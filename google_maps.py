from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import pandas as pd
from datetime import datetime

# Kullanıcıdan konum ve arama terimi al
location = input("Konum giriniz (örn: Başakşehir): ").strip()
keyword = input("Ne arıyorsunuz? (örn: pizzacı): ").strip()
search_query = f"{keyword} {location}"

print(f"'{search_query}' için Google Haritalar'da arama yapılıyor...")

# Tarayıcı başlat
# ChromeDriver'ınızın yolunu doğru bir şekilde belirttiğinizden emin olun.
service = Service("c:/Users/LENOVO/Documents/chromedriver-win64/chromedriver-win64/chromedriver.exe")
driver = webdriver.Chrome(service=service)

# Google Haritalar'ın doğru URL'sine git
driver.get("https://www.google.com/maps")

# Sayfanın yüklenmesini bekle
time.sleep(3) # İlk yükleme için kısa bir bekleme

try:
    # Arama kutusunu bul ve arama yap
    search_box = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.ID, "searchboxinput"))
    )
    search_box.send_keys(search_query)
    search_box.send_keys(Keys.ENTER)

    # Arama sonuçlarının yüklenmesini bekle
    time.sleep(5) # Sonuçların görünmesi için daha uzun bir bekleme

    # Kaydırılabilir sonuçlar panelini bul
    scrollable_div = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.XPATH, '//div[@role="feed"]'))
    )

    # Tüm işletmeleri yüklemek için aşağı kaydırma işlemi
    print("Tüm sonuçları yüklemek için kaydırılıyor...")
    last_height = driver.execute_script("return arguments[0].scrollHeight", scrollable_div)
    while True:
        driver.execute_script("arguments[0].scrollTop = arguments[0].scrollHeight", scrollable_div)
        time.sleep(2) # Kaydırma sonrası yeni içeriğin yüklenmesini bekle

        # "Daha fazla sonuç" veya benzeri bir mesajın görünmesini kontrol et
        try:
            end_of_results = driver.find_element(By.XPATH, "//div[contains(text(), 'Sonuçların sonuna ulaşıldı')]")
            if end_of_results.is_displayed():
                print("Sonuçların sonuna ulaşıldı.")
                break
        except:
            pass # Mesaj bulunamazsa devam et

        new_height = driver.execute_script("return arguments[0].scrollHeight", scrollable_div)
        if new_height == last_height:
            # Eğer kaydırma sonrası yükseklik değişmiyorsa, tüm içerik yüklenmiştir.
            break
        last_height = new_height
    print("Kaydırma tamamlandı.")
    time.sleep(2) # Kaydırma sonrası son yüklemeler için bekle

    # İşletme kartlarının bağlantılarını topla
    # Doğrudan işletme detay sayfasına giden 'a' etiketlerini bul.
    # Bu XPath, kaydırılabilir div içindeki tüm 'a' etiketlerini arar ve '/maps/place/' içeren href'e sahip olanları seçer.
    business_card_elements = driver.find_elements(By.XPATH, '//div[@role="feed"]//a[contains(@href, "/maps/place/")]')
    
    business_links = []
    if not business_card_elements:
        print("Uyarı: İşletme linkleri bulunamadı. XPath değişmiş olabilir veya sonuç yok.")
    else:
        for element in business_card_elements:
            try:
                link = element.get_attribute('href')
                if link: # href özniteliğinin boş olmadığından emin ol
                    business_links.append(link)
            except Exception as e:
                print(f"Link alınırken hata oluştu: {e}")
                continue

    print(f"Toplam {len(business_links)} işletme linki bulundu.")

    results = []
    for i, link in enumerate(business_links):
        print(f"İşletme {i+1}/{len(business_links)} detayları çekiliyor: {link}")
        driver.get(link) # İşletmenin detay sayfasına git
        time.sleep(3) # Detay sayfasının yüklenmesini bekle

        name = ""
        rating = ""
        category = ""
        address = ""
        phone = ""
        maps_link = driver.current_url # Ziyaret edilen sayfanın URL'si


        try:
            # Alternatif olarak, meta etiketinden adı almaya çalış
            meta_name_element = WebDriverWait(driver, 5).until(
                EC.presence_of_element_located((By.XPATH, '//meta[@itemprop="name"]'))
            )
            name = meta_name_element.get_attribute('content')
        except Exception as e:
            print(f"İşyeri adı meta etiketinden de bulunamadı: {e}")


        try:
            # Puan: 'yıldızlı' aria-label'ına sahip span'i bul ve metnini al.
            # Bazen puan doğrudan bu span'in metni içindedir, bazen de bir alt span'de.
            rating_element = WebDriverWait(driver, 5).until(
                EC.presence_of_element_located((By.XPATH, '//span[contains(@aria-label, "yıldızlı")]'))
            )
            # Metni doğrudan veya ilk alt span'den almaya çalış
            try:
                rating = rating_element.text.split(" ")[0]
            except:
                # Eğer doğrudan metin alınamazsa, alt span'i dene
                rating_sub_element = rating_element.find_element(By.XPATH, './span[1]')
                rating = rating_sub_element.text.split(" ")[0]
        except Exception as e:
            print(f"Puan bulunamadı: {e}")

        try:
            # Kategori: Kategori genellikle bir buton veya belirli bir sınıf yapısı içinde yer alır.
            # Birden fazla olası XPath denemesi.
            # 1. Deneme: 'Kategori:' aria-label'ına sahip buton
            try:
                category_element = WebDriverWait(driver, 3).until(
                    EC.presence_of_element_located((By.XPATH, '//button[contains(@aria-label, "Kategori:")]'))
                )
                category = category_element.text
            except:
                # 2. Deneme: 'W4Efsd' sınıfına sahip div içindeki ilk fontBodyMedium span'i (puan veya adres olmayan)
                try:
                    category_element = WebDriverWait(driver, 3).until(
                        EC.presence_of_element_located((By.XPATH, '//div[@class="W4Efsd"]//span[contains(@class, "fontBodyMedium") and not(contains(@aria-label, "yıldızlı")) and not(contains(@aria-label, "adres"))][1]'))
                    )
                    category = category_element.text
                except:
                    # 3. Deneme: 'W4Efsd' sınıfına sahip div içindeki doğrudan kategori metni
                    category_element = WebDriverWait(driver, 3).until(
                        EC.presence_of_element_located((By.XPATH, '//div[@class="W4Efsd"]//span[contains(@class, "fontBodyMedium") and not(contains(@aria-label, "yıldızlı")) and not(contains(@aria-label, "adres"))]'))
                    )
                    category = category_element.text
        except Exception as e:
            print(f"Kategori bulunamadı: {e}")


        try:
            # Adres
            # Adres için birden fazla XPath denemesi
            try:
                address_element = WebDriverWait(driver, 5).until(
                    EC.presence_of_element_located((By.XPATH, '//button[contains(@data-tooltip, "Adresi kopyala")]//div[contains(@class, "fontBodyMedium")]'))
                )
                address = address_element.text
            except:
                address_element = WebDriverWait(driver, 5).until(
                    EC.presence_of_element_located((By.XPATH, '//button[contains(@aria-label, "Adres:")]'))
                )
                address = address_element.text.replace("Adres: ", "")
        except Exception as e:
            print(f"Adres bulunamadı: {e}")

        try:
            # Telefon
            phone_element = WebDriverWait(driver, 5).until(
                EC.presence_of_element_located((By.XPATH, '//button[contains(@aria-label, "Telefon:")]'))
            )
            phone = phone_element.text.replace("Telefon: ", "")
        except Exception as e:
            print(f"Telefon bulunamadı: {e}")

        current_result = {
            "İşyeri Adı": name,
            "Puan": rating,
            "Adres": address,
            "Telefon": phone,
            "Link": maps_link,
            "Kategori": category
        }
        print(current_result)
        results.append(current_result)

        # Bir sonraki işletmeye geçmeden önce ana arama sayfasına geri dön
        driver.back()
        time.sleep(2) # Geri dönme işleminin tamamlanmasını bekle

except Exception as e:
    print(f"Genel bir hata oluştu: {e}")
    print("Lütfen ChromeDriver'ınızın güncel olduğundan ve doğru yolda olduğundan emin olun.")
    print("Ayrıca, Google Haritalar'ın arayüzü değişmiş olabilir, bu durumda XPath'lerin güncellenmesi gerekebilir.")

finally:
    # Tarayıcıyı kapat
    driver.quit()

    # Excel'e yaz
    if results:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        file_name = f"{keyword}_{location}_{timestamp}.xlsx"

        df = pd.DataFrame(results)
        df.to_excel(file_name, index=False)

        print(f"\n✅ {len(results)} sonuç '{file_name}' dosyasına kaydedildi.")
    else:
        print("\n❌ Hiç sonuç bulunamadı veya bir hata oluştuğu için veri kaydedilemedi.")
