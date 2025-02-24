import os
import time
import pandas as pd
import pyautogui
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager

TARGET_USER = input("Lütfen Instagram profil URL'sini giriniz: ")

# **Kullanıcı adını al ve Chrome profili yolunu oluştur**
user_name = os.getlogin()
chrome_profile_path = os.path.join(os.getenv("LOCALAPPDATA"), "Google", "Chrome", "User Data")

# **WebDriver başlat**
options = webdriver.ChromeOptions()
options.add_argument("--start-maximized")
options.add_argument("--profile-directory=Default")
options.add_argument(f"user-data-dir={chrome_profile_path}")  # Dinamik Chrome kullanıcı verisi yolu

# **WebDriver servis kullanarak başlat**
service = Service(ChromeDriverManager().install())
driver = webdriver.Chrome(service=service, options=options)

def save_to_excel(followers, following):
    max_len = max(len(followers), len(following))
    followers += [""] * (max_len - len(followers))
    following += [""] * (max_len - len(following))
    df = pd.DataFrame({"Takipçiler": followers, "Takip Edilenler": following})
    df.to_excel("instagram_followers_following.xlsx", index=False)
    print("Veriler Excel dosyasına kaydedildi!")


def scroll_with_middle_click(scroll_box, users, list_type):
    """Orta tık ile takipçi listesini kaydırır ve her kaydırmada Excel'e kaydeder."""
    pyautogui.moveTo(971, 620, duration=0.5)
    pyautogui.mouseDown(button='middle')
    time.sleep(0.2)

    last_count = 0  # Önceki takipçi sayısı
    same_count_times = 0  # Değişmeyen liste sayacı
    max_same_count = 2  # Kaç kez değişmezse durmalı?

    while same_count_times < max_same_count:
        new_users = [elem.text for elem in scroll_box.find_elements(By.TAG_NAME, "a") if elem.text != ""]

        if len(new_users) == last_count:
            same_count_times += 1
        else:
            same_count_times = 0  # Değişim olduysa sıfırla

        last_count = len(new_users)  # Son değeri güncelle
        users.extend([u for u in new_users if u not in users])  # Yeni kullanıcıları ekle
        save_to_excel(users, []) if list_type == "followers" else save_to_excel([], users)  # Kaydet

        # **Fareyi aşağı kaydır**
        pyautogui.moveTo(971, 750, duration=0.5)
        time.sleep(1)  # Kaydırmanın etkisini görmek için bekle

    pyautogui.mouseUp(button='middle')
    print("Kaydırma işlemi tamamlandı!")


def get_follow_list(list_type):
    driver.get(TARGET_USER)
    time.sleep(3)

    wait = WebDriverWait(driver, 10)
    if list_type == "followers":
        button = wait.until(
            EC.element_to_be_clickable((By.XPATH, "//a[contains(@href, '/followers/') and @role='link']")))
    else:
        button = wait.until(
            EC.element_to_be_clickable((By.XPATH, "//a[contains(@href, '/following/') and @role='link']")))

    button.click()
    time.sleep(3)

    # **Popup içindeki kaydırma kutusunu al**
    scroll_box = wait.until(EC.presence_of_element_located((By.XPATH, "//div[@role='dialog']//div[2]")))
    users = []

    # **Kaydırma işlemi ve Excel'e kaydetme**
    scroll_with_middle_click(scroll_box, users, list_type)
    return users


# **Çalıştırma**
followers = get_follow_list("followers")
following = get_follow_list("following")
save_to_excel(followers, following)

driver.quit()
