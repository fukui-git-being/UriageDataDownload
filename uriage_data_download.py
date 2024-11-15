"""
UriageDataDownloadモジュールは売上データをダウンロードするための機能を提供します。
"""
import os
import time
import sys
import datetime
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from dotenv import load_dotenv

load_dotenv()  # これにより .env ファイルから環境変数が読み込まれる


def main():
    """
    main
    """
    debug_mode = '--debug' in sys.argv

    # SeleniumのWebDriverを使ってブラウザを操作
    options = webdriver.ChromeOptions()
    if not debug_mode:
        options.add_argument('--headless')  # ブラウザを非表示で実行するオプション（デバッグ時は無効）
    options.add_argument('--ignore-certificate-errors')
    options.add_argument('--ignore-ssl-errors')
    options.add_argument('--no-sandbox')
    options.add_argument('--disable-dev-shm-usage')
    options.add_argument('--disable-gpu')
    options.add_argument('--disable-web-security')
    options.add_argument('--allow-running-insecure-content')    

    # 新しい設定を追加
    options.add_argument('--safebrowsing-disable-download-protection')
    options.add_argument('--safebrowsing-disable-extension-blacklist')
    options.add_argument('--disable-features=IsolateOrigins,site-per-process')

    # 特定のドメインを安全として扱う（ダウンロード元のドメインに置き換えてください）
    url = "http://sss-web/OrderEntry/main/login.asp"
    options.add_argument(f'--unsafely-treat-insecure-origin-as-secure={url}')

    # ダウンロード設定
    download_dir = "C:\\Users\\fukui\\株式会社ビーイング\\Gaia企画部 - ドキュメント\\資料\\売上データ\\売上集計\\Input\\Uriage"  # ダウンロードディレクトリを指定
    options.add_experimental_option("prefs", {
        "download.default_directory": download_dir,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": False
    })

    driver = webdriver.Chrome(options=options)

    try:
        # 与えられたURLを開く
        driver.get(url)

        # ページが読み込まれるまで待機
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.NAME, "DeptID")))

        # 部署を選択する
        department_select = Select(driver.find_element(By.NAME, "DeptID"))
        department_select.select_by_visible_text("Gaia企画部")

        # 名称をコンボボックスで選択
        name_select = Select(driver.find_element(By.NAME, "UserID"))
        name_select.select_by_visible_text("福井 洋行")

        # パスワードを入力
        password_input = driver.find_element(By.NAME, "txtpswd")
        password_input.send_keys(os.getenv("PASSWORD_"))

        # ログインボタンをクリック
        login_button = driver.find_element(By.NAME, "login")
        login_button.click()

        # メニュー画面が開くのを待つ
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CLASS_NAME, "MenuButton")))
        # 情報ボタンをクリック
        info_button = driver.find_element(By.XPATH, "//div[@class='MenuButton' and text()='情報']")
        info_button.click()

        # 情報閲覧ボタンをクリック
        info_button = driver.find_element(By.XPATH, "//div[@class='MainMenuBtn' and text()='情報閲覧']")
        info_button.click()

        # 情報閲覧メニューをクリックして資料室のページに遷移
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//div[@class='ProductBtn' and contains(text(), '【資料】営業所別明細一覧')]")))
        download_button = driver.find_element(By.XPATH, "//div[@class='ProductBtn' and contains(text(), '【資料】営業所別明細一覧')]")
        download_button.click()

        # ファイルがダウンロードされるまで少し待機（必要に応じて時間を調整） ダウンロードされても、SharePoint上なので待たないとうまくリネームできない
        time.sleep(15)
        
        today_str = datetime.datetime.now().strftime('%Y%m%d_%H%M%S')
        # ダウンロードされたファイルを特定し、名前を変更
        for file_name in os.listdir(download_dir):
            if file_name.startswith('＜情報閲覧用＞') and '営業所別明細一覧' in file_name:
                old_file_path = os.path.join(download_dir, file_name)
                new_file_name = f"{today_str}_{file_name}"
                new_file_path = os.path.join(download_dir, new_file_name)
                os.rename(old_file_path, new_file_path)
                print(f"ファイル名を変更しました: {new_file_name}")
                break
    except TimeoutException:
        print("処理がタイムアウトしました。要素が見つからなかった可能性があります。")
    finally:
        # ブラウザを閉じる
        driver.quit()


if __name__ == "__main__":
    main()
