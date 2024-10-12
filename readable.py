"""
********************************************************************************
*                                                                              *
*                      README - Automated PDF Processing Script               *
*                                                                              *
********************************************************************************

概要:
--------
このスクリプトは、指定されたフォルダ内のPDFファイルを自動で処理し、指定された条件に従って分割や結合を行います。
Selenium WebDriverを使用して、PDFファイルをオンライン翻訳サービス「readable.jp」にアップロードし、翻訳されたPDFをダウンロードします。

機能:
--------
1. PDFファイルのリストアップと分割:
   - カレントディレクトリ内のPDFファイルをリストアップし、ページ数またはファイルサイズが指定の条件を超える場合に、PDFを分割します。
   - 分割されたPDFは、新しいファイルとして保存されます。

2. Chromeプロセスの管理:
   - 実行中のChromeタスクがあれば、自動的に終了させます。
   - Chromeのユーザープロファイルを指定して、Selenium WebDriverを起動します。

3. PDFのアップロードとダウンロード:
   - PDFファイルを並行してオンラインサービス「readable.jp」にアップロードします。
   - 翻訳されたPDFファイルをダウンロードし、カレントディレクトリに保存します。

4. PDFの結合:
   - "_part"が含まれるPDFファイルをグループ化し、元の名前に基づいて結合します。
   - 結合後、元の部分ファイルは削除されます。

5. 古いファイルの管理:
   - ダウンロード後5時間以内に取得したファイルをカレントディレクトリに移動します。
   - 不要な"_part"ファイルも削除します。

使用方法:
-----------
1. Chromeのユーザープロファイル:
   - ユーザープロファイルは、`C:/Users/butsu/AppData/Local/Google/Chrome/User Data/Profile 11`に設定されています。
   - 別の環境で使用する場合は、コード内のChromeオプションを適宜修正してください。

2. PDFファイルの処理:
   - スクリプトは、カレントディレクトリにあるすべてのPDFファイルを対象とします。
   - 分割するPDFは、ページ数が100ページを超えるか、ファイルサイズが50MBを超える場合に処理されます。

3. 実行手順:
   - スクリプトを実行することで、PDFファイルの分割、アップロード、ダウンロード、結合が自動的に行われます。

注意点:
--------
- このスクリプトはSeleniumとChromeDriverを使用しますので、事前にChromeDriverのパスを設定し、Seleniumをインストールしてください。
- タイムアウトや接続エラーが発生する場合は、ネットワーク環境やSeleniumの設定を確認してください。

ライセンス:
-----------
このスクリプトは自由に使用、変更、再配布可能です。商用利用も許可されますが、自己責任でご利用ください。

********************************************************************************
"""

import os
import time
import shutil
import subprocess
from datetime import datetime, timedelta
from PyPDF2 import PdfReader, PdfWriter
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from collections import defaultdict

# カレントディレクトリのパス
current_directory = os.getcwd()

# PDFファイルをリストアップ
pdf_files = [f for f in os.listdir(current_directory) if f.endswith('.pdf')]

# Chromeのオプションを設定
chrome_options = Options()
chrome_options.add_argument("--start-maximized")
chrome_options.add_argument(f"user-data-dir=C:/Users/butsu/AppData/Local/Google/Chrome/User Data")
chrome_options.add_argument("profile-directory=Profile 11")

# Chromeのタスクが実行中か確認し、実行中の場合のみタスクを終了
def kill_chrome_if_running():
    try:
        tasklist_output = subprocess.run(["tasklist"], capture_output=True, text=True)
        if "chrome.exe" in tasklist_output.stdout:
            subprocess.run(["taskkill", "/F", "/IM", "chrome.exe"], check=True)
            print("Chromeタスクを終了しました。")
        else:
            print("Chromeは実行されていません。")
    except Exception as e:
        print(f"Chromeタスク終了中にエラーが発生しました: {e}")

# ダウンロードが完了したか確認する関数
def wait_for_download(download_path, timeout=300):
    end_time = time.time() + timeout
    while time.time() < end_time:
        if os.path.exists(download_path):
            print(f"ファイルがダウンロードされました: {download_path}")
            return True
        time.sleep(1)
    return False

# PDFを分割する関数
def split_pdf(file_path, max_pages=100, max_size_mb=50):
    file_size = os.path.getsize(file_path) / (1024 * 1024)  # サイズをMBに変換
    reader = PdfReader(file_path)
    total_pages = len(reader.pages)

    if total_pages <= max_pages and file_size <= max_size_mb:
        print(f"{file_path} はページ数やサイズが条件に適しています。")
        return [file_path]

    base_name = os.path.splitext(os.path.basename(file_path))[0]
    output_files = []
    part_number = 1

    for start_page in range(0, total_pages, max_pages):
        writer = PdfWriter()
        current_output_file = f"{base_name}_part{part_number}.pdf"
        output_files.append(current_output_file)
        
        for page_num in range(start_page, min(start_page + max_pages, total_pages)):
            writer.add_page(reader.pages[page_num])

        with open(current_output_file, "wb") as output_pdf:
            writer.write(output_pdf)

        part_number += 1

    print(f"{file_path} が分割されました: {output_files}")
    return output_files

# Chromeを終了させる（実行中であれば）
kill_chrome_if_running()

# 1. 最初に全てのPDFファイルを分割する
all_pdf_files = []
for pdf_file in pdf_files:
    # 'al-' で始まるかどうかを確認
    if pdf_file.startswith("al-"):
        # 'al-' を除去したファイル名を取得
        stripped_name = pdf_file.replace("al-", "")
        
        # 同じファイル名（'al-' を除去した名前）がカレントディレクトリに存在するか確認
        if stripped_name in os.listdir(current_directory):
            print(f"{pdf_file} は同じ名前のファイルがあるためスキップされました。")
            continue

    pdf_paths = split_pdf(os.path.join(current_directory, pdf_file))
    for pdf_path in pdf_paths:
        absolute_pdf_path = os.path.abspath(pdf_path)
        if absolute_pdf_path not in all_pdf_files:
            all_pdf_files.append(absolute_pdf_path)

print("変換するものを次に示す:")
for pdf_file in all_pdf_files:
    print(pdf_file)

# Selenium WebDriverの処理を開始
try:
    driver = webdriver.Chrome(options=chrome_options)
    
    # 2. 各PDFファイルを並行してアップロードするためにタブを開く
    for i, pdf_path in enumerate(all_pdf_files):
        if i == 0:
            driver.get("https://readable.jp/translate")  # 最初のタブでページを開く
        else:
            driver.execute_script("window.open('');")  # 新しいタブを開く
            driver.switch_to.window(driver.window_handles[i])  # 新しいタブに切り替え
            driver.get("https://readable.jp/translate")

        time.sleep(2)  # ページが完全に読み込まれるまで待機

        # PDFの絶対パスを取得
        absolute_pdf_path = os.path.abspath(pdf_path)
        print(f'{pdf_path}をアップ中です.')
        
        upload_button = WebDriverWait(driver, 30).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, 'input[type="file"]'))
        )
        upload_button.send_keys(absolute_pdf_path)

    # 3. 各タブでダウンロードリンクが現れるまで待機し、ダウンロードを開始
    for i, pdf_path in enumerate(all_pdf_files):
        driver.switch_to.window(driver.window_handles[i])
        try:
            download_link = WebDriverWait(driver, 300).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, 'a[href*="https://files.readable.jp/"]'))
            )
            download_link.click()

            download_path = os.path.join(current_directory, f"downloaded_{os.path.basename(pdf_path)}.pdf")

            if wait_for_download(download_path, timeout=1):
                print(f"ダウンロード完了: {download_path}")
            else:
                print(f"ダウンロードがタイムアウトしました: {pdf_path}")

        except Exception as e:
            print(f"ダウンロードリンクを待機中にエラーが発生しました: {e}")

except Exception as e:
    print(f"Selenium WebDriver実行中にエラーが発生しました: {e}")

finally:
    time.sleep(30)  # ダウンロードが終わるまで待機
    driver.quit()  # 終了処理でChromeを閉じる

    #==========================================================================
    # ダウンロードフォルダのパス (Windowsのデフォルトダウンロードフォルダ)
    download_directory = os.path.join(os.path.expanduser('~'), 'Downloads')

    # "_part"が含まれるPDFファイルを探してリストアップ
    part_files = [f for f in os.listdir(download_directory) if '_part' in f and f.endswith('.pdf')]

    # ファイルをベース名ごとにグループ化 (例: hoge, fuga)
    file_groups = defaultdict(list)
    for part_file in part_files:
        base_name = part_file.split('_part')[0]
        file_groups[base_name].append(part_file)

    # 各グループごとにPDFを結合
    for base_name, files in file_groups.items():
        # ファイルを数字部分に基づいてソート
        files = sorted(files, key=lambda x: int(''.join(filter(str.isdigit, x))))

        # 結合処理の開始
        print(f"結合対象のファイル: {files}")

        writer = PdfWriter()
        
        for part_file in files:
            part_path = os.path.join(download_directory, part_file)
            reader = PdfReader(part_path)
            for page in reader.pages:
                writer.add_page(page)
        
        # 結合後のPDFファイルの保存先
        output_pdf_path = os.path.join(download_directory, f'{base_name}.pdf')
        
        # 結合PDFを保存
        with open(output_pdf_path, 'wb') as output_pdf:
            writer.write(output_pdf)
        
        print(f"PDFファイルが結合されました: {output_pdf_path}")

        # 元の part ファイルを削除
        for part_file in files:
            part_path = os.path.join(download_directory, part_file)
            os.remove(part_path)
            print(f"削除されました: {part_path}")

    # 5時間以内にダウンロードされたPDFファイルをカレントディレクトリに移動
    now = datetime.now()
    five_hours_ago = now - timedelta(hours=5)

    for file in os.listdir(download_directory):
        if file.endswith('.pdf'):
            file_path = os.path.join(download_directory, file)
            file_mod_time = datetime.fromtimestamp(os.path.getmtime(file_path))

            if file_mod_time > five_hours_ago:
                shutil.move(file_path, current_directory)
                print(f"ファイルを移動しました: {file}")
    
    # ディレクトリ内のファイルを確認して、"_part" が含まれる PDF ファイルを削除
    for file in os.listdir(current_directory):
        if '_part' in file and file.endswith('.pdf'):
            file_path = os.path.join(current_directory, file)
            if os.path.exists(file_path):
                os.remove(file_path)
                print(f"{file} を削除しました。")
            else:
                print(f"{file} が見つかりませんでした。")

    
