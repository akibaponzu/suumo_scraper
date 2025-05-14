# 必要なライブラリのインポート
import requests  # HTTPリクエストを送るため
from bs4 import BeautifulSoup  # HTMLを解析するため
import openpyxl  # Excelファイル作成
from openpyxl.styles import Font  # Excelセルのフォント装飾
import time  # 待機処理などに使用
import schedule  # 定期実行スケジューラ
import logging  # ログ出力
import sys  # 標準出力制御

# ログ設定：コンソール＆ファイル出力の両方
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.StreamHandler(sys.stdout),
        logging.FileHandler("scraping.log", encoding="utf-8")
    ]
)
logger = logging.getLogger(__name__)

# ====== 定数定義 ======
BASE_URL = "https://suumo.jp"  # 詳細ページURLのベース
SEARCH_URL = "https://suumo.jp/jj/chintai/ichiran/FR301FC001/?..."  # 検索結果ページURL（長いので省略）
MAX_RETRY = 3  # 最大リトライ回数
RETRY_INTERVAL = 5  # リトライ間隔（秒）
OUTPUT_EXCEL_FILE = "賃貸物件情報.xlsx"  # 出力ファイル名
HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"  # ブラウザのふりをする
}

# ====== ページ取得処理（GET + リトライ機能付き） ======
def fetch_page(url):
    for i in range(MAX_RETRY):
        try:
            response = requests.get(url, headers=HEADERS, timeout=10)
            response.raise_for_status()
            return response.text
        except requests.exceptions.RequestException as e:
            logger.error(f"Error fetching {url}: {e}")
            if i < MAX_RETRY - 1:
                logger.info(f"Retrying after {RETRY_INTERVAL} seconds...")
                time.sleep(RETRY_INTERVAL)
    return None

# ====== 検索結果ページの解析（物件リスト） ======
def parse_search_results(html):
    soup = BeautifulSoup(html, "html.parser")
    properties = soup.find_all("div", class_="property_unit-content")  # 各物件ブロック
    data = []

    for property_unit in properties:
        try:
            name = property_unit.find("div", class_="property_unit-title").text.strip()  # 物件名
            price = property_unit.find("span", class_="price").text.strip()  # 賃料
            madori = property_unit.find("span", class_="madori").text.strip()  # 間取り

            # 駅の情報は2番目の <div> に含まれている
            stations_info = property_unit.find("div", class_="property_unit-body").find_all("div")[1].text.strip()
            stations = [s.strip() for s in stations_info.split("、")]  # 複数駅対応

            # 詳細ページURLを取得（相対URL → 絶対URLに変換）
            url = BASE_URL + property_unit.find("a", class_="js-物件概要")["href"]

            data.append({
                "物件名": name,
                "賃料": price,
                "間取り": madori,
                "最寄駅": stations,
                "URL": url
            })
        except AttributeError as e:
            # 万が一構造が異なる物件があればスキップ
            logger.warning(f"Failed to extract property: {e}")
    return data

# ====== 詳細ページから「構造」や「敷金/礼金」などを取得 ======
def parse_property_details(html):
    soup = BeautifulSoup(html, "html.parser")
    details = {"構造": "情報なし", "敷金/礼金": "情報なし"}  # 初期値

    try:
        rows = soup.select("table.data_table tr")  # 詳細情報テーブルを取得
        for row in rows:
            header = row.find("th")
            value = row.find("td")
            if not header or not value:
                continue
            key = header.text.strip()

            # 構造・敷金/礼金をチェックして辞書に格納
            if "構造" in key:
                details["構造"] = value.text.strip()
            elif "敷金" in key:
                details["敷金/礼金"] = value.text.strip()
    except Exception as e:
        logger.warning(f"Failed to parse property details: {e}")
    return details

# ====== Excelファイルへの出力処理 ======
def save_to_excel(data, filename=OUTPUT_EXCEL_FILE):
    if not data:
        logger.warning("No data to save.")
        return

    # 新しいExcelファイルを作成
    workbook = openpyxl.Workbook()
    sheet = workbook.active

    # ヘッダー行を追加
    header = ["物件名", "賃料", "敷金/礼金", "構造", "間取り", "最寄駅", "URL"]
    sheet.append(header)
    for cell in sheet[1]:
        cell.font = Font(bold=True)

    # データを1件ずつ書き込み
    for item in data:
        stations = ", ".join(item["最寄駅"])
        row = [
            item["物件名"],
            item["賃料"],
            item["敷金/礼金"],
            item["構造"],
            item["間取り"],
            stations,
            item["URL"]
        ]
        sheet.append(row)

    try:
        workbook.save(filename)
        logger.info(f"Saved data to {filename}")
    except Exception as e:
        logger.error(f"Error saving Excel: {e}")

# ====== スクレイピング本体処理 ======
def scrape_and_save():
    logger.info("Start scraping...")

    html = fetch_page(SEARCH_URL)
    if html is None:
        logger.error("Failed to get search results.")
        return

    # 検索結果から物件リスト取得
    property_list = parse_search_results(html)
    if not property_list:
        logger.warning("No properties found.")
        return

    # 各物件ごとに詳細ページへアクセス
    all_data = []
    for property_info in property_list:
        detail_html = fetch_page(property_info["URL"])
        if detail_html is None:
            logger.warning(f"Failed to fetch detail: {property_info['URL']}")
            continue

        # 詳細情報を取得してマージ
        details = parse_property_details(detail_html)
        property_info.update(details)
        all_data.append(property_info)

    # 条件フィルタ（例：「木造」を除外）
    filtered_data = [p for p in all_data if "木造" not in p["構造"]]

    # Excelへ保存
    save_to_excel(filtered_data)
    logger.info("Scraping done.")

# ====== メイン処理：毎日9:00に自動実行 ======
def main():
    schedule.every().day.at("09:00").do(scrape_and_save)
    logger.info("Scheduled at 09:00 daily.")

    while True:
        schedule.run_pending()
        time.sleep(60)  # 毎分チェック

# ====== エントリーポイント（スクリプトが直接実行されたときのみ起動） ======
if __name__ == "__main__":
    main()
