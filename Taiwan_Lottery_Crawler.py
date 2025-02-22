import os
import requests
import pandas as pd
import logging
from datetime import datetime, timedelta

class TaiwanLotteryCrawler:
    # 標示當沒有數據時的提示信息
    NO_DATA = '查無資料'
    # 台灣彩票 API 的基礎 URL
    BASE_URL = 'https://api.taiwanlottery.com/TLCAPIWeB/Lottery'

    # 設定網址所使用的標頭檔，避免反爬蟲機制
    def get_lottery_result(self, url, retries=3, timeout=30):
        # 設定了 headers，其中包含 User-Agent，
        headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/58.0.3029.110 Safari/537.3'
        }
        # 發送 HTTP GET 請求來獲取數據。
        # 為了模擬正常的瀏覽器行為，避免被反爬蟲機制攔截。
        try:
            response = requests.get(url, headers=headers, timeout=timeout)
            response.raise_for_status()
            return response.json()
        # 如果請求失敗，代碼會捕獲異常
        except requests.exceptions.RequestException as e:
            if retries > 0:
                logging.warning(f"Request failed: {e}. Retrying ({retries} retries left)...")
                return self.get_lottery_result(url, retries=retries-1, timeout=timeout)
            else:
                logging.error(f"Request failed after retries: {e}")
                raise

    def bingo_bingo(self, start_date, end_date):
        # 初始化一個空的列表 datas 用來保存抓取到的數據。
        datas = []
        # 用於日期遞增
        delta = timedelta(days=1)
        current_date = start_date
        # 使用 while 迴圈遍歷從 start_date 到 end_date 的每一天。
        while current_date <= end_date:
            # 用於設定 API 請求的日期參數。
            open_date = current_date.strftime('%Y-%m-%d')
            page = 1
            # 用來抓取該日期的所有頁面數據。
            while True:
                # API網址URL
                URL = f"{self.BASE_URL}/BingoResult?openDate={open_date}&pageNum={page}&pageSize=10"

                title = '賓果賓果_' + open_date
                try:
                    # 用設定好的URL給get_lottery_result的function
                    result = self.get_lottery_result(URL)
                    print(f"Date: {open_date}, Page: {page}, Data: {result}")

                    # 如果當前頁面沒有數據，跳出循環
                    if 'content' in result and 'bingoQueryResult' in result['content']:
                        bingo_bingo_result = result['content']['bingoQueryResult']
                        if not bingo_bingo_result:
                            break
                        for record in bingo_bingo_result:
                            # 表單內的欄位代號
                            datas.append({
                                "期別": record['drawTerm'],
                                "開獎日期": record['dDate'],
                                "號碼": record['openShowOrder'],
                                "超級獎號": record['bullEye'],
                                "猜大小": record['highLow'],
                                "猜單雙": record['oddEven']
                            })
                    else:
                        # 如果沒有'bingoQueryResult'，跳出循環
                        logging.warning(f"No data found for {title}")
                        break
                except Exception as e:
                    # 在發生錯誤時跳出循環
                    logging.warning(f"Failed to fetch data for {title}: {e}")
                    break

                # 增加頁數，繼續抓取下一頁數據
                page += 1

            # 增加一天
            current_date += delta
        # 在指定日期範圍內沒有抓取到數據，會記錄警告。
        if len(datas) == 0:
            logging.warning(self.NO_DATA + str(start_date.year))

        return datas

# 使用範例
lottery = TaiwanLotteryCrawler()
start_date = datetime(2024, 1, 1)
end_date = datetime(2024, 8, 1)
# 調用 bingo_bingo 方法抓取數據並將結果存儲在 result 中。
result = lottery.bingo_bingo(start_date, end_date)

# 將結果保存為Excel文件
# 獲取當前使用者桌面的路徑
desktop_path = os.path.join(os.path.expanduser("~"), 'Desktop')
# 將數據保存到 Excel 文件中，並指定不保存索引。
excel_file_path = os.path.join(desktop_path, 'bingo_bingo_2024_01_to_08.xlsx')

# 錯誤處理
try:
    df = pd.DataFrame(result)
    df.to_excel(excel_file_path, index=False)
    print(f'數據已保存到 {excel_file_path}')
except PermissionError:
    print(f"無法寫入文件，請確保文件沒有被其他程序打開並且有適當的權限：{excel_file_path}")
except Exception as e:
    print(f"保存文件時出錯：{e}")
