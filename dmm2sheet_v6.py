import requests
from bs4 import BeautifulSoup
import pygsheets
from google.oauth2.service_account import Credentials
import urllib.parse

# 設定 Google Sheets API 認證
SERVICE_ACCOUNT_FILE = 'file path'  # 憑證文件路徑

credentials = Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE)
client = pygsheets.authorize(service_account_file=SERVICE_ACCOUNT_FILE)
spreadsheet_id = 'target sheet ID'
sh = client.open_by_key(spreadsheet_id)

# 年齡認證
def age_verification():
    session = requests.Session()
    age_check_url = "https://www.dmm.co.jp/age_check/=/declared=yes/?rurl=https%3A%2F%2Fwww.dmm.co.jp%2Fdigital%2Fvideoa%2F-%2Flist%2F%3Factress%3D8704%26view%3Dtext"
    session.get(age_check_url)
    return session

# 搜尋出演者ID
def search_actress_id(session, actress_name):
    search_url = f'https://actress.dmm.co.jp/-/search/=/searchstr={urllib.parse.quote(actress_name)}'
    response = session.get(search_url)
    response.encoding = 'utf-8'
    soup = BeautifulSoup(response.text, 'html.parser')
    
    actress_tag = soup.find('a', href=lambda x: x and '/-/detail/=/actress_id=' in x)
    if not actress_tag:
        return None
    actress_url = actress_tag['href']
    actress_id = actress_url.split('actress_id=')[-1].split('/')[0]
    return actress_id

# 獲取影片資料
def fetch_video_data(session, actress_id):
    video_data = []
    page = 1
    while True:
        list_url = f'https://www.dmm.co.jp/digital/videoa/-/list/?actress={actress_id}&view=text&page={page}'
        response = session.get(list_url)
        response.encoding = 'utf-8'
        soup = BeautifulSoup(response.text, 'html.parser')

        # 檢查是否有影片
        video_links = soup.find_all('a', href=lambda x: x and '/digital/videoa/-/detail/=/cid=' in x)
        if not video_links or response.url != list_url:
            break
        for link in video_links:
            video_url = link['href']
            if 'cid=' in video_url:
                video_page_url = f'https://www.dmm.co.jp{video_url}'
                video_page = session.get(video_page_url)
                video_soup = BeautifulSoup(video_page.text, 'html.parser')

                try:
                    title = video_soup.find('meta', property='og:title')['content']
                    code_td = video_soup.find('td',  string='品番：')
                    if code_td:
                        code = code_td.find_next_sibling('td').text.strip()
                    else:
                        code = 'Unknown'
                    release_date = video_soup.find('td', string='配信開始日：').find_next_sibling('td').text.strip()
                    sale_date = video_soup.find('td', string='商品発売日：').find_next_sibling('td').text.strip()
                    genres = video_soup.find('td', string='ジャンル：').find_next_sibling('td').find_all('a')
                    genres_text = [genre.text.strip() for genre in genres]
                    single_work = "単体作品" if "単体作品" in genres_text else ""
                    best_of = "ベスト・総集編" if "ベスト・総集編" in genres_text else ""

                    video_data.append({
                        'code': code,
                        'title': title,
                        'video_page_url': video_page_url,
                        'release_date': release_date,
                        'sale_date': sale_date,
                        'genres': genres_text,
                        'single_work': single_work,
                        'best_of': best_of
                    })
                except Exception as e:
                    print(f"錯誤處理影片資料：{e}")
                    continue

        page += 1

    return video_data

# 寫入 Google Sheets
def write_to_google_sheets(actress_name, video_data):
    # 檢查是否已存在該出演者的分頁
    worksheet = sh.worksheet_by_title(actress_name)
    if worksheet:
        worksheet = sh.worksheet_by_title(actress_name)  
    else :
        worksheet = sh.add_worksheet(actress_name, rows=5000, cols=100)
        worksheet.update_value('A1', '下載者')
        worksheet.update_value('B1', '通用番號')
        worksheet.update_value('C1', '品番')
        worksheet.update_value('D1', '片名')
        worksheet.update_value('E1', '配信開始日')
        worksheet.update_value('F1', '商品發售日')
        worksheet.update_value('G1', '単体作品')
        worksheet.update_value('H1', 'ベスト・総集編')

    existing_codes = worksheet.get_col(3, include_tailing_empty=False)[1:]  # 獲取已存在的品番

    # 寫入新資料
    row_index = len(existing_codes) + 2
    for data in video_data:
        if data['code'] not in existing_codes:
            worksheet.update_value(f'C{row_index}', data['code'])
            worksheet.update_value(f'D{row_index}', f'=HYPERLINK("{data['video_page_url']}", "{data["title"]}")')
            worksheet.update_value(f'E{row_index}', data['release_date'])
            worksheet.update_value(f'F{row_index}', data['sale_date'])
            worksheet.update_value(f'G{row_index}', data['single_work'])
            worksheet.update_value(f'H{row_index}', data['best_of'])
            
            # 將其他的ジャンル資料寫入G欄之後的欄位
            col_index = 9  # I欄開始
            for genre in data['genres']:
                if genre not in ['単体作品', 'ベスト・総集編']:
                    worksheet.update_value((row_index, col_index), genre)
                    col_index += 1
            
            row_index += 1

# 主程式
def main():
    # 建立會話，執行年齡認證
    session = age_verification()

    # 讀取"女優列表"分頁
    actress_list_sheet = sh.worksheet_by_title("女優列表")
    actress_names = actress_list_sheet.get_col(1, include_tailing_empty=False)[1:]  # 跳過標題列，從A2開始
    
    # 取得 D 欄的所有儲存格值
    #d_col_values = actress_list_sheet.col_values(4)[1:]  # 跳過標題列，從 D2 開始

    for i, actress_name in enumerate(actress_names, start=2):
        actress_id_value = actress_list_sheet.cell(f'D{i}').value
        if not actress_id_value:
            actress_id = search_actress_id(session, actress_name)
        else:
            actress_id = actress_id_value

        if actress_id:
            # 更新女優列表中的出演者ID
            actress_list_sheet.update_value(f'D{i}', f'=HYPERLINK("https://www.dmm.co.jp/digital/videoa/-/list/?actress={actress_id}", "{actress_id}")')

            # 獲取影片資料
            video_data = fetch_video_data(session, actress_id)

            # 寫入資料到Google Sheets
            write_to_google_sheets(actress_name, video_data)

            # 將分頁女優列表中的出演者名稱建立連結到對應的分頁
            actress_name_tab = sh.worksheet_by_title(actress_name)
            actress_name_tab_gid = actress_name_tab.id
            actress_list_sheet.update_value(f'A{i}', f'=HYPERLINK("#gid={actress_name_tab_gid}", "{actress_name}")')
        else:
            print(f"查無此人：{actress_name}")
            continue
        
    # 取得所有分頁
    worksheets = sh.worksheets()

    for ws in worksheets:
    # 凍結第一行
        ws.frozen_rows = 1
        
if __name__ == "__main__":
    main()
