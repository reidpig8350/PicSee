import requests
import pandas as pd
import config

# 從 Excel 檔案讀取資料
excel_file = "data.xlsx"  # 替換為你的 Excel 檔案名稱
sheet_name = "Sheet1"  # 替換為你的工作表名稱
data = pd.read_excel(excel_file, sheet_name=sheet_name)

# 假設 Excel 中有以下欄位：page_code, utm_source, utm_medium, utm_campaign, utm_term, utm_content
for index, row in data.iterrows():
    page_code = row['page_code']
    utm_source = row['utm_source']
    utm_medium = row['utm_medium']
    utm_campaign = row['utm_campaign']
    utm_term = row['utm_term']
    utm_content = row['utm_content']

    shorten_url = f"https://www.cathaybk.com.tw/cathaybk/promo/event/ebanking/product/appdownload/index.html?action=page&content={page_code}"
    url_scheme = f"mymobibank://?action=page&content={page_code}"

    utm = {
        "utm_source": utm_source,
        "utm_medium": utm_medium,
        "utm_campaign": utm_campaign,
        "utm_term": utm_term,
        "utm_content": utm_content
    }

    url = "https://api.pics.ee/v1/links"
    params = {"access_token": config.PICSEE_ACCESS_TOKEN}
    payload = {
        "url": shorten_url,
        "target": [
            {"target": "ios_android", "url": url_scheme}
        ],
        "utm": {
        "utm_source": utm_source,
        "utm_medium": utm_medium,
        "utm_campaign": utm_campaign,
        "utm_term": utm_term,
        "utm_content": utm_content
        },
        "gtag": "GTM-M27PT9"
    }

    try:
        response = requests.post(url, params=params, data=payload)
        response.raise_for_status()  # 如果回應狀態碼不是 200 OK，會拋出 HTTPError 異常
        print(f"第 {index + 1} 列 API 呼叫成功：", response.json())  # 印出成功的回應
    except requests.exceptions.RequestException as e:
        print(f"第 {index + 1} 列 API 呼叫發生錯誤：{e}")
        if response is not None:
            print(f"回應狀態碼：{response.status_code}")
            print(f"回應內容：{response.text}")