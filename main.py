import requests
import json
import pandas as pd

def create_pics_ee_link(page_code_prefix, access_token='YOUR_ACCESS_TOKEN', proxy=None):
    """
    使用 Pics.ee API 創建短連結。
    """
    url = 'https://api.pics.ee/v1/links'
    params = {'access_token': access_token}
    headers = {'Content-Type': 'application/json'}
    long_url = f"https://www.cathaybk.com.tw/cathaybk/promo/event/ebanking/product/appdownload/index.html?action=page&content={page_code_prefix}"
    deep_link = f"mymobibank://?action=page&content={page_code_prefix}"
    data = {
        "url": long_url,
        "domain": "cube.cathaybk.com.tw",
        "targets": [
            {"target": "ios", "url": deep_link},
            {"target": "android", "url": deep_link}
        ],
        "gTag": "GTM-W32HBPPR"
    }
    proxies = {}
    if proxy:
        proxies = {'http': f'http://{proxy}', 'https': f'https://{proxy}'}

    try:
        response = requests.post(url, headers=headers, params=params, data=json.dumps(data), proxies=proxies)
        response.raise_for_status()
        response_json = response.json()
        if 'data' in response_json and 'picseeUrl' in response_json['data']:
            return response_json['data']['picseeUrl']
        else:
            print(f"警告：API 回應中缺少 'data' 或 'picseeUrl'。代碼前綴：{page_code_prefix}")
            return None
    except requests.exceptions.RequestException as e:
        print(f"為代碼前綴 '{page_code_prefix}' 創建短連結時發生錯誤：{e}")
        return None
    except json.JSONDecodeError:
        print(f"警告：無法解析 API 回應為 JSON。代碼前綴：{page_code_prefix}")
        return None


def process_csv_and_create_links(csv_file_path, access_token='YOUR_ACCESS_TOKEN', proxy=None):
    """
    從 CSV 檔案讀取 "CUBE App Page Code" 欄位，擷取前六個字元，
    使用 Pics.ee API 創建短連結，並將結果寫回 CSV 的 "H" 欄位。
    如果 "H" 欄位已經有值，則跳過該列的 API 請求。
    """
    try:
        # 讀取 CSV 檔案
        df = pd.read_csv(csv_file_path)

        if "CUBE App Page Code" in df.columns:
            # 如果 "H" 欄位不存在，則新增該欄位
            if "H" not in df.columns:
                df["H"] = None

            short_url_list = []
            for index, row in df.iterrows():
                page_code = row["CUBE App Page Code"]
                existing_short_url = row["H"]

                # 如果 "H" 欄位已有值，跳過該列
                if pd.notna(existing_short_url):
                    print(f"第 {index + 2} 行已存在短連結，跳過 API 請求。")
                    short_url_list.append(existing_short_url)
                    continue

                if isinstance(page_code, str) and len(page_code) >= 6:
                    prefix = page_code[:6]
                    short_url = create_pics_ee_link(prefix, access_token, proxy)
                    short_url_list.append(short_url)
                else:
                    print(f"'{page_code}' 無法擷取前六個字元或不是字串 (第 {index + 2} 行)。")
                    short_url_list.append(None)

            # 更新 DataFrame 的 "H" 欄位
            df["H"] = short_url_list

            # 儲存回同一個 CSV 檔案
            df.to_csv(csv_file_path, index=False)
            print(f"短連結已成功寫回 CSV 檔案 '{csv_file_path}' 的 'H' 欄位。")

        else:
            print(f"錯誤：在 CSV 檔案 '{csv_file_path}' 中找不到 'CUBE App Page Code' 欄位。")
    except FileNotFoundError:
        print(f"錯誤：找不到 CSV 檔案 '{csv_file_path}'。")
    except Exception as e:
        print(f"讀取/寫入 CSV 檔案時發生錯誤：{e}")


if __name__ == "__main__":
    csv_file = 'deeplink.csv'  # 替換為你的 CSV 檔案名稱
    access_token = '84a1a4dd917ad11f514bdde7165ed4ae775a52b9'  # 替換為你的 Pics.ee API 存取權杖
    process_csv_and_create_links(csv_file, access_token)
