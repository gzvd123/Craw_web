import requests
import pandas as pd 
import re
from time import sleep 
import html  # Thêm thư viện để giải mã HTML entities
from concurrent.futures import ThreadPoolExecutor

cout_th = int(input("Nhap So luong` can chay: (1-10): "))

# Đọc tệp Excel
file_path = "data.xlsx"  # Thay bằng đường dẫn thực tế của tệp
df = pd.read_excel(file_path)

# Lấy dữ liệu từ cột "Searching URLs"
searching_urls = df["Searching URLs"]

# Danh sách chứa dữ liệu
data_list = []

# Hàm xử lý dữ liệu của từng URL
def process_url(i, count):
    status = "Active"
    print(f"Round: {count}")
    match = re.search(r"searchTerm=(\d+)", i)
    if match:
        key_p = match.group(1)
    else:
        key_p = None

    response = requests.get(f'https://www.target.com/p/-/A-{key_p}')
    data_html = response.text
    
    # get buy_url
    buy_url = None
    match = re.search(r'"buy_url\\":\\"(https://.*?)\\"', data_html)
    if match:
        buy_url = match.group(1)

    # get title
    title = None
    title_dec = None
    match = re.search(r'product-title[^>]*>(.*?)</h1>', data_html)
    if match:
        title = match.group(1)
        if "This item is not available" in title:
            status = "Obsolete"
            try:
                title = title.split('alt="')[1].split('"')[0] 
            except:
                pass
        title_dec = html.unescape(title) 

    # get cur_price
    cur_price = None
    if 'current_retail\\":' in data_html:
        cur_price = data_html.split('current_retail\\":')[1].split(",")[0]

    # get reg_retail
    reg_retail = None
    if 'reg_retail\\":' in data_html:
        reg_retail = data_html.split('reg_retail\\":')[1].split(",")[0]

    # get save_dollar & save_percent
    save_dollar, save_percent = None, None
    try:
        if 'save_dollar\\":' in data_html:
            save_dollar = data_html.split('save_dollar\\":')[1].split(",")[0]
        if 'save_percent\\":' in data_html:
            save_percent = data_html.split('save_percent\\":')[1].split(",")[0]
    except:
        pass

    # Thêm vào danh sách
    # Chuyển cur_price thành float nếu có dữ liệu
    if cur_price:
        try:
            cur_price = float(cur_price)
        except ValueError:
            cur_price = None

    # Chuyển reg_retail thành float nếu có dữ liệu
    if reg_retail:
        try:
            reg_retail = float(reg_retail)
        except ValueError:
            reg_retail = None

    # Chuyển save_dollar thành float nếu có dữ liệu
    if save_dollar:
        try:
            save_dollar = float(save_dollar)
        except ValueError:
            save_dollar = None

    # Chuyển save_percent thành float nếu có dữ liệu
    if save_percent:
        try:
            save_percent = float(save_percent)
        except ValueError:
            save_percent = None

    data_list.append([i, buy_url, key_p, status, title_dec, cur_price, reg_retail, save_dollar, save_percent])
    print("URL Product:", buy_url)
    print("Title:", title_dec)
    print("cur_price:", cur_price)
    print("reg_retail:", reg_retail)
    if save_dollar:
        print("save_dollar:", save_dollar)
    print("_".center(100, "_"), "\n")

# Chạy đa luồng với ThreadPoolExecutor
with ThreadPoolExecutor(max_workers=cout_th) as executor:
    for count, i in enumerate(searching_urls, start=1):
        executor.submit(process_url, i, count)

# Tạo DataFrame
df_output = pd.DataFrame(data_list, columns=[
    "Searching URLs", "Item URLs", "TCIN", "Status", "Item Title", "Item Retail Price",
    "Regular Retail Price", "Save Dollar", "Sale Off %"
])

# Xuất ra file Excel
df_output.to_excel("output.xlsx", index=False)
print("Đã lưu file output.xlsx thành công!")
