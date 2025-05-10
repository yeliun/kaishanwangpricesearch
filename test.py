import requests
from bs4 import BeautifulSoup
from urllib.parse import quote
import re
import pandas as pd


# 读取 Excel 文件中的商品编码
def read_product_codes(file_path):
    df = pd.read_excel(file_path)
    return df['商品编码'].astype(str).tolist()  # 确保所有商品编码为字符串


# 获取商品价格
def get_product_price(product_code):
    try:
        encoded_code = quote(product_code, safe='')
        url = f"http://xinyu.k3.cn/search/web,xinyu,{encoded_code},,1,0.html"
        headers = {
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36",
            "Referer": "http://xinyu.k3.cn/",
        }
        cookies = {
            "user_username": "",
            "user_hash": "",
            "user_login_type": "",
            "user_is_user_login": "1",
            "daily_login": "1",
            "k3cn": "",
            "user_login_ip": "",
            "user_login_time": "",
            "user_phash": "",
            "user_type": "0",
            "user_user_id": ""
        }

        resp = requests.get(url, headers=headers, cookies=cookies)
        if resp.status_code != 200:
            print(f"❌ 请求失败，商品编码: {product_code}")
            return None

        resp.encoding = resp.apparent_encoding
        soup = BeautifulSoup(resp.text, "html.parser")
        items = soup.select("li")

        for item in items:
            title_tag = item.select_one('a.tradeName')
            price_tag = item.select_one('div.price > b')

            if title_tag:
                title_text = re.sub(r'\s+', '', title_tag.text).lower()  # 标题去空格转小写
                code_lower = product_code.lower()  # 编码转小写

                if code_lower in title_text:
                    return price_tag.text.strip() if price_tag else None
        return None
    except Exception as e:
        print(f"⚠️ 商品编码 {product_code} 出现错误: {e}")
        return None


# 保存结果到 Excel
def save_to_excel(product_data, output_file):
    df = pd.DataFrame(product_data, columns=['商品编码', '商品价格'])
    df.to_excel(output_file, index=False)
    print(f"✅ 已保存至: {output_file}")


# 主程序
def main(input_file, output_file):
    product_codes = read_product_codes(input_file)
    product_data = []

    for code in product_codes:
        print(f"🔍 查询: {code}")
        price = get_product_price(code)
        if price:
            print(f"✅ 价格: ¥{price}")
            product_data.append([code, price])
        else:
            print(f"⚠️ 未找到价格: {code}")
            product_data.append([code, '未找到价格'])

    save_to_excel(product_data, output_file)


# 路径设置
input_file = "D:\\Downloads\\1.xlsx"  # 输入文件路径
output_file = "D:\\Downloads\\2.xlsx"  # 输出文件路径

main(input_file, output_file)
