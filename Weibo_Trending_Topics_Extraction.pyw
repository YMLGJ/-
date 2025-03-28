from selenium import webdriver
from selenium.webdriver.edge.service import Service as EdgeService
from selenium.webdriver.edge.options import Options as EdgeOptions
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from datetime import datetime
import time
import pandas as pd

# 指定 EdgeDriver 的路径
edgedriver_path = "E:\\Application program\\edgedriver\\msedgedriver.exe"

# 创建一个 EdgeOptions 对象
edge_options = EdgeOptions()

# 设置浏览器全屏
edge_options.add_argument("--start-maximized")
edge_options.add_argument('--disable-gpu')
edge_options.add_argument('--headless')
edge_options.add_argument('--allow-running-insecure-content')
edge_options.add_argument('--ignore-certificate-errors')

# 创建一个 EdgeService 对象
service = EdgeService(executable_path=edgedriver_path)

# 启动 Edge 浏览器
driver = webdriver.Edge(service=service, options=edge_options)

# 打开目标网址
driver.get("https://weibo.com/hot/search")  # 以微博为例

# 等待页面加载完成
time.sleep(5)

# 添加 cookie，模拟登录
# 这里的 cookie 需要根据实际登录后的 cookie 进行替换
# cookie 的格式为 {'name': 'cookie_name', 'value': 'cookie_value', 'domain': 'weibo.com'}
cookies = [
    {"name": "SINAGLOBAL", "value": "2725083852807.675.1718858016495", "domain": "weibo.com"},
    {"name": "SCF", "value": "OEVpDrrd8n37HSwWzzVM2-xD_wkXM8RaXB-CdwuVCbr8sbT4IbrmkPICMRf0tni9qo4WkzCilsROOrCSXH9eyw.", "domain": "weibo.com"},
    {"name": "SUB", "value": "_2A25K253YDeRhGeFH41sV9SfMzz2IHXVpmJ8QrDV8PUNbmtAYLXbskW9NejSJxzvvibLBqVL37ZjZqvXamG8BQgQ5", "domain": "weibo.com"},
    {"name": "SUBP", "value": "0033WrSXqPxfM725Ws9jqgMF55529P9D9WFejspRO-ypc2V5zes8BQhW5NHD95QN1Kn4Sh-4ehBpWs4DqcjMi--NiK.Xi-2Ri--ciKnRi-zNS0.R1KBf1K5XeBtt", "domain": "weibo.com"},
    {"name": "ALF", "value": "02_1745320584", "domain": "weibo.com"},
    {"name": "_s_tentry", "value": "cn.bing.com", "domain": "weibo.com"},
    {"name": "UOR", "value": ",,cn.bing.com", "domain": "weibo.com"},
    {"name": "Apache", "value": "3836407582383.925.1742789053851", "domain": "weibo.com"},
    {"name": "ULV", "value": "1742789053852:4:3:3:3836407582383.925.1742789053851:1742739233553", "domain": "weibo.com"},
    {"name": "XSRF-TOKEN", "value": "DTTnZBT1f94cNOFw5rpt5-xf", "domain": "weibo.com"},
    {"name": "WBPSESS", "value": "b4Nof7TUCgiz7ZaZYzH9kzkRBczHgkwlrnNyBFMBaTGm-LMy8VAuLLKm2PjI3M8m33CT8uPGuo_saTuVATnEjIXxk-k6resLfvFranvlfga9hjvnO4SLnHMhK7fu_wbtdUhvWXLTpYwXpMHatJj0NQ==", "domain": "weibo.com"}
]

# 将列表中的每个 cookie 添加到浏览器会话中
for cookie in cookies:
    driver.add_cookie(cookie)

time.sleep(5)

try:
    # 等待按钮可点击
    button = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.CSS_SELECTOR, "button.woo-button-main.woo-button-flat.woo-button-primary.woo-button-m.woo-button-round.woo-dialog-btn"))
    )
    
    # 尝试直接点击按钮
    button.click()
except Exception as e:
    print(f"直接点击按钮失败: {e}")
    # 如果直接点击失败，尝试滚动页面
    driver.execute_script("arguments[0].scrollIntoView(true);", button)
    try:
        # 再次尝试点击
        button.click()
    except Exception as e:
        print(f"滚动后点击按钮失败: {e}")
        # 如果仍然点击失败，使用 JavaScript 点击
        try:
            driver.execute_script("arguments[0].click();", button)
        except Exception as e:
            print(f"使用 JavaScript 点击按钮失败: {e}")

# 设置等待条件，确保页面元素加载完成
try:
    wait = WebDriverWait(driver, 10)
    # 根据页面结构调整这里的选择器，这里选择class="HotTopic_item_VkggB"下的热点元素
    elements = wait.until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, '.HotTopic_item_VkggB')))
except Exception as e:
    print(f"等待页面元素加载时出错: {e}")

# 初始化数据列表
new_data = []
nowtime = datetime.now()

# 如果你想以特定格式存储，比如“年-月-日 时:分”
datanowtime = nowtime.strftime("%H:%M")

# 获取当前日期
current_date = datetime.now().strftime("%Y-%m-%d")

try:
    existing_df = pd.read_excel(f'微博热点{current_date}.xlsx')
except FileNotFoundError:
    existing_df = pd.DataFrame(columns=['热点名称'])

for element in elements:
    # 查找热点名称
    try:
        span_element = element.find_element(By.CSS_SELECTOR, 'a.HotTopic_tit_eS4fv span')
        textcut_text = span_element.text if span_element else 'N/A'
    except Exception as e:
        print(f"查找热点名称时出错: {e}")
        textcut_text = 'N/A'

    # 查找关注度数字
    try:
        clb_element = element.find_elements(By.CSS_SELECTOR, '.HotTopic_num_1H-j8')
        clb_text = clb_element[0].text if clb_element else 'N/A'
    except Exception as e:
        print(f"查找关注度数字时出错: {e}")
        clb_text = 'N/A'

    # 检查是否有关注度数字，如果有则添加到DataFrame中
    if clb_text != 'N/A':
        if textcut_text in existing_df['热点名称'].values:
            index = existing_df['热点名称'] == textcut_text
            existing_df.loc[index, datanowtime] = clb_text
        else:
            # 如果热点名称不存在于Excel中，则添加新的行
            new_data.append({
                '热点名称': textcut_text,
                datanowtime: clb_text
            })

print(new_data)
# 创建 DataFrame
df = pd.DataFrame(new_data)

# 合并已有的DataFrame和新数据的DataFrame
final_df = pd.concat([existing_df, df], ignore_index=True)

# 将数据写入Excel文件
final_df.to_excel(f'微博热点{current_date}.xlsx', index=False)

print("数据已成功写入Excel文件")

# 关闭浏览器
driver.quit()
