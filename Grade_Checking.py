
from selenium import webdriver
from selenium.webdriver.common.action_chains import ActionChains    # 鼠标操作
from selenium.webdriver.chrome.options import Options   # 配置参数
import time
from selenium.webdriver.common.by import By  # 导入By类

import openpyxl

#打开Excel文件
workbook = openpyxl.load_workbook(r'C:\Users\65349\Desktop\test.xlsx')

#选择指定的工作表
worksheet = workbook['Sheet1']  # 替换为你的工作表名称

for row in worksheet.iter_rows():
    username = row[0].value
    name = row[1].value

    print(username,name)

    try:
        # 创建ChromeOptions对象
        chrome_options = Options()


        # 启用无痕模式
        chrome_options.add_argument("--incognito")

        # 创建Chrome浏览器对象并传入选项
        driver = webdriver.Chrome(options=chrome_options)

        driver.get("https://824305.yichafen.com/public/queryscore/sqcode/NsDcEnzmNDgzNnwwZDYyYjI4M2ZiNDliNzE5Y2YyNTJmN2ZlOGQyZWRmY3w4MjQzMDUO0O0O.html")  # get方式访问查分网站.

        driver.find_element(By.XPATH,"/html/body/div/div[2]/div[4]/div/div[2]/form/table/tbody/tr[1]/td[2]/input").send_keys(username)  # 输入学号
        driver.find_element(By.XPATH,"/html/body/div/div[2]/div[4]/div/div[2]/form/table/tbody/tr[2]/td[2]/input").send_keys(name)  # 输入姓名

        driver.find_element(By.XPATH, "/html/body/div/div[2]/div[4]/div/div[2]/form/div/button").click()  # 点击登陆

        time.sleep(1)
        tbody_element = driver.find_element(By.XPATH, "/html/body/div/div[2]/div[1]/div[2]/table/tbody/tr[2]")

        text_content = tbody_element.text
        print(text_content)

        words = tbody_element.text.split()
        print(words)

        row[2].value = words[4]
        row[3].value = words[5]

    except:
        print("查询失败", name)
        row[2].value = '查询失败'


#保存Excel文件
workbook.save(r'C:\Users\65349\Desktop\test.xlsx')  # 替换为你的保存路径
# 关闭Excel文件
workbook.close()

