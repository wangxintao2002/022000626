import asyncio
from pyppeteer import launch
import os
from lxml import etree
import time
import random





async def pyppteer_fetchContent(url):

    #以headless模式打开浏览器，将浏览器处理标准输出导入process.stdout，并设置自动关闭
    browser = await launch({'headless': True, 'dumpio': True, 'autoClose': True})
    page = await browser.newPage()
    # 绕过浏览器检测
    await page.evaluateOnNewDocument('() =>{ Object.defineProperties(navigator,'
                                     '{ webdriver:{ get: () => false } }) }')

    #设置UA，模拟人手动打开页面
    await page.setUserAgent(
        'Mozilla/5.0 (Linux; Android 6.0; Nexus 5 Build/MRA58N) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/105.0.0.0 Mobile Safari/537.36 Edg/105.0.1343.27')
    await page.goto(url)
    # await asyncio.sleep(10)

    #等待页面转到一个新链接或者是刷新
    await asyncio.wait([page.waitForNavigation()])

    #获取爬下来的内容
    str = await page.content()
    await browser.close()
    return str


# 获取每一个有防控数据页面的内容（页面源代码）
def get_page_souce(url):

    #获取当前事件循环，并不断获取每一个页面的内容
    return asyncio.get_event_loop().run_until_complete(pyppteer_fetchContent(url))


#总共42页数据，获取每一页的url
def get_page_url():
    url_list = []

    #遍历42页
    for page in range(1, 42):
        if page == 1:
            url_list.append('http://www.nhc.gov.cn/xcs/yqtb/list_gzbd.shtml')
        else:
            page_url = 'http://www.nhc.gov.cn/xcs/yqtb/list_gzbd_' + str(page) +'.shtml'
            url_list.append(page_url)
    return url_list


# 通过 getDateUrl发布日期。
def get_files(html):
    file_list =[]
    html_tree = etree.HTML(html)
    li_lsit = html_tree.xpath("/html/body/div[3]/div[2]/ul/li")
    for li in li_lsit:
        file = li.xpath("./a/text()")
        file = "".join(file)
        date = li.xpath('./span/text()')
        date = "".join(date)
        file_name = date + "-" + file
        file_list.append(file_name)
    return file_list


#拼接要爬取的目标页面的url
def get_link_url(html):
    day_urls = []
    html_tree = etree.HTML(html)

    #解析出url位置
    li_list = html_tree.xpath("/html/body/div[3]/div[2]/ul/li")
    for li in li_list:
        day_url = li.xpath('./a/@href')
        day_url = "http://www.nhc.gov.cn" + "".join(day_url[0])
        day_urls.append(day_url)
    return day_urls

#获取真正的、有用的内容(每日新增确诊，无症状)
def get_content(html):
    html_tree = etree.HTML(html)
    text = html_tree.xpath('/html/body/div[3]/div[2]/div[3]//text()')
    day_text = "".join(text)
    return day_text


#保存文件
def save_file(path, filename, content):
    if not os.path.exists(path):
        os.makedirs(path)
    # 保存文件
    with open(path + filename + ".txt", 'w', encoding='utf-8') as f:
        f.write(content)


if "__main__" == __name__:
    start = time.time()
    for url in get_page_url():
        s = get_page_souce(url)
        time.sleep(random.randint(3, 6))
        filenames = get_files(s)
        index = 0
        links = get_link_url(s)
        for link in links:
            html = get_page_souce(link)
            content = get_content(html)
            print(filenames[index] + "爬取成功")
            save_file('D:/数据/', filenames[index], content)
            index = index + 1
        print("-----"*20)
    end = time.time()
    print({end-start})
