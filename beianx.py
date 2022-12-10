import argparse
import time
import urllib.parse

import pandas as pd
import requests
import urllib3
import xlwt
from lxml import etree

urllib3.disable_warnings()


def save_excel(datas, keyword):
    now_time = time.strftime('%Y%m%d%H%M%S', time.localtime(time.time()))
    File_Found = xlwt.Workbook(encoding='utf-8')  # 工作簿
    File_Sheet = File_Found.add_sheet(keyword, cell_overwrite_ok=True)  # 工作表
    File_Sheet.col(0).width = 256 * 30
    File_Sheet.col(1).width = 256 * 20
    File_Sheet.col(2).width = 256 * 20
    File_Sheet.col(3).width = 256 * 20
    File_Sheet.col(4).width = 256 * 10
    File_Sheet.col(5).width = 256 * 30
    File_Sheet.col(6).width = 256 * 30
    col = ('主办单位名称', '域名', '备案许可证号', '备案号', '单位性质', '网站名称', '网站首页')
    print(f'\n获取结果数量：{len(datas)}')
    for i in range(len(datas)):
        data = datas[i]
        for k in range(0, 7):
            File_Sheet.write(0, k, col[k])
            File_Sheet.write(i + 1, k, data[k])
    File_Found.save(f"{now_time}.xls")
    print(f'Excel文件生成成功，已保存至{now_time}.xls\n')


def sendate(con, output):
    datas = []
    key = urllib.parse.quote(con)
    url = f"https://www.beianx.cn:443/search/{key}"
    cookies = {"HWWAFSESTIME": "1657250177535", "HWWAFSESID": "510ea7c9cb28802932",
                     "__51huid__JfwpT3IBSwA9n8PZ": "2d542d45-8b1a-562b-b897-7cfbb8aa8deb",
                     "__51uvsct__JfvlrnUmvss1wiTZ": "1",
                     "__51vcke__JfvlrnUmvss1wiTZ": "3270ccc4-ace1-578b-8f94-354e79bae94e",
                     "__51vuft__JfvlrnUmvss1wiTZ": "1657250179340",
                     ".AspNetCore.Antiforgery.1QshF0qLdoU": "CfDJ8CV960aQmVdCqF54fR0paPCgalP6b-aYmayuwueaGlF-LSw9OL0BzmNDtELlpG91tphIp0zYvEtfODVD_iOLqoDCF1J1Ta9jUB7OALSgG-fKsh97BOj3DUKVQkE8vpWmQzNJjx36elTBXzi8RkRFxds",
                     "__vtins__JfvlrnUmvss1wiTZ": "%7B%22sid%22%3A%20%2268c4fb78-802d-579d-819a-8c5e6bc5dc32%22%2C%20%22vd%22%3A%2011%2C%20%22stt%22%3A%20317656%2C%20%22dr%22%3A%205029%2C%20%22expires%22%3A%201657252296993%2C%20%22ct%22%3A%201657250496993%7D"}
    headers = {"Cache-Control": "max-age=0", "Upgrade-Insecure-Requests": "1",
                     "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/103.0.0.0 Safari/537.36",
                     "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9",
                     "Sec-Fetch-Site": "same-origin", "Sec-Fetch-Mode": "navigate", "Sec-Fetch-User": "?1",
                     "Sec-Fetch-Dest": "document",
                     "Sec-Ch-Ua": "\".Not/A)Brand\";v=\"99\", \"Google Chrome\";v=\"103\", \"Chromium\";v=\"103\"",
                     "Sec-Ch-Ua-Mobile": "?0", "Sec-Ch-Ua-Platform": "\"Windows\"",
                     "Referer": "https://www.beianx.cn/search", "Accept-Encoding": "gzip, deflate",
                     "Accept-Language": "zh-CN,zh;q=0.9,en-US;q=0.8,en-GB;q=0.7,en;q=0.6", "Connection": "close"}
    res = requests.get(url, headers=headers, cookies=cookies, verify=False)

    html = etree.HTML(res.text)
    table = html.xpath('/html/body/div[2]/table')
    table = etree.tostring(table[0], encoding='utf-8').decode()
    df = pd.read_html(table, encoding='utf-8', header=0)[0]
    results = list(df.T.to_dict().values())  # 转换成列表嵌套字典的格式
    for i in range(len(results)):
        s_number = results[i]['序号']
        o_name = results[i]['主办单位名称']
        o_nature = results[i]['主办单位性质']
        website_number = results[i]['网站备案号']
        website_name = results[i]['网站名称']
        if isinstance(website_name, float):
            website_name = ""
        website_address = results[i]['网站首页地址']
        audit_date = results[i]['审核日期']
        restrict_access = results[i]['是否限制接入']
        if not "没有查询到记录" in o_name:
            length = len(results[i]['网站备案号'].split("-"))
            if length == 2:
                permitNumber = results[i]['网站备案号'].split("-")[0]
            elif length == 3:
                permitNumber = results[i]['网站备案号'].split("-")[0] + "-" + results[i]['网站备案号'].split("-")[1]

            domain = results[i]['网站首页地址'].split(".", 1)[1]

        print(
            f"序号：{s_number}  主办单位名称：{o_name}  域名：{domain}   备案许可证号：{permitNumber}   备案号：{website_number}  单位性质：{o_nature}   网站名称：{website_name} 网站首页：{website_address}")
        datas.append((o_name, domain, permitNumber, website_number, o_nature, website_name, website_address))
    if output:
        save_excel(datas, con)


def run(args):
    output = True if args.output else False
    if args.file:
        with open(args.file, 'r', encoding='utf-8') as file:
            for x in file.readlines():
                target = x.strip("\n")
                sendate(target, output)
    elif args.target:
        sendate(args.target, output)


if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    parser.add_argument('-t', dest='target', help='需查询的目标（域名、备案号、公司名称）')
    parser.add_argument('-f', dest='file', help='从文件导入查询目标')
    parser.add_argument('-o', dest='output', help='结果是否导出Excel（default：false）')
    args = parser.parse_args()
    run(args)
