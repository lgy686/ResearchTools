# -*- coding: utf-8 -*-
import requests
from bs4 import BeautifulSoup

# 测试GPS SOLUTIONS的搜索URL
journal_name = "GPS SOLUTIONS"
search_url = 'https://www.letpub.com.cn/index.php?page=journalapp&view=search&searchname={}'.format(journal_name)
headers = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
}

print("测试搜索URL: {}".format(search_url))
try:
    response = requests.get(search_url, headers=headers, timeout=10)
    response.encoding = 'utf-8'
    print("响应状态码: {}".format(response.status_code))
    
    # 保存响应内容到文件，方便查看
    with open('search_result.html', 'w', encoding='utf-8') as f:
        f.write(response.text)
    print("搜索结果已保存到 search_result.html")
    
    # 解析HTML
    soup = BeautifulSoup(response.text, 'html.parser')
    
    # 尝试找到搜索结果
    print("\n尝试解析搜索结果...")
    # 查看页面中的链接
    links = soup.find_all('a')
    print("找到 {} 个链接".format(len(links)))
    
    # 打印前20个链接，查看是否有期刊详情页链接
    for i, link in enumerate(links[:20]):
        href = link.get('href', '')
        text = link.text.strip()
        if 'journal' in href or 'GPS' in text.upper():
            print("{}: {} -> {}".format(i+1, text, href))
            
except Exception as e:
    print("错误: {}".format(e))
