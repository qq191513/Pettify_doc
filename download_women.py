# __author__ = 'ccav'
import urllib
import urllib.request
import re
from os import path
import  os

def download_page(url):
    request = urllib.request.Request(url)
    response = urllib.request.urlopen(request)
    data = response.read()
    return  data

def get_image(html):
    regx = r'http://[\S]*\.jpg'
    pattern = re.compile(regx)
    get_img = re.findall(pattern,repr(html))
    num = 1
    if not path.exists('output'):
        os.makedirs('output')

    for img in get_img:
        image = download_page(img)
        with open('output/%s.jpg' % num,'wb') as fp:
            fp.write(image)
            num += 1
            print('第 %d 副图片下载完成' % num)

    return

url = 'http://http://www.mm131.com/qingchun/'
html = download_page(url)
get_image(html)
