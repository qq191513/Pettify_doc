# coding=utf-8
from bs4 import BeautifulSoup
from win32com import client as wc
import shutil
from os import path,remove
import os
import re


def change_picShow(input_file,output_dir,output_file):

    if not path.exists(output_dir):
        os.makedirs(output_dir)

    file_xml = r'temp.html'
    file_xml_rm = path.join(output_dir,file_xml.replace('html','files'))

    word = wc.Dispatch('Word.Application')
    doc1 = word.Documents.Open(input_file)
    doc1.SaveAs(path.join(output_dir,file_xml),8)
    doc1.Close

    my_width = 1
    my_height = 400

    word.Visible = False
    word.DisplayAlerts = 0

    word_soup = BeautifulSoup(open(path.join(output_dir,file_xml)),'lxml')

    for tag in word_soup.find_all('p',class_='MsoNormal'):
        for tag1 in tag.find_all(re.compile(r"^img$")):
            # if(tag1["height"] == my_height):
            #     continue

            if(int(tag1["width"])>30):
                p = int(tag1["height"]) / my_height
                tag1["height"] = my_height

                my_width =str(int(tag1["width"])/p)
                tag1["width"] = my_width

                tag["style"] = tag["style"].replace("left", "center")
                print(tag)
                # style = "text-indent:24.0000pt;mso-char-indent-count:2.0000;" \
                #         "mso-pagination:widow-orphan;text-align:center;"

        # temp_str1 = re.sub(re.compile(r"width:[0-9]*\.[0-9]+]"),tag["width"],("%d" % my_width))
        # if(int(temp_str1)>20):
        #     tag["width"] = temp_str1
        # print(tag["width"])

        # temp_str1 = re.sub(re.compile(r"height:[0-9]*\.[0-9]+]"),tag["height"],("%d" % my_height))
        # if (int(temp_str1) > 20):
        #     tag["height"] = temp_str1
        #
        # print(tag)


    word.Quit()
    # remove(path.join(work_path,file_xml))
    with open(path.join(output_dir,file_xml),"wb") as file:
        # print(word_soup.prettify(word_soup.original_encoding))
        file.write(bytes(word_soup.prettify(word_soup.original_encoding),encoding ='utf-8'))

        # file.write(bytes("hello\n",encoding ='utf-8'))
        file.close()

    word = wc.Dispatch('Word.Application')
    doc2 = word.Documents.Add(path.join(output_dir,file_xml))
    doc2.SaveAs(path.join(output_dir,output_file))
    doc2.Close()
    word.Quit()
    remove(path.join(output_dir,file_xml))
    shutil.rmtree(path.join(output_dir,file_xml_rm))



if __name__ == '__main__':

    # 手动输入一个
    path_file =r'F:\test\pythonTest\doc2\log\2016.docx'
    name_doc1 =path.basename(path_file)
    name_doc2 =name_doc1.replace('.docx','_v1.docx')
    output_dir =path.join(path.dirname(path_file),'output_dir')
    change_picShow(path_file,output_dir,name_doc2)


    # 遍历所有文件
    # path_file = r'F:\test\pythonTest\doc2\log'
    # list = os.listdir(path_file)  # 列出文件夹下所有的目录与文件
    # for i in range(len(list)):
    #     path = os.path.join(path_file, list[i])
    #     if os.path.isfile(path):
    #         print(list[i])





