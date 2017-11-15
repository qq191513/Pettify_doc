# coding=utf-8
import win32com
from win32com.client import Dispatch, DispatchEx


word = Dispatch('Word.Application')  # 打开word应用程序
# word = DispatchEx('Word.Application') #启动独立的进程
word.Visible = 0  # 后台运行,不显示
word.DisplayAlerts = 0  # 不警告
path = 'F:/test/pythonTest/doc2/log/200520062007_1.docx' # word文件路径
doc = word.Documents.Open(FileName=path, Encoding='gbk')

print('----------------')
print('段落数: ', doc.Paragraphs.count)
print ('-------------------------')

#改字体
iter_Paragraph = doc.Paragraphs.count
for i in range(iter_Paragraph):

    if(i>= iter_Paragraph):

        break

    para = doc.Paragraphs[i]
    print("%d : %s" % (i, para.Range.text))
    if(para.Range.text.find(u"年") != -1):
        if (para.Range.text.find(u"月") != -1):
            if (para.Range.text.find(u"天气") != -1):
                # para.Range.Font.Size = (16)  #三号
                # para.Range.Font.Size = (15)  #小三   标题
                  para.Range.Font.Size = (14)  #四号   日期
                # para.Range.Font.Size = (13)  #13号
                # para.Range.Font.Size = (12)  #小四   正文
                  print("%d : %s" % (i,para.Range.text))


                  para = doc.Paragraphs[i-2]
                  para.Range.Font.Size = (15)  #小三   标题

                  para = doc.Paragraphs[i - 1]
                  text_len = len(para.Range.text)
                  if(text_len <=4):
                      para.Range.text =''
                      iter_Paragraph = iter_Paragraph-1
                  else:
                      para = doc.Paragraphs[i - 1]
                      para.Range.Font.Size = (15)  # 小三   标题
    else:
        print("%d : no" % i)


doc.Save()     # 存檔
doc.Close()  # 关闭word文档




#change





