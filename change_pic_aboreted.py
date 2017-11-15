import os
import os.path
from PIL import Image
import glob

def ResizeImage(filein, fileout, width, height, type):



    img = Image.open(filein)
    if(img.width<50):
        return -1
    out = img.resize((width, height), Image.ANTIALIAS)  # resize image with high-quality
    out.save(fileout, type)


if __name__ == "__main__":



    for filein in glob.iglob(r'F:\test\pythonTest\doc2\2016test.files\*.png'):
        width = 500
        height = 660
        type = 'png'
        ResizeImage(filein, filein, width, height, type)















