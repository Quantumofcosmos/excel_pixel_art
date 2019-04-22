import xlsxwriter
from skimage import io

img = io.imread('exp1.jpg')
workbook = xlsxwriter.Workbook('hello_world.xlsx')
worksheet = workbook.add_worksheet()

for row in list(range(0,img.shape[0])):
    for col in list(range(0,img.shape[1])):
        r,g,b=img[row,col].tolist()
        hexx=hex(r)+hex(g)[2:]+hex(b)[2:]
        hexx='#'+hexx[2:]
        print(hexx)
        cell_format = workbook.add_format()
        cell_format.set_bg_color(hexx)
        worksheet.write(row,col, ' ',cell_format)
#       break
#    break


workbook.close()

# img[10,20]
# array([127, 138, 106], dtype=uint8)
#>>> from skimage import io
#>>> img = io.imread('exp1.jpg')

