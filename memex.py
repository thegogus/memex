import xlsxwriter
from PIL import Image
def rgb2hex(r, g, b):
    return '#{:02x}{:02x}{:02x}'.format(r, g, b)

def getColumn(num):
    ch = 65
    zk = 64
    rt = num // 26
    lst = num - (26 * rt)
    if rt == 0:
        out = chr(64+lst)
    else:
        zk += rt
        out = chr(zk) + chr(64+lst)
    return out



img = Image.open('./img.png')
pix = img.load()
meme = xlsxwriter.Workbook('./meme.xlsx')
worksheet = meme.add_worksheet()

for i in range(img.size[1]):
    for x in range(0,img.size[0]):
        color = rgb2hex(pix[x,i][0], pix[x,i][1], pix[x,i][2])
        format = meme.add_format()
        format.set_bg_color(color)
        worksheet.write_blank(i+3, x, '', format)

for x in range(3, img.size[1]+3):
    worksheet.set_row(x, 6, None, None)
column = 'A:' + getColumn(img.size[0])
worksheet.set_column(column, 0.75, None, None)
meme.close()
