import fitz
import os

file_path = os.path.dirname(os.path.abspath(__file__))
if not os.path.exists(file_path+'/pictures'):
    os.mkdir(file_path+'/pictures')
    save_file_path = os.mkdir(file_path+'/pictures')
    print(save_file_path)

rotate = int(0)
zoom_x = 2.0
zoom_y = 2.0
trans = fitz.Matrix(zoom_x, zoom_y).preRotate(rotate)
open_file_path = 'E:/python/PDF2image/01.pdf'
save_file_path = 'E:/python/PDF2image/pictures'

pdf = fitz.open(open_file_path)
for i in range(pdf.pageCount):
    pm = pdf[i].getPixmap(matrix=trans, alpha=False)
    pm.writePNG(save_file_path + '/%s.png' % i)