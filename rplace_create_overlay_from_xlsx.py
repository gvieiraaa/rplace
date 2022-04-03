import numpy as np
import png
import openpyxl as xl
from PIL import ImageColor

wb = xl.load_workbook(filename='rbrasil.xlsx', data_only=True)
sh = wb["Planilha1"]

dt = np.full(shape=(3000, 6000, 4), dtype=np.uint8, fill_value=(0, 0, 0, 0))

pl_coords_x1, pl_coords_x2 = 23, 200
pl_coords_y1, pl_coords_y2 = 23, 75
pn_coords_x, pn_coords_y = 909, 566

for icolumn, column in enumerate(range(pl_coords_x1, pl_coords_x2 + 1)):
    for irow, row in enumerate(range(pl_coords_y1, pl_coords_y2 + 1)):
        cll = sh.cell(row,column)
        rgb = cll.fill.bgColor.index
        if isinstance(rgb,str):
            color = [i for i in ImageColor.getrgb('#' + rgb[2:])]
            color.append(255)
        else:
            color = [255,255,255,0]
        dt[(pn_coords_y + irow) * 3 + 1, (pn_coords_x + icolumn) * 3 + 1] = np.array(color,dtype=np.uint8)

img = png.from_array(dt.reshape(3000,6000 * 4),'RGBA')
img.save('test.png')
