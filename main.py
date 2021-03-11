import xlrd
import json

loc = ("D:/ma_mau_dat.xlsx")
wb = xlrd.open_workbook(loc)
sheet = wb.sheet_by_index(0)

data = []
for i in range(sheet.nrows):
    if (i != 0):
        R = int(float(sheet.cell_value(i, 3)))
        G = int(float(sheet.cell_value(i, 4)))
        B = int(float(sheet.cell_value(i, 5)))
        hex = '#%02x%02x%02x' % (R, G, B)

        data.append({
            'name': sheet.cell_value(i, 1),
            'ma': sheet.cell_value(i, 2),
            'color_hex': hex
        })

with open('D:/data.json', 'w') as outfile:
    json.dump(data, outfile)



