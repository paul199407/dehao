import xlrd
import xlwt
from decimal import Decimal

book = xlrd.open_workbook("1.xlsx")
print ("\ntable count: {}".format(book.nsheets))

workbook = xlwt.Workbook(encoding = 'utf-8')
worksheet = workbook.add_sheet('用地统计表')

worksheet.write(0, 0, label = '用地类型\所属范围')

mc_list = ['mc', '用地大类', 'Shape_Area']

type_list = ['住宅', '工业用地', '行政办公', '商业金融', '教育文体', '医疗卫生', '公共设施', '备用地', '绿地', '交通用地和水域']

area_list = []
sh = book.sheet_by_index(0)

for r in range(1, sh.nrows):
    if (area_list.count(sh.cell_value(r,1))) <= 0:    area_list.append(sh.cell_value(r,1))

for a in range(0, len(type_list)):
    worksheet.write(a+1, 0, label = type_list[a])
	
for b in range(0, len(area_list) ):
    worksheet.write(0, b+1, label = area_list[b])

for x in range(0, len(type_list) ):
    temp_str = type_list[x];

    for y in range(0, len(area_list)):	
        temp_str1 = area_list[y];
        count_value = 0.00;
        start_row = 1
        for r in range(start_row, sh.nrows):
            start_col = 1
            value1 = ""
            value2 = ""
            value3 = ""
			
            for c in range(start_col, sh.ncols):
                if(sh.cell_value(0, c) == "mc"):    value2 = sh.cell_value(r, c)
                elif(sh.cell_value(0, c) == "用地大类"):    value1 = sh.cell_value(r, c)
                elif(sh.cell_value(0, c) == "Shape_Area"):    value3 = sh.cell_value(r, c)
			
            if ( (temp_str == value2) and (temp_str1 == value1) ):    
                count_value += float(value3) / 10000.0
        
        worksheet.write(x+1, y+1, str(Decimal(str(count_value)).quantize(Decimal('0.000000'))))	

workbook.save('宋慧乔.xls')
 
print('done')