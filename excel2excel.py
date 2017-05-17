
# Write contents from one excel to another


import xlrd


from xlutils.copy import copy


rb_tar = xlrd.open_workbook('material.xls',formatting_info=True)# xlsx to be written
#   Note that formatting=True works for xls not xlsx


rb_ori = xlrd.open_workbook('group_student.xlsx')   # xlsx for data

sheet_ori = rb_ori.sheet_by_index(0)    # Get the shee reference to use

wb = copy(rb_tar)

sheet_tar = wb.get_sheet(0)


for i_tar in range(2,27):
    sheet_tar.write(i_tar,0,sheet_ori.cell(i_tar+2,4).value)
    sheet_tar.write(i_tar,1,sheet_ori.cell(i_tar+2,5).value)
    sheet_tar.write(i_tar,2,sheet_ori.cell(i_tar+2,8).value)

wb.save('material2.xls')

print('Congratulations!')
