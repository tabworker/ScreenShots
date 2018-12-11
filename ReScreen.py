import time
import xlrd
import os

num = 0
book = xlrd.open_workbook(input('Drag file Excel here to: '))
print('Number of sheets in the workbook {0}'.format(book.nsheets))
k = book.sheet_names()
for i in k:
	print(num, ' - ', k[num])
	num+=1
ind = input('Enter the worksheets number: ')
sheet = book.sheet_by_index(int(ind))
print('Worksheets name - {0}\nNumber of lines - {1}\nNumber of columns - {2}'.format(sheet.name, sheet.nrows, sheet.ncols))
link_of_folder = input('To move a directory with screenshots here: ')
print('Processing ...')
os.chdir(link_of_folder)
print('{}'.format(os.getcwd()))
time.sleep(5)
j=0
number_end = sheet.nrows
for k in range(0, number_end, 1):
	name_files = '{}'.format(sheet.cell_value(k, 3)) #old file
	name_new_files = '{}'.format(sheet.cell_value(k, 0)) #new file
	if os.path.isfile(name_files):
		if name_new_files == '' or name_new_files == ' ' or name_new_files == '.jpg':
			os.rename(name_files, 'Пустое_имя_{}'.format(str(k)))
			print('No file name')
		else:
			os.rename(name_files, name_new_files)
			print('File {0} successfully renamed in {1}')
		j+=1
print(j)
input('Я вроде всё. Для завершения, нажмем Enter')