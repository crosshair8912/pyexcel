from xlutils.copy import copy
from xlrd import open_workbook
from os import listdir
from os.path import isfile, join
from datetime import datetime,timedelta
from xlwt import easyxf
import sys,getopt,math

def excel_time_to_string(excel_time, fmt):
    dt = datetime(1899,12,30) + timedelta(days=excel_time)
    return dt.strftime(fmt)	
	
def main(argv):
	#------------Variables--------------------------------
	input_path = ''
	sheet_name = ''
	date_cols = [8,9,10,11]
	percentage_cols = [14,25,26]
	#------------Zone for script parameters----------------
	try:
		opts, args = getopt.getopt(argv,"hi:s:",["ifile=","sname="])
	except getopt.GetoptError:
		print ('test.py -i <input_path> -s <sheet_name>')
		sys.exit(2)
	for opt, arg in opts:
		if opt == '-h':
			print ('test.py -i <input_path> -s <sheet_name>')
			sys.exit()
		elif opt in ("-i","--ifile"):
			input_path = arg
		elif opt in ("-s","--sname"):
			sheet_name = arg
	#------------Working with data zone---------------------		
	onlyfiles = [f for f in listdir(input_path) if (isfile(join(input_path, f)) and not(f.startswith('~')))]
	
	for i in onlyfiles:
		if '.xlsx' in i:
			file = open_workbook(input_path+i)
			tmp = copy(file)
			w_sheet = tmp.get_sheet(0)
			
			for rownum in range(1,file.sheet_by_index(0).nrows):
				cell_values = file.sheet_by_index(0).row_values(rownum)
				#Changing cells to display $
				if cell_values[16]!='':
					w_sheet.write(rownum,16,'$'+str(cell_values[16]))
				#Changing cells to display time correctly
				for col in date_cols:
					if type(cell_values[col]) == type(3.14):
						cell_values[col] = excel_time_to_string(int(cell_values[col]), '%Y%m%d')
						w_sheet.write(rownum,col,cell_values[col])
				#Changing cells to display %		
				for col in percentage_cols:
					style_percent = easyxf(num_format_str='0.0%')
					w_sheet.write(rownum,col,cell_values[col],style_percent)
			
			#Renaming list zone
			if file.sheet_names() != list('Page1'):
				tmp.get_sheet(0).name = u'Page1'
				tmp.save(input_path + i[:len(i)-5]+'.xls')
				
if __name__ == "__main__":
   main(sys.argv[1:])