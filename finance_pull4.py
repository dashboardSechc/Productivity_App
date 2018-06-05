import xlrd
from tkinter import filedialog
import xlwt
import numpy as np
import collections
from xlutils.copy import copy
import string
import pyexcel as p
import os
from shutil import copy2
import datetime



def main(raw_location=0, dash_location=0):

	print("Hello! Program starting")
	#TODO: add in error check if no summary sheet
	
	#inputting the two files to be used
	#raw_location = "C:/Users/jprivera/Documents/Dashboard_contruction/Files/financial_indicators_ws.xlsx"
	#dash_location = "C:/Users/jprivera/Documents/Dashboard_contruction/Files/finance_dashboard_workspace3.xlsx"
	
	#TODO if they dont exist



#hi  this is aptrick K 

	if raw_location ==0:
		raw_location = "P:/Dashboard KPI location/Departments/Finance/financial_indicators.xlsx"

	if dash_location ==0:
		dash_location = "P:/Dashboard KPI location/Departments/Finance/Dashboard_files/finance_dashboard.xlsx"


	directory = "P:/Dashboard KPI location/Departments/Finance/Dashboard_files/old_input_files"
	if not os.path.exists(directory):
		os.makedirs(directory)





	date = str(datetime.date.today())
	#date = "00.00302.2"

	new_loc = "P:/Dashboard KPI location/Departments/Finance/Dashboard_files/old_input_files/" + "finance_dashboards_"+date+".xlsx"


	copy2(dash_location, new_loc)





	#opening the excel sheets and placing them into xlrd book object
	raw_book =  xlrd.open_workbook(raw_location)
	dash_book = xlrd.open_workbook(dash_location)

	
	#copying the book, to then write to it
	copy_book = copy(dash_book)

	#putting the data from the summary sheet from the file in raw location into an xlrd object
	#TODO add in check to use the first sheet if the summary sheet does not exist
	worksheet_summ = raw_book.sheet_by_name('Summary Sheet') 




	#getting the max rows/columns from the summary sheet 
	max_rows = worksheet_summ.nrows
	max_cols = worksheet_summ.ncols



	finance_list = [[]for i in range(max_rows)]

	finance_list = extract_finance_data(worksheet_summ)

	check_data(dash_book,copy_book, finance_list)



	copy_book.save("P:/Dashboard KPI location/Departments/Finance/Dashboard_files/finance_dashboard.xls")

	p.save_book_as(file_name = "P:/Dashboard KPI location/Departments/Finance/Dashboard_files/finance_dashboard.xls", dest_file_name= "P:/Dashboard KPI location/Departments/Finance/Dashboard_files/finance_dashboard.xlsx")


	os.remove("P:/Dashboard KPI location/Departments/Finance/Dashboard_files/finance_dashboard.xls")


	print("Program ending")




def check_data(dash_book,copy_book,finance_list):



	

	for z in range(0,8):
		wsheet_re = dash_book.sheet_by_index(z)
		dash_sheet = copy_book.get_sheet(z)

		print( "main.check sheet",wsheet_re.name)

		dash_rows = wsheet_re.nrows
		dash_cols = wsheet_re.ncols
		sheet_y =0
		sheet_x =0
		finance_x=0

		#range is at list[0]  to just get the length of 1d
		#for i in range(0,len(finance_list[0])):
			#print("main.check",finance_list[1][i])

			#if finance_list[1][i] is "None":
			#	finance_x =i-1
				#print("main.check found none----------------------", finance_y)
				#print("main.check",finance_list[0][finance_y])
			#	break





		for y_dash in range(0, dash_rows):
			#starts at 1 to skip over "Month"
			for x_ind in range(0, len(finance_list[0])):
				
				dash_month = str(wsheet_re.cell(y_dash,0))
				dash_month = dash_month.replace("text:","")
				dash_month =dash_month.replace("'","")
				ind_month = finance_list[0][x_ind]



				if dash_month == ind_month:

					print("The months are equal", ind_month)
				

					for x_dash in range(0,dash_cols):

						for y_ind in range(len(finance_list)):

							dash_name = str(wsheet_re.cell(0, x_dash))
							dash_name = dash_name.replace("text:","")
							dash_name =dash_name.replace("'","")


							ind_name = finance_list[y_ind][0]


							if dash_name == ind_name:


								print("ind name is",ind_name, finance_list[y_ind][x_ind])
								dash_sheet.write(y_dash,x_dash, finance_list[y_ind][x_ind])
				#print("   ")






		# # for loop going through the dashboard file to look where to write the value
		# for y in range(0,max_rows):
		# 	for x in range(1,max_cols): 
		# 		cell_val = wsheet_re.cell(y,x)

		# 		#if str(cell_val) == "empty:''":
		# 		sheet_y = y
		# 		sheet_x = x



		# 		month_dash = str(wsheet_re.cell(sheet_y,0))
		# 		month_dash = month_dash.replace("text:","")
		# 		month_dash = month_dash.replace("'","")

		# 		dash_name = str(wsheet_re.cell(0,sheet_x))
		# 		dash_name = dash_name.replace("text:","")
		# 		dash_name = dash_name.replace("'","")




		# 		month_finance = finance_list[0][x]
		# 		month_fin2_foo = finance_list[0][x]


		# 		if month_dash == month_fin2_foo:
		# 			foo=4
		# 			#print("-----------------------finance month", month_fin2_foo, month_dash)

		# 		else:
		# 			foo=3
		# 			#print(month_fin2_foo, month_dash,"-------------------------------")


					
		# 		#print("dashname",month_dash, month_finance)
		# 		if month_dash == month_finance:
		# 			#print("main.check months are equal",month_dash, month_finance)
		# 			print("equal", month_dash, month_finance)

		# 			for n in range(0,len(finance_list)):
		# 				finance_name = finance_list[n][0]

		# 				if finance_name == dash_name:
		# 					print("main.check names are equal", finance_name, dash_name, finance_list[n][finance_x])
		# 					dash_sheet.write(sheet_y,sheet_x, finance_list[n][finance_x])









def extract_finance_data(sheet):


	max_rows = sheet.nrows
	max_cols = sheet.ncols

	#dtype = {
	#	'names': ('data_name','data_value'),
	#	'formats': ('U40','U40','U40')}

	#finance_data = np.zeros(600,dtype)



	raw_list = [[]for i in range(max_rows)]

	month_cnt = 0

	for i in range(0,max_rows):
		raw_list[i] =  sheet.row(i)



	for i in range(0,max_rows):
		for j in range(0,max_cols):
			
			str_raw_data = str(raw_list[i][j])

			if 'text:' in str_raw_data:
				str_raw_data = str_raw_data.replace("text:","")
				string2 = str_raw_data.replace("'","")
				#print("extract", string2)
				#print("extract before", raw_list[i][j])
				raw_list[i][j] =  string2
				#print("extract after", raw_list[i][j])


			if 'number:' in str_raw_data:
				raw_num= str_raw_data.split(":")[1]
				numb = float(raw_num)
				raw_list[i][j] = numb


			if 'empty:' in str_raw_data:
				#print("I found empty")
				raw_list[i][j] = "None"


			#print("main.extract raw val is", string)

	

	#for k in range(0,10):
		#for n in range(0,10):
		#print(raw_list[k])
	
	return raw_list
	
	

if __name__ == "__main__":
	main()