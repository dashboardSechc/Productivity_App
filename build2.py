import xlrd
from tkinter import filedialog
import xlwt
import numpy as np
import collections
import string
import sum_h
import sys
import summary_format



def main():


	print("Hello! Program starting")
	#new_file = "C:/Users/jprivera/Desktop/NewVals.xlsx"
	#true_file = "C:/Users/jprivera/Desktop/TrueVals.xlsx"
	#true_prov = "C:/Users/jprivera/Documents/Productivity/Automation/excel_files/true_providers.xlsx"

	true_prov = "../excel_files/true_providers.xlsx"
	


	#commment : this is the argument section of the  code
	if len(sys.argv) > 1:
		print("main sys argv >0")

		if sys.argv[1] =="-d":
			print("running default version")
			raw_data = "../input_files/provider_productivity_april.xls"

		elif sys.argv[1] == "-f":
	 		raw_data = sys.argv[2]
	 		print("File path is",raw_data)

	else:
	 	raw_data = filedialog.askopenfilename(initialdir = "/")



	
	#READ:opening the files and getting sheet(0)
	try:
		raw_sheet = get_sheet(raw_data)

	except:
		print("the file", raw_data,"you have selected does not exist")


	true_sheet = get_sheet(true_prov)
	#new_sheet = get_sheet(new_file)
	#true_sheet = get_sheet(true_file)


	#WRITE:
	write_wb = xlwt.Workbook()


	max_rows = raw_sheet.nrows
	max_cols = raw_sheet.ncols

	true_max_rows = true_sheet.nrows



	#array to raw array
	array = [[] for i in range(max_rows)]

	dtype = {
			'names': ('department','type','name'),
			'formats': ('U40','U40','U40')}


	true_provs = np.zeros(200,dtype)


	true_provs = extract_true_prov(true_sheet)
	




	array = extract_data(raw_sheet,max_rows, max_cols)

	array_length = len(array)

	array_num=0
	for i in range(0,array_length):
		if not array[i]:
			#print("I found the None ins the array")
			array_num+=1

	


	#print("Main:array is 0 is ",array[0])
	#print("Main:array is 154 is ",array[194])

	#print("Main:max_row is",max_rows)
	#print("Main:array_num is",array_num)


	max_rows = max_rows-array_num

	print("Main:New max_row is",max_rows)

	#make_sheet_tabs(wb, check_array, wrte_array,max_rows, max_cols, true_max_rows)

	#duplicate_removal(write_wb, array, max_rows, max_cols)

	output_file = make_sheet_tabs(write_wb, true_provs, array, max_rows, max_cols, true_max_rows)
	#print(array[max_rows-1])

	#wb_file = "C:/Users/jprivera/Documents/Productivity/Automationnew_writtenwb.xls"
	#xls = xlrd.open_workbook(wb_file)

	print("Raw data is from",str(raw_data))



	#sum_h.test_print_hello()
	#summary_page.main()
	print("Hello! Program ending")
	print("           ")
	print("           ")

	summary_format.main()






#This is the function to extract the data from the true providers sheet
#That sheet contains a breakdown of all the staff and their role/type at SECHC
#examples include nurse, provider, clinicain, etc.

def extract_true_prov(sheet):
	nrows = sheet.nrows
	ncols = sheet.ncols
	dtype = {
			'names': ('department','type','name'),
			'formats': ('U40','U40','U40')}
	
	#TOdo fix to not just be static 200

	providers =  np.zeros(600,dtype)
	

	#print("tue rows is", nrows)

	# = [sheet.cell_value(c,0)for c in range(nrows)]
	depts = [sheet.cell_value(c,0)for c in range(1,nrows)]

	types = [sheet.cell_value(c,1)for c in range(1,nrows)]
	
	names = [sheet.cell_value(c,2)for c in range(1,nrows)]



	#print(names[0])


	i = 0
	for item in depts:
		#print(item)
		providers['department'][i] = item
		i+=1


	i = 0
	for item in types:
		#print(item)
		providers['type'][i] = item
		i+=1


	i = 0
	for item in names:
		#if i < 60:
		#	print(item)

		item = item.replace("'","")	
		providers['name'][i] = item
		i+=1

	#print("---------------------------",providers)
	#print("---------------------------",providers['department'][0])
	#print("---------------------------",providers[134]['name'])

	return providers





#This function extracts the data from the raw sheet and 
#massages the data into the correct formats to be written to the excel sheet.


def duplicate_removal(wb, raw_list, max_rows, max_cols):

	print("inside of duplicate removal")

	

	dupe_list = [[] for i in range(max_rows)]
	pre_sort_list = [1]*max_rows
	set_list = [1]*max_rows
	post_set_list = [1]*max_rows

	#raw_sheet = wb.add_sheet("No Duplicate data")


	sorted_cnt= 0
	for i in range(0,max_rows):
		pre_sort_list[i] = raw_list[i][1]
		



	set_list = set(pre_sort_list)


	i=0
	for item in set_list:
		post_set_list[i] = str(item)
		sorted_cnt+=1
		i+=1



	z=0
	found_name = 0
	for x in range(0,sorted_cnt):
		name_1 = post_set_list[x]
		for y in range(0,max_rows):
			name_2 = raw_list[y][1]
			if name_1 == name_2:
				#if "ROBYN" in name_1:
				dupe_list[z] = raw_list[y]
				#print("found a match",z, name_1,raw_list[x][4], name_2,)
				z+=1	
				#print("found a duplicate",name_1,name_2)

		if y == max_rows-1:
			
			print(" new  ")
			for i in dupe_list:
				print("New findings",i)
				print(" new  ")

		del dupe_list[:]




def extract_data(sheet,max_rows,max_cols):

#This section of the program makes sure that any garbage before the data is ignored
	#print("inside of extract data")

	prov_array = [[] for i in range(max_rows)]

	
	Found_department= False

	true_row_count=0
	array_count=0
	for i in range(max_rows):
		row = sheet.row(i)

		string_row = str(row[0])

		#if 'department' in row:

		if 'Department Name' in string_row:

			Found_department=True



		if Found_department ==True:

		 	prov_array[array_count] = row
		 	array_count+=1

		#prov_array[i] = row

	#print("Extract_data:max rows is ",max_rows)
	#print("Extract_data:array count is ",array_count)

	max_rows = array_count

#This is the section tat starts extracting the data into a list

	#TODO: fix the number of percentage vars
	for i in range(max_rows):
		for x in range(max_cols):
			str_prov = str(prov_array[i][x])
			

			if 'number:' in str_prov:
				prov = str_prov.split(":")[1]
				numb = float(prov)
				prov_array[i][x] = numb
				#print("true", foo2)


			if 'empty' in str_prov:
				#print(str_prov)
				prov_array[i][x]="None"


			if 'text' in str_prov:
				prov = str_prov.split(":")[1]
				prov2 = prov.replace('"', "'")
				prov2 = prov2.replace("'","")
				#print(prov2)
				prov_array[i][x] = prov2


			if '%' in str_prov:
				#this makes the number into a float
				prctg = str_prov.split(":")[1]
				prctg2 = prctg.replace("%","")
				prctg3 = prctg2.replace("'","")
				prctg4 = float(prctg3)
				prctg4 = prctg4/100
				prov_array[i][x] = prctg4
				# print(prov_array[i][x])
				# print(prctg4)
				# print(type(prov_array[i][x]))
		#print("Prov array---------------",prov_array[i])

	#print(prov_array[max_rows-1])

	return prov_array








def make_sheet_tabs(wb, check_array, wrte_array,max_rows, max_cols, true_max_rows):
	print("in make sheet tabs")
	
	dept = "'Department Name'"
	empty = "empty:''"
	report = "'Report ID"
	num = 1
	i = 0
	num_dept=0
	dtype = {
			'names': ('department','type','name'),
			'formats': ('U40','U40','U40')}
	

	true_rec_arr = np.zeros(200,dtype)

	hass_data = [[] for i in range(max_rows)]

	rawP_array = [[] for i in range(max_rows)]

	nurse_array = [[] for i in range(max_rows)]

	#obot_array = [[] for i in range(max_rows)]

	provider_array = [[] for i in range(max_rows)]

	prescriber_array  = [[] for i in range(max_rows)]

	clinician_array  = [[] for i in range(max_rows)]

	true_dept = [[] for i in range(max_rows)]

	null_array = [[] for i in range(max_rows)]

	#setting the different styling options


	#col_style = xlwt.easyxf('align:wrap on, pattern:pattern solid, fore_colour pale_blue')

	col_style_month = xlwt.easyxf('pattern:pattern solid, fore_colour pale_blue ;align: wrap on')
	col_style_ytd = xlwt.easyxf('pattern:pattern solid, fore_colour rose ;align: wrap on')


	#print(rawP_array)

	new_array = [1]*max_rows
	

	#print(max_rows)
	for i in range(0,max_rows):
		new_array[i] =  wrte_array[i][0] 
	
	dept_names = set(new_array)

	#print("AFTER THE BREAK STATEMENT!!!!!!!!!")

	#print(check_array[''])


	#TODO FIX:assuming that there are no more than 100 departments
	dept_array = [1]*100
	

	i = 0
	
	for item in dept_names:
		
		dept_array[i] = str(item)
		#print("-----------------",type(dept_array[i]))
		i = i+1
		#print(item)


	#removing any unwanted items in the dept array
	while report in dept_array:dept_array.remove(report)
	#print("1")

	while dept in dept_array:dept_array.remove(dept)
	#print("2")

	while "Department Name" in dept_array:dept_array.remove("Department Name")

	

	while "SE 85 BEHAVIORAL HEALTH" in dept_array:dept_array.remove("SE 85 BEHAVIORAL HEALTH")

	while empty in dept_array:dept_array.remove(empty)
	#print("3")

	while num in dept_array:dept_array.remove(num)
	#print("4")



	#counting the number of itmes in the dept array to be used as a counter below
	for item in dept_array:
		num_dept = num_dept+1
		#print(type(item), item)

	#print(num_dept)


	
	sorted_1 = sorted(dept_array)
	
	#for num in sorted_1:
	#	print("srted item is",num)

	#Adding in Raw sheet of data
	raw_sheet = wb.add_sheet("Raw data")
	for x in range(0,max_rows):
		for y in range(0,max_cols):
			raw_dat = wrte_array[x][y]

			#if x <30:
				#if type(raw_dat)==str :
					#raw_dat = raw_dat.replace('"', "'")
					#print(raw_dat, type(raw_dat))


			raw_sheet.write(x,y,raw_dat)

			#raw_sheet.write(x,y,raw_dat)

	#print("----------------------------------- after made raw data")
	#adding in all of the regular sheets of data

	hass_sheet = wb.add_sheet("Hass Center", cell_overwrite_ok = True)
	count_true = 0

	hass_placeholder = [1]*max_rows
	hass_names = [1]*100


	for r in range(0,max_cols):
				title = str(wrte_array[0][r])

				if 1 < r < 10:
					hass_sheet.write(0,r,title,col_style_month)

				elif 9 < r < 18: 
					hass_sheet.write(0,r,title,col_style_ytd)


				else:
					hass_sheet.write(0,r,title)


	#section to put the Hass data into one array----------------------------



	count_hass=0
	for n in range(0,max_rows):
		if "400" in str(wrte_array[n][0]):
			hass_placeholder[n] = wrte_array[n][0]
			hass_data[count_hass] = wrte_array[n]
			#print(hass_data[count_hass])
			count_hass +=1
			#print(wrte_array[n][0])		

	#print(count_hass)

	#TODO: add in a srt feature, since the list is random
	hass_set=set(hass_placeholder)

	i=0
	for item in hass_set:
		hass_names[i] = str(item)
		i+=1


	#while empty in hass_array:hass_array.remove(empty)
	#print("3")

	#print(hass_array)

	while 1 in hass_names:hass_names.remove(1)

	while "1" in hass_names:hass_names.remove("1")

	#hass_names.remove("1")


	# for i in hass_names:
	# 	print("hass names is",i)



	row_count = 0
	num_hass_data = len(hass_data)
	#print(num_hass_data)

	num_hass = len(hass_names)

	#print("   ")
	for x in range(0,num_hass):
		row_count+=2
		start = row_count
		for n in range(0, count_hass):
			#row_count = n
			dept_name = str(hass_names[x])
			data = str(hass_data[n][0])
			#print(data, dept_name)


			if dept_name == data:
				for i in range(0,max_cols):
					hass_sheet.write(row_count,i,hass_data[n][i])
				#break


				#print(hass_data[n][0])
			#print(hass_names[x])
			#print("has data is",hass_data[n])


				row_count+=1

		#print("End is ",row_count, start)
		hass_sheet.write(row_count, 0,dept_name+" total")
		#sum_h.make_summations(hass_sheet,start,row_count-1)
		#row_count==1
		#print("  ")





	#hass section array--------------------------------------------------------











	#\\-----------------------------------------------------Creating the sheets of data------------------------------------------------------------------------------//
	#THIS FOR LOOP IS CHANGED TO LESSEN THE DEPARTMENT NUMBERS
	#print(num_dept)

	#This for loop goes through all the department levels
	for i in range(0,num_dept):
	#for i in range(0,5):
		foo = str(sorted_1[i])
		sheet_name = foo.replace("'","")
		#print(sheet_name)
		#Because the Hass center is a department, and with 400 in the title it won't match up
		#if "400" in sheet_name:

			
		#TODO: FIx the overwrite feature, to not overwrite the page. instead ignore each error individually	
		#else:


		#TODO fix fatal error, when the sheet name is not changed to null, was working fine before
		if "400" in sheet_name:
						


			sheet_name="Null"
			#sheet_wt = wb.add_sheet(sheet_name, cell_overwrite_ok = True)
			pass

		else:
			sheet_wt = wb.add_sheet(sheet_name, cell_overwrite_ok = True)




		#TODO make into a function
		#For loop adding the column names to each sheet
			for r in range(0,max_cols):
				title = str(wrte_array[0][r])

				if 1 < r < 10:
					sheet_wt.write(0,r,title,col_style_month)

				elif 9 < r < 18: 
					sheet_wt.write(0,r,title,col_style_ytd)


				else:
					sheet_wt.write(0,r,title)


		#-----------------------------------------------------------------------------------


		dept_r=0
		num_hass_raw=0
		num_hass_true=0

		#-------------------------------------Creating the raw department array for 
		g=0
		for dept_r in range(1,max_rows):

			dept = str(wrte_array[dept_r][0])
			dept2 = dept.replace("'","")

			#print("dept is ----------------------------------",dept)
			#print(sheet_name, dept2)

			if sheet_name == dept2:
				#print("in if state---------------------------------------")
				for z in (0,max_cols-1):
					foo = wrte_array[dept_r]
					#print("printing foo-----------------",foo)
					rawP_array[num_hass_raw] = foo
					#print(num_hass_raw,foo)
				num_hass_raw+=1

			

		#------------------------------------Creating the true department array 


		for dept_n in range(0,true_max_rows):
			dept = str(check_array[dept_n]['department'])
			#if "400" in dept:

			#PLacing the items from the true record array if the departments are equaal to the current sheet
			if sheet_name == dept:
				#print("Foobar dept----------------",dept)
				for z in (0,max_cols-1):
					foo_true = check_array[dept_n]

					#print("foobar foo true-------------------------",foo_true)
					true_rec_arr[num_hass_true] = foo_true
					#true_dept[num_hass_true] = foo_true
				num_hass_true+=1

		#print("num of true is", num_hass_true, max_rows, max_cols)


		#print(true_rec_arr[0])
		#print(true_dept[0])
		#print("----------------------------",wrte_array[0])

		#print("-------------------------------------",rawP_array[5])
		
		#print("sheet name is----------",sheet_name)




		#if "ADULT" in sheet_name: 
		nurse_array = make_provider(true_rec_arr,rawP_array,num_hass_true,num_hass_raw,'nurse', sheet_name)
		provider_array =  make_provider(true_rec_arr,rawP_array,num_hass_true,num_hass_raw,'provider', sheet_name)
		obot_array =  make_provider(true_rec_arr,rawP_array,num_hass_true,num_hass_raw,'obot', sheet_name)
		special_array =  make_provider(true_rec_arr,rawP_array,num_hass_true,num_hass_raw,'Specialist', sheet_name) 
		clinician_array =  make_provider(true_rec_arr,rawP_array,num_hass_true,num_hass_raw,'Clinician', sheet_name)
		prescriber_array =  make_provider(true_rec_arr,rawP_array,num_hass_true,num_hass_raw,'prescriber',sheet_name)
		IBH_array =  make_provider(true_rec_arr,rawP_array,num_hass_true,num_hass_raw,'IBH',sheet_name)
		null_array = make_provider(true_rec_arr,rawP_array,num_hass_true,num_hass_raw,'null', sheet_name)
		unknwn_array = make_unknwn_array(true_rec_arr,rawP_array,num_hass_true,num_hass_raw,'null')
		#null_array = make_null_array(check_array,true_rec_arr,rawP_array,num_hass_true,num_hass_raw,unknwn_array)






		prov_count=0
		for item in provider_array:
			if item:
				#if "women" in sheet_name:
				#print("the item is",item)
				prov_count+=1


		nurse_count=0
		for item in nurse_array:
			if item:
				nurse_count+=1

		obot_count=0
		for item in obot_array:
			if item:
		 		obot_count+=1


		clin_count=0
		for item in clinician_array:
			if item:
				#print(item)
				clin_count+=1


		pres_count=0
		for item in prescriber_array:
			if item:
				#print(item)
				pres_count+=1


		ibh_count=0
		for item in IBH_array:
			if item:
				#print(item)
				ibh_count+=1

		spec_count=0
		for item in special_array:
			if item:
			#	print(item)
				spec_count+=1


		unknwn_count=0
		for item in unknwn_array:
			if item:
				#print(item)
				unknwn_count+=1

		null_count=0
		for item in null_array:
			if item:
				#print(item)
				null_count+=1


		#------------------------------Printing out the provider array-----------------------------------------



		#med_staff_write_to_sheet(max_cols,max_staff,staff_type, staff_array ,sheet_name, current_excel_sheet):


		dept_name = str(sheet_name)

		sheet_y_position =1

	#	print(sh)

		if "1601" in dept_name:
			#print("in the Adult department ---------------")



			
			sheet_y_position = med_staff_write_to_sheet(sheet_y_position, max_cols,prov_count,"provider", provider_array ,sheet_name, sheet_wt)
			sheet_y_position = med_staff_write_to_sheet(sheet_y_position, max_cols,clin_count,"clinician", clinician_array ,sheet_name, sheet_wt)
			sheet_y_position = med_staff_write_to_sheet(sheet_y_position, max_cols,nurse_count,"nurse", nurse_array ,sheet_name, sheet_wt)
			sheet_y_position = med_staff_write_to_sheet(sheet_y_position, max_cols,obot_count,"obot", obot_array ,sheet_name, sheet_wt)
			sheet_y_position = med_staff_write_to_sheet(sheet_y_position, max_cols,pres_count,"prescriber", prescriber_array ,sheet_name, sheet_wt)
			sheet_y_position = med_staff_write_to_sheet(sheet_y_position, max_cols,ibh_count,"ibh", IBH_array ,sheet_name, sheet_wt)
			sheet_y_position = med_staff_write_to_sheet(sheet_y_position, max_cols,spec_count,"specialty", special_array ,sheet_name, sheet_wt)
			sheet_y_position = med_staff_write_to_sheet(sheet_y_position, max_cols,unknwn_count,"unknown", unknwn_array ,sheet_name, sheet_wt)
			sheet_y_position = med_staff_write_to_sheet(sheet_y_position, max_cols,null_count,"null", null_array ,sheet_name, sheet_wt)

		


	#print("Report used",raw_data)
	file_name = 'final_report.xls'
	wb.save(file_name)
	#foo = wb.add_sheet("shet1")

	return(file_name)


#This is the  function that writes the individuals sections to the department sheets
#For example for a department this would write the provider, nurse, and specialty sections
def med_staff_write_to_sheet(excel_y_position, max_cols, max_staff, staff_type, staff_array, sheet_name, current_excel_sheet):

#------------------------------Writing out the staff array-----------------------------------------

	#print("Inside of the med staff to sheet function")

#	print("write_sheet", sheet_name, staff_type,max_staff)

	#print("         ")
	#print("         ")


	if max_staff ==0:
		#print(sheet_name)
		return(excel_y_position)
		
	#print(sheet_name)

	#params

	#staff_count is the number of individuals in a particluar section
	#staff_type = "Provider"

	#max_staff =0

	#max_staff2 = len(staff_array)


	#print("The max staff is", max_staff,max_staff2)

	#print("excel y position",excel_y_position)
	#print("excel y position",type(excel_y_position))


	#This is an offset from the top of the excel sheet
	if "provider" in staff_type:
		pass

	else:
		excel_y_position +=2 

	#print(staff_type)

	#if staff_type == "null":
		#pass
	#else:
	staff_type = staff_type + " total"



		#TOdo keep track of current position in excel file
	#excel_y_position=0
		
	for staff_count in range(0,max_staff):

		#excel_y_position +=1
		for j in range(0,max_cols):
			dept_name=staff_array[staff_count][0].replace("'","") 
			#print("staff name is",staff_array[staff_count][1])

			#print("Sheet name and prov deot are",sheet_name, prov_deot1)
			if  "400" in sheet_name:



				pass

			if dept_name == sheet_name:
				#print("sheet name true",provider_array[cnt][0]) 
				cell_value = staff_array[staff_count][j]
				#print(cell_value)

				current_excel_sheet.write(excel_y_position,j,cell_value)
				#print("written to sheet is",cell_value)
				#print(cell_value)
				if staff_count == max_staff-1:
					current_excel_sheet.write(excel_y_position+1,0,staff_type)
					#num =-2
					#current_excel_sheet.write(+1,0,staff_type)
					#make_summations(current_excel_sheet,num,excel_y_position)
 					#print("Total cell value is", cell_value)

		if dept_name == sheet_name:

			excel_y_position +=1



	#print("Leaving the med staff to sheet report")

	return(excel_y_position)




def make_summations(sheet_wt,cnt1,cnt2):


	style_col = xlwt.easyxf('pattern:pattern solid, fore_colour aqua',num_format_str='#,##0.00')

	for x in range(2,6):
		lettr = string.ascii_uppercase[x]
		sheet_wt.write(cnt2+1,x,xlwt.Formula("SUM(%s%g:%s%g)" % (lettr,cnt1 + 4,lettr,cnt2+1)))

	sheet_wt.write(cnt2+1,6,xlwt.Formula("%s%g/(%s%g + %s%g)" % ('E',cnt2+2,'E',cnt2+2,'D',cnt2+2)))
	sheet_wt.write(cnt2+1,7,xlwt.Formula("%s%g/%s%g)" % ('D',cnt2+2,'C',cnt2+2)))
	sheet_wt.write(cnt2+1,8,xlwt.Formula("%s%g/%s%g)" % ('F',cnt2+2,'C',cnt2+2)))
	sheet_wt.write(cnt2+1,9,xlwt.Formula("%s%g/%s%g)" % ('F',cnt2+2,'D',cnt2+2)))

	for x in range(10,14):
		lettr = string.ascii_uppercase[x]
		sheet_wt.write(cnt2+1,x,xlwt.Formula("SUM(%s%g:%s%g)" % (lettr,cnt1 + 4,lettr,cnt2+1)))

	sheet_wt.write(cnt2+1,14,xlwt.Formula("%s%g/(%s%g + %s%g)" % ('M',cnt2+2,'M',cnt2+2,'L',cnt2+2)))
	sheet_wt.write(cnt2+1,15,xlwt.Formula("%s%g/%s%g)" % ('L',cnt2+2,'K',cnt2+2)))
	sheet_wt.write(cnt2+1,16,xlwt.Formula("%s%g/%s%g)" % ('N',cnt2+2,'K',cnt2+2)))
	sheet_wt.write(cnt2+1,17,xlwt.Formula("%s%g/%s%g)" % ('N',cnt2+2,'L',cnt2+2)))


def make_unknwn_array(true_rec_arr,rawP_array,num_hass_true,num_hass_raw,type_prov):

	#bool_type =check_equal(true_rec_arr[cnt]['type'],type_prov)

	# print("OOOOOOOOOOO----------------------------------inside of make unknwn funciton")
	# print("    ")
	# print("    ")
	# print("    ")
	#print("hi")
	cnt3 =0
	array = [[] for i in range(200)]
	array2 = [[] for i in range(200)]

	#print("True record array is-------------",true_rec_arr['department'][2])

	# for i in range(0,num_hass_raw):
	# 	print(rawP_array[i])

	# for i in range(0,num_hass_true):
	# 	print(true_rec_arr[i])

	

	if num_hass_true ==0:
	# 	#num_hass_true=1

	 	for n in range(0,num_hass_raw):
	 		array[n] = rawP_array[n]
	 		#print(array2[n])


	# 	return(rawP_array)


	#print("before for loop", num_hass_raw, num_hass_true)
	for cnt in range(0,num_hass_raw):

	 	for cnt2 in range(0,num_hass_true):
	 		
	 		true_name =  true_rec_arr[cnt2]['name']
	 		true_dept = true_rec_arr[cnt2]['department']

	 		#print("Hello, inside for loop")
	 		#print("The true department is --------------------------",true_dept)


 			raw_string = rawP_array[cnt][1].replace("'","")

	 		if raw_string == true_name:
	 			break
	 		else:
	 			
	 			if cnt2 ==  num_hass_true-1:
	 				#print(raw_string)
	 				#print(rawP_array[cnt])
	 				array[cnt3] = rawP_array[cnt]
	 				cnt3+=1


	#print("Array is....................",array)
	
	# print("    ")
	# print("    ")
	# print("     ")
	# print("CCCCCCC---------------------------------leaving unknwn function")

	return(array)


# def make_null_array(true_rec_arr,rawP_array,num_hass_true,num_hass_raw,unknwn_array):


# 	#
# 	return array


def make_provider(true_rec_arr,rawP_array,num_hass_true,num_hass_raw,type_prov, sheet_name):

	#print("inside of make provider funciton",type_prov,sheet_name)
	#print("num hass true", num_hass_true , "num_hass raw", num_hass_raw)
	#print("     ")
	#print("   ")

	array = [[] for i in range(600)]

	cnt=0
	cnt3=0

	#print(rawP_array[0][0])
	#for a in range(0,num_hass_raw):
		#print("raw array is",rawP_array[a])


	#print(num_hass_raw,"-----------------",type_prov)
	
	#print(num_hass_true)

	for cnt in range(0,num_hass_true):
		match = 0


		#if sheet_name == 

		bool_type =check_equal(true_rec_arr[cnt]['type'],type_prov)

		#print(bool_type, true_rec_arr[cnt]['type'], type_prov)


		if bool_type is True:
			#print("--------true found nurse", check_array[cnt])
			#print(num_hass_raw)
			for cnt2 in range(0,num_hass_raw):
				#print(cnt2)
				true_string =  true_rec_arr[cnt]['name']
				#raw_string = rawP_array[cnt2][1].replace("'","")
				raw_string = rawP_array[cnt2][1]
				#print("raw string is ----------",raw_string, true_string)

				if raw_string == true_string:					

					#dept_name = str(rawP_array[cnt2][0])
					#if "FAMILY" in dept_name:
					#	print()

					match = 1
					break;
				#else:
					#print("The non matching names", raw_string, true_string)

			if match ==1:
				newarray = rawP_array[cnt2]
				array[cnt3] = newarray
				cnt3+=1				
				#print("True--------------------------", raw_string, true_string)

			
	#print("leaving function")
	#print("    ")
	#print("  ")
	return array




#Fucntion to count the number of providers in a sheet

def provider_loc(sheet_p):
	#this assumes that the coumn names start at (0,0)
	#placing the 
	nrows = sheet_p.nrows
	ncols = sheet_p.ncols

	col_names = [sheet_p.cell_value(0,c)for c in range(ncols)]

	for x in range(0,ncols):
		#print(col_names[x])
		#TODO make it able to handle different cases
		if col_names[x] == "Provider":
			#print(sheet_p.cell_value(x,true_prov_col))
			#print("true")
			true_prov_col = x

	ret_list = [None]* 2
	ret_list[0] = true_prov_col


	prov_count =0

	for x in range(1,nrows):
		if sheet_p.cell_value(x,true_prov_col) != '':
			#print(tsheet.cell_value(x,true_prov_col))
			prov_count+=1
		else:
			break

	ret_list[1] = prov_count
	return ret_list
	



def get_sheet(file_loc):
	#TODO check if file exists

	workbook = xlrd.open_workbook(file_loc)
	sheet = workbook.sheet_by_index(0)

	return sheet


def check_equal(a,b):
	try:
		return a.upper() == b.upper()

	except AttributeError:
		return a== b




if __name__ == "__main__":
	main()