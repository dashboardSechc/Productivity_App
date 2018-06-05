import xlrd
import xlwt
import numpy as np
from xlutils.copy import copy
import sum_h
import xlsxwriter
from xlutils.filter import process, XLRDReader, XLWTWriter



def main():

	print("Starting the summary program foobar!")

	exisiting_prod_path = "final_report.xls"

	rb = xlrd.open_workbook(exisiting_prod_path,formatting_info=True)

	col_style_month = xlwt.easyxf('pattern:pattern solid, fore_colour pale_blue ;align: wrap on')
	col_style_ytd = xlwt.easyxf('pattern:pattern solid, fore_colour rose ;align: wrap on')
	style_col = xlwt.easyxf('pattern:pattern solid, fore_colour pale_blue',num_format_str='#,##0.00')
	style2 = xlwt.easyxf('pattern:pattern solid, fore_colour red',num_format_str='#,##0.00')

	#r_sheet = rb.sheet_by_index(3)

	wb_book = copy(rb)
	empty = "empty:''"

	dtype = {
			'names': ('department','raw_sum','null_val','summ_val'),
			'formats': ('U40','U40','U40','U40')}
	

	dept_verification = np.zeros(600,dtype)


	

	style = xlwt.easyxf(num_format_str='#,##0.00')

	#ws_sheet = wb.get_sheet(3)
	#hass_sheet = wb.get_sheet(1)

	#try to remove the overwrite feature
	try:
		ws_sheet = wb_book.add_sheet('Summary_sheet',cell_overwrite_ok=True)
		verification_sheet = wb_book.add_sheet('verification is now',cell_overwrite_ok=True)
	except:
		pass

	first_sheet = rb.sheet_by_index(0)
	
	#places the raw data from the 1st shet into an array
	#purpose is to count the number of departments
	raw_array = extract_data(first_sheet, first_sheet.nrows, first_sheet.ncols)
	
	y_position=0
	dept_count =0
	dept_sum ="null"


	#print(" summ.main length", len(raw_array[0]))
	raw_len = len(raw_array[0])

	#this section is to write the column names to the file, and skipping over the provider name
	skip = False
	for i in range(0,raw_len-1):

		raw_col_name = raw_array[0][i].replace("(XF","")
		
		if i ==1:
			skip =True
			#print("I found 1", skip)

		if skip == True:
			#print("-----------------skip is true")
			raw_col_name = raw_array[0][i+1].replace("(XF","")
			#print(raw_col_name)

		#print("I found 1", skip)


		
		#print("summ.main name",raw_col_name)


		if 0<i<9: 
			ws_sheet.write(0,i,raw_col_name, col_style_month) 

		elif 8<i<17:
			ws_sheet.write(0,i,raw_col_name, col_style_ytd)

		else: 
			ws_sheet.write(0,i,raw_col_name)









	#TODO: make into a function

	sheet_count=0
	v_count=0
	for sht in rb.sheets():
		max_rows = sht.nrows
		max_cols = sht.ncols
		#cell = sht.row_values(max_rows)
		
		#writing out the department names to the sheet

		#TODO:this should not be in an except statement
		#This threw an error before throwing a that it couldn't find ws
		#try:
		#print("summ.main",sht.name)
		
		#TODO potential error with this before ws_sheet.write below
		wb_sheet = wb_book.get_sheet(sheet_count)



		#if "Raw" in str(sht.name) or "CARE COORDINATION" in str(sht.name):


		if "CARE COORDINATION" in str(sht.name) or "Raw" in str(sht.name):
		#if "Raw" in str(sht.name):
			pass
			#print("summar raw sheet", sht.name)

		else:
			#print("summary normal", sht.name)

			ws_sheet.write(y_position,0,sht.name)


		if "Raw" not in str(sht.name):


			dept_verification = sum_h.find_null(sht,dept_verification, v_count)

			dept_sum = make_dept_sum(wb_sheet, sht)

			dept_verification['summ_val'][v_count] = dept_sum

			v_count+=1


		
			#if not dept_verification[k]:
			#	pass


		auto_summations(wb_sheet,sht)

		#first_val = sht.cell(0,0)	

		sheet_arr=extract_data(sht,max_rows, max_cols)
		while empty in sheet_arr:sheet_arr.remove(empty)
		#foobar_sheet = wb.add_sheet("summary page_foo")


		#For the number of rows in a particular department sheet
		y_excel_position=0
		#for cell in sheet_arr:

		len_sheet = len(sheet_arr)



		#Start of for loop creating the department section in the summary sheet-----------------------------
		total_count =0
		for x in range(0,len_sheet):

			#print("cell is",cell[0], "        sheet_arr is", sheet_arr[5][0]   )
			#print("the cell is",cell[0])
			#print("sheet_arr is",sheet_arr[y_excel_position][0])
			#if cell[0] =="'Provider total'":

			#var1 = cell[0]
			cell_string = sheet_arr[y_excel_position][0]

			

			if type(cell_string)==xlrd.sheet.Cell :
						#print("------------------------------------------ value is None",cell_value)
						cell_string="None"

			
			#instead of searching total of each section(i.e. nursing, providers) I just search for the keyword total to grab all of them
			#Only for the summary sheet
			if "total" in cell_string and "null" not in cell_string:
				if "CARE COORDINATION" in str(sht.name):
					print("summ.mainI found care coord")
					
				else:

					total_count +=1
					y_position+=1
					skip =False

					for x_position in range(0, max_cols-1): 
											
						#should be indexed at 0, so should be -1
						col_letter = xlsxwriter.utility.xl_col_to_name(x_position)
						#should not need to be increased by 1
						link = "'" + str(sht.name) + "'" + "!$" + col_letter + "$" + str(y_excel_position+1)
						

						if x_position ==1:
							skip =True
							#print("I found 1", skip)

						if skip == True:
							#print("-----------------skip is true")
		
							col_letter = xlsxwriter.utility.xl_col_to_name(x_position+1)
							#should not need to be increased by 1
							link = "'" + str(sht.name) + "'" + "!$" + col_letter + "$" + str(y_excel_position+1)
							

						ws_sheet.write(y_position,x_position,xlwt.Formula(link), style)

						x_position+=1





			dept_name_formatted = str(sht.name).replace("SE1601","")

			dept_name_formatted = dept_name_formatted.lower()
			#print("The formatted name is-----", dept_name_formatted)
			dept_wrt_name = dept_name_formatted.title() + " total"

			
			#this final section is when eveything is being written to the file
			y_excel_position+=1
			if(y_excel_position == len_sheet):
				if "Raw" in dept_wrt_name or "Care Coordination" in dept_wrt_name:
					pass
				else:
					#print("summ_main", dept_wrt_name)
					ws_sheet.write(y_position+1,0,dept_wrt_name, style2)
					start_position =y_position - total_count
					sum_h.make_summations(ws_sheet,start_position,y_position,1,5,9,13,0,0, False)
					#print("You've reached the end", y_position, start_position)


 		# End of for loop per sheet ------------------------------------------------------------------
	




		#adding 2 spaces between the department sections in the sheet
		y_position+=3
	
		
		#print("max rows is",max_rows)

		if max_rows ==0:
			print("its time to break")
			#break
			pass



		#print("I am here")
		#print(first_val)
		#dept_array[dept_count] = sht.name
		dept_count+=1
		sheet_count+=1
		




	sorted_dept_list = sum_h.department_names(first_sheet.nrows, raw_array)


	dept_verification = verification(first_sheet,verification_sheet, raw_array, sorted_dept_list, first_sheet.nrows, dept_verification)





	g=0
	for i in range(0,len(dept_verification)):
		#print("summar.main dept rec array", dept_verification[i])
		g+=1
		if not dept_verification[i]:
			break


	

	sum_h.dept_summ2(verification_sheet, dept_verification,g)


	try:
		wb_book.save('correct_format_sheet.xls')
	except:
		print("ERROR: You currently have the output file open, please close it and re run the program")
		print("Execution terminated")
		exit(0)


	print("Output is, correct_format_sheet.xls")
	print("Summary Ending ")
	#print("        ")
	#print("    ")


#Function that creates the summation of sums
def make_dept_sum(wb_sheet, rb):

	sheet_rows=rb.nrows
	#print(" in the make dept sum function", rb.name)

	#print("sheet rows is", sheet_rows)
	style2 = xlwt.easyxf('pattern:pattern solid, fore_colour red',num_format_str='#,##0.00')

	dept_name = str(rb.name)

	sum_of_totals = "SUM("
	sum_of_totals_2 = "SUM("
	sum_of_totals_3 = "SUM("
	sum_of_totals_4 = "SUM("

	sum_of_ytd = "SUM("
	sum_of_ytd_2 = "SUM(" 
	sum_of_ytd_3 = "SUM("
	sum_of_ytd_4 = "SUM("


	sum_test="SUM(A1:A2)"

	if "raw" in dept_name:
		pass
	else:

		for i in range(0,sheet_rows):
			if "total" in str(rb.cell(i,0)):
				#print(rb.cell(i,0), i)
				sum_of_totals = sum_of_totals + "C" + str(i+1) + ","
				sum_of_totals_2 = sum_of_totals_2 + "D" + str(i+1) + ","
				sum_of_totals_3 = sum_of_totals_3 + "E" + str(i+1) + ","
				sum_of_totals_4 = sum_of_totals_4 + "F" + str(i+1) + ","

				sum_of_ytd = sum_of_ytd + "K" + str(i+1) + ","
				sum_of_ytd_2 = sum_of_ytd_2 + "L" + str(i+1) + ","
				sum_of_ytd_3 = sum_of_ytd_3 + "M" + str(i+1) + ","
				sum_of_ytd_4 = sum_of_ytd_4 + "N" + str(i+1) + ","



				#print("sum 4", sum_of_totals_4)

			#total_count +=1
	sum_of_totals = sum_of_totals + ")"
	sum_of_totals_2 = sum_of_totals_2 + ")"
	sum_of_totals_3 = sum_of_totals_3 + ")"
	sum_of_totals_4 = sum_of_totals_4 + ")"

	sum_of_ytd = sum_of_ytd + ")"
	sum_of_ytd_2 = sum_of_ytd_2 + ")"
	sum_of_ytd_3 = sum_of_ytd_3 + ")"
	sum_of_ytd_4 = sum_of_ytd_4 + ")"

	#TODO probably should not do a try except here


	wb_sheet.write(i+5,0,rb.name)

	try:
		wb_sheet.write(i+5,2, xlwt.Formula(sum_of_totals))
		wb_sheet.write(i+5,3, xlwt.Formula(sum_of_totals_2))
		link_location = i+6
		wb_sheet.write(i+5,4, xlwt.Formula(sum_of_totals_3))
		wb_sheet.write(i+5,5, xlwt.Formula(sum_of_totals_4))


		wb_sheet.write(i+5,10, xlwt.Formula(sum_of_ytd))
		wb_sheet.write(i+5,11, xlwt.Formula(sum_of_ytd_2))
		wb_sheet.write(i+5,12, xlwt.Formula(sum_of_ytd_3))
		wb_sheet.write(i+5,13, xlwt.Formula(sum_of_ytd_4))


	except:
		pass

	excel_end = i+5

	wb_sheet.write(excel_end,6,xlwt.Formula("%s%g/(%s%g + %s%g)" % ('E',excel_end+1,'E',excel_end+1,'D',excel_end+1)))
	wb_sheet.write(excel_end,7,xlwt.Formula("%s%g/%s%g)" % ('D',excel_end+1,'C',excel_end+1)))
	wb_sheet.write(excel_end,8,xlwt.Formula("%s%g/%s%g)" % ('F',excel_end+1,'C',excel_end+1)))
	wb_sheet.write(excel_end,9,xlwt.Formula("%s%g/%s%g)" % ('F',excel_end+1,'D',excel_end+1)))

	wb_sheet.write(excel_end,14,xlwt.Formula("%s%g/(%s%g + %s%g)" % ('M',excel_end+1,'M',excel_end+1,'L',excel_end+1)))
	wb_sheet.write(excel_end,15,xlwt.Formula("%s%g/%s%g)" % ('L',excel_end+1,'K',excel_end+1)))
	wb_sheet.write(excel_end,16,xlwt.Formula("%s%g/%s%g)" % ('N',excel_end+1,'K',excel_end+1)))
	wb_sheet.write(excel_end,17,xlwt.Formula("%s%g/%s%g)" % ('N',excel_end+1,'L',excel_end+1)))


	linked_sum = 0
	try:
		linked_sum ="'" + str(rb.name) + "'" + "!$D" + str(link_location)
		#print("summary.make_dept_sum", linked_sum)
	except:
		pass

	
	return linked_sum

	#print("summar_format.make_dept_sum Sum of totals",sum_of_totals)






def extract_data(sheet,max_rows,max_cols):

	prov_array = [[] for i in range(max_rows)]

	for i in range(max_rows):
		row = sheet.row(i)
		prov_array[i] = row


	#TODO: fix the number of percentage vars
	for i in range(max_rows):
		for x in range(max_cols):
			str_prov = str(prov_array[i][x])
			if 'number:' in str_prov:
				prov = str_prov.split(":")[1]
				#numb = float(prov)
				numb = prov
				prov_array[i][x] = numb
				#print("true", foo2)
			#print(str_prov)


			if 'text' in str_prov:
				prov = str_prov.split(":")[1]
				prov_array[i][x] = prov
			if '%' in str_prov:
				prctg = str_prov.split(":")[1]
				prctg2 = prctg.replace("%","")
				prctg3 = prctg2.replace("'","")
				print(prctg3)
				prctg4 = float(prctg3)
				prctg4 = prctg4/100
				prov_array[i][x] = prctg4
				# print(prov_array[i][x])
				# print(prctg4)
				# print(type(prov_array[i][x]))
		#print("Prov array---------------",prov_array[i])

	#print(prov_array[max_rows-1])

	return prov_array


def verification(raw_sheet,verf_sheet, raw_array, sorted_list, max_rows, dept_verification):

	print("Inside of verification")




	v_list_count=0
	for y in range(1,len(dept_verification)):
		#print("summary.verficatoin dept array",y)

		v_list_count+=1

		if not dept_verification[y]:
			break
		

	v_count =0
	#for i in sorted_list:
	#	print(i)

	sorted_num = len(sorted_list)
	current_dept = "not been set"


	#for loop to change all of the 400s into the hass center
	for i in range(0,max_rows):
		if "400" in str(raw_array[i][0]):
			raw_array[i][0] = "Hass center"

			#print("found 400")


	#high level for loop that goes through all the sorted departments
	for x in range(0,sorted_num):
		#print("summar_format.verfication sorted list",sorted_list[x])
		start = 0
		end =0

		#loop going through the maximum rows within the raw data sheet
		for n in range(0,max_rows):

			raw_str = raw_array[n][0].replace("(XF","")
			sorted_str = sorted_list[x]
			#print("verification current dept",current_dept)




			if raw_str == sorted_str and start ==0:
				
				start =1
				start_excel = n
				#current_dept = raw_str.replace("'","")
				current_dept = raw_str
				#print("summar.verf",start_excel,current_dept)
				
			
			if current_dept != raw_str and start ==1:
				
				#print("They are not equal", current_dept, raw_str, n, x)

				end_excel = n
				
				#print("summary_format.verf  current_dept 1",current_dept)
				sum_var = "'" + raw_sheet.name +"'"+"!" +"D" + str(start_excel+1) + ":" + "D"+ str(end_excel) 
				#print("summary.verf sum var", sum_var)

				current_dept = current_dept.replace("'","")

				for k in range(0,v_list_count):

					verif_str = str(dept_verification['department'][k])


					verif_str =  verif_str.replace(" ","")

					current_dept= current_dept.replace(" ","")

					if verif_str == current_dept:
						#print("summary.verification 	Found a match", current_dept,verif_str, sum_var)
						#dept_verification['department'][v_count] = current_dept


						#print("summary.verf",sum_var)
						dept_verification['raw_sum'][k] = sum_var
						#print("summary",dept_verification[v_count])
					

				#print("called verf")
				break
				
				#print("no longer equal", raw_str, end_excel)


	return dept_verification








def auto_summations(wb_sheet,rb_sheet):


	#print("     ")
	#Make sure their names are the same? Don't know when it wouldn't be the case
	#print("Entering the aut_summations function", wb_sheet.name, rb_sheet.name)


	max_rows = rb_sheet.nrows
	max_cols = rb_sheet.ncols
	nurse_visits_m = 0
	nurse_visits_y = 0

	total_count = 0

	#print("Max colls is:",max_rows)
	#print(rb_sheet.name)

	#if "SE1601" in str(rb_sheet.name):

		#loop to iterate over the columns in a rb_sheet, counting the number of staff total sectoins
	for x in range(0,max_rows):
		#print(rb_sheet.cell(x,1))

		if "NURSE" in str(rb_sheet.cell(x,1)):
			#print("summ.auto",rb_sheet.cell(x,1))
			#print("summ.auto",rb_sheet.cell(x,3))

			count_str = str(x+1)

			nurse_visits_m ='D' + count_str

			nurse_visits_y = 'L' + count_str


		
		if "total" in str(rb_sheet.cell(x,0)):
			total_count +=1
			#print("I found the total-----------------------------------------", total_count)





	nurse_sec = False
	start_position = 1
	for n in range(0,total_count):		

		#print("The start position",start_position)
		start =0
		for i in range(start_position, max_rows):
			cell_type = rb_sheet.cell_type(i,0)
			#ifel statement to find the start of each section
			if cell_type is xlrd.XL_CELL_EMPTY:
				pass

			elif start ==0:
				sum_start = i
				start=1
				#print(sum_start,i,rb_sheet.cell(i,0), rb_sheet.cell(i,1))


			if "total" in str(rb_sheet.cell(i,0)):

				if "nurse" in str(rb_sheet.cell(i,0)):
					nurse_sec = True
					#print("auto_summ" ,nurse_sec)



				end_position = i+1
				#og_start_position = start_position +1
				start_position = end_position
				#call summation function
			

				sum_end = i-1

				#print("--------------------")
				#print("The start", sum_start)
				#print("the position is", rb_sheet.cell(sum_start,1))

				#print("The end is", rb_sheet.cell(sum_end,1))
				#print("The end positions are", sum_end)
				#print("-------------------")
				sum_h.make_summations(wb_sheet,sum_start,sum_end,2,6,10,14, nurse_visits_m, nurse_visits_y, nurse_sec)
				nurse_sec = False

				break



	#print("Leaving the auto summations function ")
	#print("    ")










def get_sheet(file_loc):
	#TODO check if file exists

	workbook = xlrd.open_workbook(file_loc)
	sheet = workbook.sheet_by_index(0)

	return sheet






if __name__ == "__main__":
	main()
