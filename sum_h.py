
import string
import xlwt




def test_print_hello():

	print(" i am in the sum_h.py file")	




def remove_duplicates(raw_array):
	print("inside of duplicate removal")









def department_names(max_rows, wrte_array):

	dept = "'Department Name'"

	dept_xf = "'Department Name' (XF"
	empty = "empty:''"
	report = "'Report ID"
	num = 1
	num_dept = 0



	new_array = [1]*max_rows
	

	#print(max_rows)
	for i in range(0,max_rows):
		if "400" in str(wrte_array[i][0]):
			pass
		else:

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
	dept_array[i] = "Hass center"


	#removing any unwanted items in the dept array
	while report in dept_array:dept_array.remove(report)
	#print("1")

	while dept in dept_array:dept_array.remove(dept)
	#print("2")

	while dept_xf in dept_array:dept_array.remove(dept_xf)

	while empty in dept_array:dept_array.remove(empty)
	#print("3")

	while num in dept_array:dept_array.remove(num)
	#print("4")

	while "1" in dept_array:dept_array.remove("1")


	#counting the number of itmes in the dept array to be used as a counter below
	for item in dept_array:
		num_dept = num_dept+1
		#print("sum_h.department_names",type(item), item)

	#print(num_dept)
	
	sorted_1 = sorted(dept_array)

	i = 0
	for n in sorted_1:
		n = n.replace("(XF","")
		sorted_1[i] = n
		i+=1
		#print("inside of sum_h", )
	#sorted_1[i] = "Hass center"
	#print("sum_h.dept_sum i is",i)



	return sorted_1




def find_null(read_sheet,verf_array,v_count):
	#print("inside of find null")


	


	sheet_rows = read_sheet.nrows
	null_flag=0
	#v_count =0

	
	for n in range(0,sheet_rows):
		#cell_val = 
		if "null" in str(read_sheet.cell(n,0)):

			null_val = "'" +str(read_sheet.name)+ "'" + "!$D" + str(n+1)


			verf_array['department'][v_count] = read_sheet.name
			verf_array['null_val'][v_count] = null_val


			#print("sum_h.find_null ",read_sheet.name, null_val)
			null_flag=1
	if null_flag==0:
		verf_array['department'][v_count] = read_sheet.name
		verf_array['null_val'][v_count] = "0"
		
		#print("foobar-------------------------------------sum_h.find_null no Null value", read_sheet.name)

	return verf_array




#Function to write the verification out to excel
def dept_summ2(v_sheet, dept_verif, length_of_verif):
	#print("inside of dept 2", length_of_verif, dept_verif)

	style2 = xlwt.easyxf('pattern:pattern solid, fore_colour green',num_format_str='#,##0.00')

	n=0
	i=0
	empty_flag=0


	#for i in range(0,length_of_verif):
		#print(" dept_sum2 value is",dept_verif[i]['raw_sum'], type(dept_verif[i]['raw_sum']),i)
		#print(" dept_sum2 value is", dept_verif[i]['department'],i)

		##if not dept_verif[i]['raw_sum']:
			#print("--------------------------------dept_summ i is", i)



	#length_of_verif = len(dept_verif)

	v_sheet.write(2,1,"Department Name")
	v_sheet.write(2,2,"Null values")
	v_sheet.write(2,3,"Deptartment total")
	v_sheet.write(2,4,"Raw total")
	

	for n in range(0,length_of_verif-1):
		for i in range(0,3):
			#print(dept)

				#break
			#if not dept_verif[i]:
				#break


			#TODO check for all 3 to be none
			if not dept_verif[i]['raw_sum']:
					#print("--------------------------------dept_summ i is", i)
					dept_verif[i]['raw_sum'] = 0
					dept_verif[i]['null_val']=0



			if empty_flag ==1:
				pass

			else:

				null_value = str(dept_verif['null_val'][n])
				summ_value = str(dept_verif['summ_val'][n])
				raw_value = str(dept_verif['raw_sum'][n])



				#cell_value = "'Raw data'!G4:G9"
				#print("dept_sum2",cell_value)
				#print("sum.summ2 depts are",dept_verif['department'][n])

				cell_dept_value = dept_verif['department'][n]
				v_sheet.write(n+3,1,cell_dept_value,style2)

				
				#print("sum_h.summdept2", null_value)

				#try:
				#print("sum_h.dept_sum", null_value)
				v_sheet.write(n+3,2,xlwt.Formula("SUM(%s)"%null_value))
				v_sheet.write(n+3,3,xlwt.Formula("SUM(%s)"%summ_value))
				v_sheet.write(n+3,4,xlwt.Formula("SUM(%s)"%raw_value))
				#except:
					#pass
				#write(xlwt.formula(dept_verif['null_val']))
				#write(xlwt.formula(dept_verif['summ_val']))



	#v_sheet.write(y_position+4,2,xlwt.Formula(sum_var))






def department_summation(v_sheet,start, end, raw_array, y_position,name_of_sheet, dept_name):


	#print("inside of department summation")
	#print("The start and end are", start, end, y_position)




	sum_var = "SUM(" + "'" + name_of_sheet +"'"+"!" +"D" + str(start) + ":" + "D"+ str(end) +")"
	#print("sum_h.department_summation",sum_var, dept_name)

	#for i in range(0,3):

	#v_sheet.write(y_position+4, 1, dept_name)
	#v_sheet.write(y_position+4,2,xlwt.Formula(sum_var))







def make_summations(wb_sheet,excel_start,excel_end,start_1, end_1,start_2, end_2,nurse_m,nurse_y, nurse_bool):


	#adding in the format values for adding color to the sheet

	#format_col = wb_book.add_format()
	#format_col.set_bg_color('green')

	#xlwt.add_palette_colour("custom_col",0x0B )


	style = xlwt.easyxf('pattern:pattern solid, fore_colour pale_blue',num_format_str='#,##0.00')
	style2 = xlwt.easyxf('pattern:pattern solid, fore_colour green',num_format_str='#,##0.00')



	print("sum_h.make_summ", nurse_y, nurse_m)


	for x in range(start_1,end_1):
		#TODO change from ascii to the top of an excel sheet
		#print("Inside the make summations", excel_start, excel_end)
		lettr = string.ascii_uppercase[x]
		if nurse_bool == True and x == start_1+1:
			wb_sheet.write(excel_end+1,x,xlwt.Formula("SUM(%s%g:%s%g,%s)" % (lettr,excel_start+1,lettr,excel_end+1,nurse_m)), style2)
		else:
			wb_sheet.write(excel_end+1,x,xlwt.Formula("SUM(%s%g:%s%g)" % (lettr,excel_start+1,lettr,excel_end+1)), style)




	wb_sheet.write(excel_end+1,end_1,xlwt.Formula("%s%g/(%s%g + %s%g)" % ('E',excel_end+2,'E',excel_end+2,'D',excel_end+2)),style)
	wb_sheet.write(excel_end+1,end_1+1,xlwt.Formula("%s%g/%s%g)" % ('D',excel_end+2,'C',excel_end+2)),style)
	wb_sheet.write(excel_end+1,end_1+2,xlwt.Formula("%s%g/%s%g)" % ('F',excel_end+2,'C',excel_end+2)),style)
	wb_sheet.write(excel_end+1,end_1+3,xlwt.Formula("%s%g/%s%g)" % ('F',excel_end+2,'D',excel_end+2)),style)


	for x in range(start_2,end_2):
		lettr = string.ascii_uppercase[x]
		if nurse_bool == True and x == start_2 +1:
			#wb_sheet.write(excel_end+1,x,xlwt.Formula("SUM(%s%g:%s%g)" % (lettr,excel_start + 1,lettr,excel_end+1)),style2)
			wb_sheet.write(excel_end+1,x,xlwt.Formula("SUM(%s%g:%s%g,%s)" % (lettr,excel_start+1,lettr,excel_end+1,nurse_y)), style2)
		else:
			wb_sheet.write(excel_end+1,x,xlwt.Formula("SUM(%s%g:%s%g)" % (lettr,excel_start + 1,lettr,excel_end+1)),style)


	wb_sheet.write(excel_end+1,end_2,xlwt.Formula("%s%g/(%s%g + %s%g)" % ('M',excel_end+2,'M',excel_end+2,'L',excel_end+2)),style)
	wb_sheet.write(excel_end+1,end_2+1,xlwt.Formula("%s%g/%s%g)" % ('L',excel_end+2,'K',excel_end+2)),style)
	wb_sheet.write(excel_end+1,end_2+2,xlwt.Formula("%s%g/%s%g)" % ('N',excel_end+2,'K',excel_end+2)),style)
	wb_sheet.write(excel_end+1,end_2+3,xlwt.Formula("%s%g/%s%g)" % ('N',excel_end+2,'L',excel_end+2)),style)

