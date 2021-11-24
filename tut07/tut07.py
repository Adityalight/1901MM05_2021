import csv            # importing csv library
import os             # importing os library
os.system("cls")
 
if os.path.isfile(r'C:\Users\hp\Dropbox\PC\Documents\GitHub\1901MM05_2021\tut07\course_feedback_remaining.xlsx'):
	print("course_feedback_remaining.xlsx already exists.")                #checking if the course_feedback_remaining.xlsx file already exists
else:
	stud_info={}                                                           #dictionary to store required info of students
	header=[]                                                              #header files to store the header list from each excel file
	header1=[]
	header2=[]                                                             
	header3=[]
	with open('studentinfo.csv','r') as file:                              #reading studentinfo.csv file
		reader = csv.reader(file)
		counter=0
		for row in reader:                                                 #iterating through each row of studentinfo.csv file
			if counter==0:
				header=row                                                 #header list stores the 1st row of studentinfo.csv
				counter=counter+1
				continue
			if row[1] not in stud_info:                                    #checking if current roll no. exists in the std_info dictionary
				stud_info[row[1]]=[row[0], row[8], row[9], row[10]]        #if current roll no. doesn't exist then appending name, email, aemail and contact info corresponding to current roll no. as key in stud_info dictionary

			
	course_ltp_dict={}
	with open("course_master_dont_open_in_excel.csv","r")as file:          #reading course_master_dont_open_in_excel.csv file
		reader=csv.reader(file)
		counter=0
		for row in reader:                                                 #iterating through each row of course_master_dont_open_in_excel.csv file
			if counter==0:
				header1=row                                                #header list stores the 1st row of course_master_dont_open_in_excel.csv
				counter=counter+1
				continue
			if row[0] not in course_ltp_dict:                              #checking for current subno. in course_ltp_dict
				non_zero_bits=0                                            #initializing variable for counting the total no. of lecture , tut and practical
				ltp=row[2].split('-')                                      #storing the credits for lecture tutorial and practical in a list
				for x in ltp:
					if x != '0':                                           #checking for non zero credits 
						non_zero_bits+=1                                   #if non zero credit is present then increasing the count by 1
				course_ltp_dict[row[0]]=non_zero_bits                      #finally storing the total no. of feedback count corresponding to current subno. as key in course_ltp_dict

	course_registered={}
	with open("course_registered_by_all_students.csv","r")as file:         #reading course_registered_by_all_students.csv file
		reader=csv.reader(file)
		counter=0
		for row in reader:                                                 #iterating through each row of course_registered_by_all_students.csv file
			if counter==0:
				header2=row                                                #header list stores the 1st row of course_registered_by_all_students.csv
				counter=counter+1
				continue
			if row[0] not in course_registered:                            #checking for current roll no. in course_registered dictionary
				course_registered[row[0]]=[]                               
				course_registered[row[0]].append(row[1:4])                 #if current roll no. is not present then appending regidter_sem, schedule_sem and subno info corresponding to current roll no. as key in course_registerd dictionary
			else:
				course_registered[row[0]].append(row[1:4])

	already_submitted_feedback={}
	with open("course_feedback_submitted_by_students.csv","r")as file:     #reading course_feedback_submitted_by_students.csv file
		reader=csv.reader(file)
		counter=0
		for row in reader:                                                 #iterating through each row of course_feedback_submitted_by_students.csv file
			if counter==0:
				header3=row                                                #header list stores the 1st row of course_feedback_submitted_by_students.csv
				counter=counter+1
				continue
			if row[3] not in already_submitted_feedback:                   #checking for current roll no. in already_submitted_feedback dictionary
				already_submitted_feedback[row[3]]={}                      #if current roll no. is not present then creating dictionary corresponding to current roll no.
				if row[4] not in already_submitted_feedback[row[3]]:       #checking for subno in already_submitted_feedback[current roll no.] dictionary
					already_submitted_feedback[row[3]][row[4]]=1           #if subno is not present then storing 1 as value corresponding to current roll no. as key in above dictionary
				else:
					already_submitted_feedback[row[3]][row[4]]+=1          
			else:
				if row[4] not in already_submitted_feedback[row[3]]:
					already_submitted_feedback[row[3]][row[4]]=1
				else:
					already_submitted_feedback[row[3]][row[4]]+=1

	feedback_to_be_filled=[]
	heading=['rollno','register_sem','schedule_sem','subno','Name','email','aemail','contact']
	feedback_to_be_filled.append(heading)                                                                #appending the header in feedback_to_be_filled list
	for key,value in course_registered.items():                                                          #iterating through course_registered dictionary
		if key in already_submitted_feedback:                                                            #checking if the current key also corresponds to same key in already_submitted_feedback dictionary
			for x in value:
				if x[2] in already_submitted_feedback[key]:                                              #further checking if current subno. is present in already_submitted_feedback dictionary
					if already_submitted_feedback[key][x[2]]<course_ltp_dict[x[2]]:                      #here checking that wheather the no. of already submitted feedbacks is less than the total no. of feedbacks to be submitted
						if key in stud_info:                                                             #if the above condition id true then checking for same current roll no. key in stud_info dictionary
							lst = [key,x[0],x[1],x[2]]                                                   #appending 'rollno','register_sem','schedule_sem','subno','Name','email','aemail','contact' details in a list
							lst.extend(stud_info[key])
							feedback_to_be_filled.append(lst)                                            #finally appending the above list with required info in feedback_to_be_filled list
						else:
							feedback_to_be_filled.append(['NA','NA','NA','NA','NA','NA','NA','NA'])      #if the current roll no. key is not present in stud_info dictionary then returing NA
				else:
					if key in stud_info:
						lst = [key,x[0],x[1],x[2]]
						lst.extend(stud_info[key])
						feedback_to_be_filled.append(lst)
					else:
						feedback_to_be_filled.append(['NA','NA','NA','NA','NA','NA','NA','NA'])
		else:
			for x in value:
				if key in stud_info:
					lst = [key,x[0],x[1],x[2]]
					lst.extend(stud_info[key])
					feedback_to_be_filled.append(lst)
			else:
				feedback_to_be_filled.append(['NA','NA','NA','NA','NA','NA','NA','NA'])

	from openpyxl import Workbook                  #importing workbook from openpyxl library
	wb = Workbook()                                #creating workbook
	sheet = wb.active                              #making sheet active
	for i in feedback_to_be_filled:
		sheet.append(i)                            #appending info from feedback_to_be_filled into the excel sheet
	wb.save('course_feedback_remaining.xlsx')      #saving the xlsx file





