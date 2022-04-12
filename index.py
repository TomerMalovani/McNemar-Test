from asyncio.windows_events import NULL
import tkinter as tk
from attr import define
# import pandas for reading excel sheet
import pandas as pd
import re
from tkinter import filedialog
import numpy as np
from datetime import datetime
from tkinter import messagebox
# import mcnemar from statsmodels
from statsmodels.stats.contingency_tables import mcnemar
# try and catch
try:
	root = tk.Tk()
	root.withdraw()
	log_file_path=NULL
	csv_file_path=NULL
	file_path = filedialog.askopenfilename()
	if(file_path == ''):
		raise Exception("No file was selected")
	# copy file_path to a new variable as root directory

	root_dir = file_path
	# get the root dir from the file path
	root_dir = root_dir.split('/')
	# remove the last element of the list
	filename=root_dir.pop().split('.')[0]

	# join the list back to a string
	root_dir = '/'.join(root_dir)
	log_file_path = "%s/%s.log" % (root_dir,filename)
	csv_file_path = "%s/%s.csv" % (root_dir,filename)
	if(file_path.split(".")[-1] != "xlsx"):
		raise Exception("bad file ,Please select an excel file")
	df = pd.read_excel(file_path)
	
	users = []
	count=0
	
	if((len(df.columns)-1) % 2 != 0):
		raise Exception("bad file format, there should be 2 columns for each user")


	for i in range(1, len(df.columns), 2):
		new_name=re.sub("[^0-9]", "", df.columns[i])
		users.append({"code":new_name,"11":0,"10":0,"01":0,"00":0})
		for j in range(0, len(df.index)):
			first = df.iloc[j][i]
			second=df.iloc[j][i+1]
			if pd.isna(first):
				raise Exception("bad file format, row %s column %s is empty" % (j,i))
			if pd.isna(second):
				raise Exception("bad file format, row %s column %s is empty" % (j,i+1))
			if first == 1 and second==1:
				users[count]["11"] +=1
			elif first == 1 and second==0:
				users[count]["10"] += 1
			elif first == 0 and second==1:
				users[count]["01"] += 1
			elif first == 0 and second==0: 
				users[count]["00"] += 1
			else:
				raise Exception("bad file format, row %s is bad, all values should be 0 or 1" % (j+2))
		users[count]["mcnemar"]=mcnemar([[users[count]["00"],users[count]["01"]],[users[count]["10"],users[count]["11"]]], exact=True).pvalue

		count+=1

	final_data=[]
	# loop over users
	for user in users:
		user_stats={
			 "user_code":user["code"],
			 "A_succsess_rate":round(((user["10"]+user["11"])/(len(df.index)))* 100, 1),
			 "B_succsess_rate":round(((user["01"]+user["11"])/(len(df.index)))* 100, 1),
			 "mcnemar_pvalue":user["mcnemar"]
		}
		final_data.append(user_stats)


	df = pd.DataFrame.from_dict(final_data)
	# create a new csv file with the new data
	print(csv_file_path)

	df.to_csv(csv_file_path,index = False, header=True)
	# write to log file
	with open(log_file_path, "a+") as log:
		# loop over users with index
		now = datetime.now()
		date_time = now.strftime("\n%m/%d/%Y, %H:%M:%S")
		log.write("results::%s:\n" % (date_time))
		
		for i, user in enumerate(users):
			log.write(str(i)+"."+str(user)+"\n")
		log.write("csv file was created at:%s\n" % (csv_file_path))
	




	messagebox.showinfo('completed', "csv and log files was created at:\n%s\n" % (root_dir))

# catch error
# check if variable is defined



except Exception as e:
	if  log_file_path is NULL:
		if file_path == ' ':
			messagebox.showwarning('alert title', 'error log was created where script was activated')
		log_file_path = "mcnemar_error_log.log"
	with open(log_file_path, "a+") as log:
		# loop over users with index
		now = datetime.now()
		date_time = now.strftime("\n%m/%d/%Y, %H:%M:%S")
		log.write("ERROR::"+date_time+":"+str(e)+"\n")
	messagebox.showwarning('ERROR', str(e))







