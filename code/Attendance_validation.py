import pandas as pd
import difflib
import csv


#https://stackoverflow.com/questions/41192424/python-how-to-correct-misspelled-names
def is_similar(first, second, ratio):
    return difflib.SequenceMatcher(None, first, second).ratio() > ratio

#S1 = 'Amar Singh Kosta'
#S2 = 'AMAR KOSTA'
#R = is_similar(S1,S2.lower(), 0.5)
#print(R)


# read by default 1st sheet of an excel file
OSRTC_df = pd.read_excel('../data/customer/Dec Attendence.xlsx')
INTAG_df = pd.read_excel('../data/attendance/Driver Attendance Sheet-01-Dec-31-Dec.xlsx')

#OS_LIST = ['1','2','3','4','5','6','7','8','9',10,'11','12','13','14','15','16','17','18','19',20,'21']
#IN_LIST = ['01-Dec','02-Dec','03-Dec','04-Dec','05-Dec','06-Dec','07-Dec','08-Dec','09-Dec','10-Dec','11-Dec','12-Dec','13-Dec','14-Dec','15-Dec','16-Dec','17-Dec','18-Dec','19-Dec','20-Dec','21-Dec']
OS_LIST = [10,'11','12','13','14','15','16','17','18','19',20,'21']
IN_LIST = ['10-Dec','11-Dec','12-Dec','13-Dec','14-Dec','15-Dec','16-Dec','17-Dec','18-Dec','19-Dec','20-Dec','21-Dec']


#print(OSRTC_df)
#print(OSRTC_df[['Name of SO', 'Total']])
#print(INTAG_df[['Driver Name', 'Total Working Days']])

with open('../results/Result.csv', 'w', newline='') as file:

	writer = csv.writer(file)
	field = ["OSRTC Name", "Intangle Name", "Accuracy", "Match Days", "Miss Match Days"]
	writer.writerow(field)

	for os_drv_cnt in range(0,len(OSRTC_df['Name of SO'])):
		#print(OSRTC_df['Name of SO'][os_drv_cnt])
		NAME = OSRTC_df['Name of SO'][os_drv_cnt]
		print('---------------------------------')

		FLAG = 0 

		for in_drv_cnt in range(0, len(INTAG_df['Driver Name'])):
			OSRTC_S = NAME
			INTAG_S = INTAG_df['Driver Name'][in_drv_cnt]
			R = is_similar(OSRTC_S.lower(),INTAG_S.lower(), 0.9)
			DAY_ACC = 0
			if(R):
				print('OSRTC ENTRY----->',OSRTC_S)
				#print(OSRTC_df['Total'][os_drv_cnt])
				print('INTANGLE ENTRY----->', INTAG_S)
				#print(INTAG_df['Total Working Days'][in_drv_cnt])
				
				#print(INTAG_df['01-Dec'][in_drv_cnt])
				for day_cnt in range(0,len(OS_LIST)):
					OS_A = OSRTC_df[OS_LIST[day_cnt]][os_drv_cnt]
					#print(OS_A)
					IN_A = INTAG_df[IN_LIST[day_cnt]][in_drv_cnt]
					#print(float(IN_A[0:2]))
					if((OS_A == 1) and (float(IN_A[0:2]) > 0.5)):
						DAY_ACC = DAY_ACC + 1
						#print('MATCH')
					if((OS_A == 'WO') and (float(IN_A[0:2]) == 0)):
						DAY_ACC = DAY_ACC + 1
						#print('MATCH1')
					if((OS_A == 'A') and (float(IN_A[0:2]) == 0)):
						DAY_ACC = DAY_ACC + 1
						#print('MATCH2')
					if((OS_A == 'L') and (float(IN_A[0:2]) == 0)):
						DAY_ACC = DAY_ACC + 1
						#print('MATCH3')
				#exit(0)
				FLAG = 1
				print('ACCURACY------------------->', (DAY_ACC/len(OS_LIST))*100)
				print('Number Of match days------->', DAY_ACC)
				print('Number Of miss-match days-->', (len(OS_LIST) - DAY_ACC))
				ROW = [OSRTC_S, INTAG_S , (DAY_ACC/len(OS_LIST))*100 , DAY_ACC , (len(OS_LIST) - DAY_ACC)]
				writer.writerow(ROW)


		if(FLAG < 1):
			print('OSRTC ENTRY----->',OSRTC_S)
			print('NO INTANGLE ENTRY FOUND FOR DRIVER')
			ROW = [OSRTC_S, 'NA' , 'NA' , 'NA' , 'NA']
			writer.writerow(ROW)
	


		

