import unicodedata
import pandas as pd
import re
source = pd.read_excel("data/Canada_Hosp1_COVID_InpatientData.xlsx", header = 0,  dtype=str)
source2 = pd.read_excel("data/Canada_Hosp2_COVID_InpatientData.xlsx", header = 0,  dtype=str)
data = []
count = 0

def normalize_col (ele, col_src):
	
	for i in col_src.split(','):
		temp = ""
		i = re.sub("^\\s", "", i)
		for l in re.split('\\s', i):
			if (l!='other'):
						
				temp+='_'+l
		ele.append(re.sub("^_", "", temp))

for x in source.values:
	try:
		if (x[1].lower().find('covid') != -1):
			x[8] = x[8].lower().replace("\\", "")
			x[8] = x[8].replace("\"", "")
			x[8] = re.sub("\\]|\\[", "", x[8])
			ele = []

			normalize_col(ele, x[8])
			normalize_col(ele, x[9])
			normalize_col(ele, x[13])


			

			
			ele.append('covid')
			ele_cov =""
			for i in ele:
				if (i!=""):
					ele_cov+=","+i



			data.append([count, re.sub('^[,\\s]',"",ele_cov)])
	except Exception as e:
		pass
	count+=1

for x in source2.values:
	
	try:
		if (x[1].lower().find('covid') != -1):
			x[8] = x[8].lower().replace("\\", "")
			x[8] = x[8].replace("\"", "")
			x[8] = re.sub("\\]|\\[", "", x[8])

			ele = []
			normalize_col(ele, x[8])
			normalize_col(ele, x[9])
			normalize_col(ele, x[13])

			
			ele.append('covid')
			ele_cov =""
			for i in ele:
				if (i!=""):
					ele_cov+=","+i

			data.append([count, re.sub('^[,\\s]',"",ele_cov)])
	except Exception as e:
		pass
		
	count+=1
	


export_excel = pd.DataFrame(
        data,
        columns=[
        'id',
        'mediacation_history'])
       

export_excel.to_excel('hey.xlsx', 'SOL', index = False)