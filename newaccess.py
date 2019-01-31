import pyodbc
import csv

def titles_and_values(sql):
	print("Connect to Database")
	conn_str = (
		r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'
		r'DBQ=MATERIAL_DATABASE.accdb;'
		)
	cnxn = pyodbc.connect(conn_str)
	crsr = cnxn.cursor()
	try:
		crsr.execute(sql) 
		rows = crsr.fetchall()
		titles = [x[0] for x in crsr.description]
		values = []
		for row in rows:
			values.append(list(row))
	finally:
		crsr.close()
		cnxn.close()
		print("Closed!")
	return titles , values

def access_to_csv(savename,sql):
	print("Creating Workbook")
	filename = "{0}.csv".format(savename)
	data = titles_and_values(sql)
	titles = data[0]
	values = data[1]
	with open(filename, "wb") as wb:
		ws = csv.writer(wb)
		ws.writerow(titles)
		for val in values:
			ws.writerow(val)
	print("Saved.")
	
def categories(sql):
	print("Connect to Database")
	conn_str = (
		r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'
		r'DBQ=MATERIAL_DATABASE.accdb;'
		)
	cnxn = pyodbc.connect(conn_str)
	crsr = cnxn.cursor()
	try:
		crsr.execute(sql)
		titles = [str(x[0]) for x in crsr.description]
	finally:
		crsr.close()
		cnxn.close()
		print("Closed!")
	return titles

def limits(limit_sql):
	limit_array = titles_and_values(limit_sql)[1]
	return limit_array

def macro(file_name,limit_sql,sql):
	new_arr = titles_and_values(sql)[0][1:14]
	limit_arr = limits(limit_sql)
	values = limit_arr[0]
	pairs = [(i,i+1) for i in range(1,len(values)) if i % 2 != 0]
	mix_type = values[0]
	with open(file_name,'w') as new_text:
		new_text.write("GMACRO \n")
		for idx, item in enumerate(new_arr):
			idx = pairs[idx]
			lsl , usl = round(values[idx[0]],5) , round(values[idx[1]],5)
			new_text.write("\nLayout. \n")
			new_text.write("IChart '{0}';\n".format(item))
			new_text.write("Title '{0} {1}';\n".format(mix_type, item))
			new_text.write("Reference 2 {0} {1};\n".format(lsl , usl))
			new_text.write("LABEL 'LSL= {0}' 'USL= {1}'.\n".format(lsl , usl))
			new_text.write("Endlayout. \n")
		new_text.write("\nENDMACRO")	
