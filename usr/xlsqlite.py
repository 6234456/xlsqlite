#!/src/bin/env python
# -*- coding: utf-8 -*-


# #####
#  @param  db:    name of the sqlite database
#  @param  tbl:   name of the table in the given db
#  @param  wb:    name of the workbook from which to fetch the data, default to be ".xls" format. no extension needed
# #####

def xl2sql(wb = None, sht = None, db = "src.db", tbl = "src"):
  # xlrd is needed to handle the excel io
	from xlrd import open_workbook
	import os
	import sqlite3
	os.chdir("\\".join(str(__file__).split("\\")[:-1]))
	
	if not wb:
		wb = "src"
	
	book = open_workbook(wb + ".xls", encoding_override="utf-8")
	
	if not sht:
		sheet = book.sheet_by_index(0)
	else:
		sheet = book.sheet_by_name(sht)
		
	conn = sqlite3.connect(db)
	cursor = conn.cursor();
	
	# check the data type of the cells, store in the dict col => (callback, name)
	types = {}
	callbacks = {}
	
	sql_create_table = "CREATE TABLE [" + tbl + "] ( "
	
	for c in range(sheet.ncols):
		callbacks[c], types[c] = type_mapping(sheet.cell(1,c).ctype)
		if types[c]:
			sql_create_table = sql_create_table + " [" + sheet.cell(0,c).value +"] " + types[c] + ","
		
	sql_create_table = sql_create_table[:-1] + " );"

	cursor.execute(sql_create_table)

	# insert the records
	sql_insert_value = "INSERT INTO [" + tbl + "] VALUES ( "
	for r in range(1, sheet.nrows):
		for c in range(sheet.ncols):
			if types[c]:
				if types[c] == "TEXT" or types[c] == "DATETEXT":
					sql_insert_value = sql_insert_value + callbacks[c](sheet.cell(r,c).value).decode(encoding='UTF-8',errors='strict') + ","
				else:
					sql_insert_value = sql_insert_value + str(callbacks[c](sheet.cell(r,c).value)) + ","
				
		sql_insert_value = sql_insert_value[:-1] + ");"
		
		try:
			cursor.execute(sql_insert_value)
		except sqlite3.OperationalError:
			print sql_insert_value
			return

		sql_insert_value = "INSERT INTO [" + tbl + "] VALUES ( "
	
	conn.commit()
	conn.close()
			
def xldate2str(d):
	from xlrd import xldate_as_tuple
	from datetime import date

	a =  xldate_as_tuple(d,0) 
	return date(a[0],a[1],a[2]).strftime("'%Y-%m-%d'")
	
def sqlstr(s):
	try:
		res =  "'" + str(s).replace("'","''") + "'"
	except UnicodeEncodeError:
		res = "'" + s.encode(encoding='UTF-8',errors='strict').replace("'","''") + "'"
	return res
		
def type_mapping(t):
	if t == 1:
		return (sqlstr, "TEXT")
	elif t == 2:
		return (float, "REAL")
	elif t == 3:
		return (xldate2str,"DATETEXT")
	elif t == 4:
		return (int,"INTEGER")
	else:
		return (None,None)
		
# #####
#  @param  db:    name of the sqlite database
#  @param  tbl:   name of the table in the given db
#  @param  query: SQL to be executed
#  @param  wb:    name of the workbook to store the query result, default to be ".xls" format
# #####
def sql2xl(db = "src.db", tbl = "src", query = None, wb = None):
	
	import os
	import sqlite3
	import time
	from xlwt import Workbook

	os.chdir("\\".join(str(__file__).split("\\")[:-1]))

	if not wb:
		wb = time.strftime("%d%m%Y", time.localtime())
		
	sql = sqlite3.connect(db)
	cursor = sql.cursor()
	w = Workbook(encoding = "utf-8")
	sht = w.add_sheet(tbl)
	r = 1
	c = 0
	cnt = 1
		
	if not query:
		query = "SELECT * FROM ["+ tbl + "];" 

	for i in cursor.execute(query):
		for j in i:
			sht.write(r,c,j)
			c = c + 1
		r = r + 1
		c = 0

	for i in cursor.description:
		sht.write(0, c, i[0])
		c = c + 1


	sql.close()

	w.save(wb + ".xls")
	print "DONE!"
	
		

if __name__ == "__main__":
	# xl2sql()
	sql2xl(query = "SELECT * FROM src WHERE ")
