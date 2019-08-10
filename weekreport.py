#!/usr/bin/python2.7
#coding:utf-8
import xlrd
import os
import time
import datetime
import mysql.connector as mc

def listDir(rootDir):
	for filelist in os.listdir(rootDir):
		if(filelist == '.DS_Store'):
			continue
		path = os.path.join(rootDir,filelist)
		if os.path.isdir(path):
			listDir(path)
		elif(os.path.exists(path)):
			print(path)
			process_task_file(path)

def process_task_file(path):
	data = xlrd.open_workbook(path)
	version = 0
	
	if(len(data.sheet_names()) > 1):
		version = 1;
	
	if (version == 0 ):
		table = data.sheets()[0]
	else:
		table = data.sheet_by_name('丁智君'.decode('utf-8'))
	
	
	if(version == 0):
		for i in range(2,table.nrows):
			if(table.row_values(i)[0] == '遗留问题'.decode('utf-8')):
				break
			task_date = table.row_values(i)[1]
			task_name = table.row_values(i)[2]
			task_desc = table.row_values(i)[3]
			task_desc_list = task_desc.split('\n')
			task_name_list = task_name.split('\n')
			if(len(task_desc_list) != len(task_name_list)):
				print("Error:%s."%(path))
				break
			for j in range(0,len(task_desc_list)):
				sql = 'insert into t_tasks values('+'str_to_date('+'\''+task_date+'\','+'\'%Y.%m.%d\'),'+'\''+task_name_list[j]+'\','+'\''+task_desc_list[j]+'\','+'\'\',\'\');'
				cursor.execute(sql)
	elif(version == 1):
		for i in range(2,table.nrows):
			if(table.row_values(i)[1] == ''):
				break
			task_date = xlrd.xldate_as_datetime(table.row_values(i)[1],0)
			task_name = table.row_values(i)[2]
			task_desc = table.row_values(i)[7]
			task_start = datetime.date.today()
			task_end = datetime.date.today()
			try:
				t = xlrd.xldate.xldate_as_tuple(table.row_values(i)[5],0)
				p = xlrd.xldate.xldate_as_tuple(table.row_values(i)[6],0)
				task_start = datetime.time(t[3],t[4],t[5])
				task_end = datetime.time(p[3],p[4],p[5])
			except:
				pass
			sql = 'insert into t_tasks values('+'str_to_date('+'\''+task_date.strftime('%Y-%m-%d')+'\',' \
			+'\'%Y-%m-%d\'),'+'\''+task_name+'\',\''+task_desc+'\','+'str_to_date('+'\''+task_start.strftime('%T') \
			+'\','+'\'%T\'),'+'str_to_date('+'\''+task_end.strftime('%T') + '\','+'\'%T\')'+');'
			#print(sql)
			cursor.execute(sql)
	db.commit()
	data.release_resources()
		
db = mc.connect(user='root',password='doremi',database='moondb',use_unicode=True)
cursor = db.cursor()
listDir('/Volumes/Lion/ProjectDocuments/北明/工作周报/recent')
db.close()
