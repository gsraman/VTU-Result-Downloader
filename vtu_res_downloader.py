#!/usr/bin/python

import re
import urllib2
import xlwt

clg_name = raw_input("Enter College Name ( BMSIT ) : ");
clg_code = raw_input("Enter College Code ( 1BY12 ) : ");
branch = raw_input("Enter Branch ( CS ) : ");
strength = raw_input("Enter Branch Strength : ");
path = raw_input("Enter Path To Save Result(/home/suraj/Desktop) : ");

workbook = xlwt.Workbook()
sheet = workbook.add_sheet(clg_name+"_"+branch)
style = xlwt.easyxf('font: bold 1')
sheet.write(0,0,'NAME',style)
sheet.write(0,1,'USN',style)
sheet.write(0,2,'TOTAL MARKS',style)
sheet.write(0,3,'RESULT',style)

#---------------- PARSER ---------------------------------------------------------
def parser(usn):
	url = "http://results.vtu.ac.in/vitavi.php?rid="+usn+"&submit=SUBMIT"
	data_read = urllib2.urlopen(url)
	data = data_read.read()
	rxp1 = re.findall(r'<table>(.*?)</table>',data,re.DOTALL)
	if len(rxp1) == 0:
			print 'No Result!!'
			sheet.write(i,0,'NILL')
			sheet.write(i,1,usn)
			sheet.write(i,2,'NILL')
			sheet.write(i,3,'NILL')
			workbook.save(path+'/'+clg_name+'_'+branch+'.xls')
	else:
			rxp2 = re.findall(r'<B>(.*?)</B>',data,re.DOTALL)
			rxp3 = re.findall(r'<td>(.*?)</td>',rxp1[2],re.DOTALL)
			rxp4 = re.findall(r'<b>(.*?)</b>',rxp1[0],re.DOTALL)
			res_lst = rxp4[2].split(';')
			res_lst_1 = res_lst[2].split(' ')
			nm = rxp2[0].split('(')
			name = nm[0]
			if res_lst_1[0] == 'FAIL':
				print name,'\t',0
				sheet.write(i,0,name)
				sheet.write(i,1,usn)
				sheet.write(i,2,0)
				sheet.write(i,3,'FAIL')
				workbook.save(path+'/'+clg_name+'_'+branch+'.xls')
			elif len(rxp3) == 0:
				print name,'\t','-'
				sheet.write(i,0,name)
				sheet.write(i,1,usn)
				sheet.write(i,2,'NILL')
				sheet.write(i,3,'PASS')
				workbook.save(path+'/'+clg_name+'_'+branch+'.xls')
			else:
				tm = rxp3[3].split(' ')
				total_marks = int(tm[1])
				print name,'\t',total_marks
				sheet.write(i,0,name)
				sheet.write(i,1,usn)
				sheet.write(i,2,total_marks)
				sheet.write(i,3,'PASS')
				workbook.save(path+'/'+clg_name+'_'+branch+'.xls')
#---------------- PARSER ----------------------------------------------------------

i = 1;

while  i <= int(strength) :
	if  i <= 9 :
		code = clg_code+branch+"00";
		usn = code + str(i);
		parser(usn);
	elif  i <= 99 :
		code = clg_code+branch+"0";
		usn = code + str(i);
		parser(usn);
	else:
		code = clg_code+branch;
		usn = code + str(i);
		parser(usn);
	i = i + 1;
workbook.save(path+'/'+clg_name+'_'+branch+'.xls')
