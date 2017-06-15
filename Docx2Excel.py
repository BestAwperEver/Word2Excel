# -*- coding: utf-8 -*-
import re
import unicodedata
import os
import sys
#import pythoncom
#import fnmatch
#from subprocess import Popen, PIPE
#import textract
#import win32com.client
import docx2txt
from openpyxl import Workbook

files = os.listdir("./docx")

wb = Workbook()
ws = wb.active
ws.append([u"ФИО", u"Телефон", "E-mail", u"Последнее место работы", u"Последняя должность", u"Год получения высшего образования", u"Резюме обновлено"])

for file in files:
	file = str(file)
	FILENAME = file.split('.')
	if len(FILENAME) < 2 or FILENAME[1] != "docx":
		continue
	FILENAME = FILENAME[0]
	PATH = os.getcwd() + "\\docx\\" + file
	print(file)

	#wordapp = win32com.client.gencache.EnsureDispatch("Word.Application")

	#try:
	#	for path, dirs, files in os.walk(sys.argv[1]):
	#		for doc in [os.path.abspath(os.path.join(path, filename)) for
	#		filename in files if fnmatch.fnmatch(filename, '*.doc')]:
	#			print("processing %s" % doc)
	#			wordapp.Documents.Open(doc)
	#			docastxt = doc.rstrip('doc') + 'txt'
	#			wordapp.ActiveDocument.SaveAs(docastxt,
	#			FileFormat=win32com.client.constants.wdFormatTextLineBreaks)
	#			wordapp.ActiveWindow.Close()
	#finally:
	#	wordapp.Quit()

	#def doc_to_text_catdoc(filename):
	#	#p = Popen('catdoc -w "%s"' % filename, shell=True, close_fds=True)
	#	p = Popen('catdoc -w "%s"' % filename, shell=True, bufsize=-1,
	#		  stdin=PIPE, stdout=PIPE, stderr=PIPE)
	#	(fi, fo, fe) = (p.stdin, p.stdout, p.stderr)
	#	fi.close()
	#	retval = fo.read()
	#	erroroutput = fe.read()
	#	fo.close()
	#	fe.close()
	#	if not erroroutput:
	#		return retval
	#	else:
	#		raise OSError("Executing the command caused an error: %s" %
	#		erroroutput)

	#text = doc_to_text_catdoc(PATH)

	#doc = win32com.client.GetObject(PATH)
	#text = str(unicodedata.normalize('NFKD',
	#str(doc.Range().Text))).replace('\r','\n').replace('\x01','').replace('\x07','').replace('\x08','')

	text = docx2txt.process("docx\\" + FILENAME + ".docx").replace('\r','\n').replace('\x01','').replace('\x07','').replace('\x08','')#.replace('ё','е').replace('Ё','Е').replace(u'\u0306','')
	text = unicodedata.normalize('NFKD', text)
	ind = 0

	fio = '?'
	number = '?'
	email = '?'
	company = '?'
	position = '?'
	graduation_year = '?'
	resume_updated = '?'

	#text = textract.process("Bushnev_Yuri_Alekseevich.doc")

	fio_match = re.search(r'([\w\u0306]+[\n \t]?){2,3}', text)
	if fio_match != None:
		fio = str(fio_match.group(0)).replace('\n','').replace('\t','')
	number_match = re.search(r'((8|\+7)[\- ]?)?(\(?\d{3}\)?[\- ]?)?[\d\- ]{7,10}', text)
	if number_match != None:
		number = number_match.group(0)
	email_match = re.search(r'[\n\t ].{2,30}@\w+\.([a-zA-z]{1,20}){1,2}[^\w]', text)
	if email_match != None:
		email = email_match.group(0)[1:-1]

	exp = re.search(r'(Опыт работы)|(Work experience)', text)
	if exp != None:
		ind = exp.span()[1]

		time = re.compile(r'(\d+ ((года?)|(лет))( \d+ месяц(а|(ев)))?)|(\d+ месяц(а|(ев)))')
		time_eng = re.compile(r'(\d+ (years?)( \d+ months?)?)|(\d+ months?)')

		#time_match = time.search(text, time.search(text, ind).span()[1])
		#ind = time_match.span()[1]

		time_match_eng = time_eng.search(text, ind)
		time_match = time.search(text, ind)

		if time_match_eng != None:
			ind = time_match_eng.span()[1]
			time_match_eng = time_eng.search(text, ind)
			if time_match_eng != None:
				ind = time_match_eng.span()[1]
		elif time_match != None:
			ind = time_match.span()[1]
			time_match = time.search(text, ind)
			if time_match != None:
				ind = time_match.span()[1]
		else:
			ind = 0
		

		#company_text = text[ind:ind+400]
		#job = text[ind+40:ind+180]
		#yearmonth = re.search(r'\d+ ((года?)|(лет)) \d+ месяц(а|(ев))', text)
		#company = re.search(r'\n.+\n',
		#text[ind+yearmonth.span()[1]:ind+yearmonth.span()[1]+50])

		line = re.compile(r'\n.+\n')
		empty_line = re.compile(r'\n(.+)?\n')
		asdf = text[ind:ind+20]

		if ind != 0:
			company_match = line.search(text, ind)
			company = company_match.group(0)[1:-1]
			ind = company_match.span()[1]
			ind = empty_line.search(text, ind)
			ind = ind.span()[1]
			position = line.search(text, ind).group(0)[1:-1]

	education_match = re.search(r'(Высшее)|(Higher)|(Магистр)', text)
	year_pattern = re.compile(r'[0-9]{4,4}')
	if education_match != None:
		ind = education_match.span()[1]
		graduation_year = year_pattern.search(text, ind).group(0)

	resume = re.search(r'(Резюме обновлено)|(Resume updated)', text)
	if resume != None:
		ind = resume.span()[1]
		date = re.compile(r'[0-9]{1,2} \w+ [0-9]{4,4}')
		date_match = date.search(text, ind)
		if date_match != None:
			resume_updated = date_match.group(0)

	#fio = fio.replace("Й","И").replace("й","и").replace(u'\u0306','')
	#company = company.replace("Й","И").replace("й","и").replace(u'\u0306','')
	#position =
	#position.replace("Й","И").replace("й","и").replace(u'\u0306','')

	#print(fio)
	#print(number)
	#print(email)
	#print(company)
	#print(position)
	#print(graduation_year)
	#print(resume_updated + "\n\n")

	ws.append([fio, number, email, company, position, graduation_year, resume_updated])
	wb.save("list.xlsx")