#! /usr/bin/env python3
#
# A python script to extract articles from a .docx file and write these to a database
# articles are considered to start with a level 2 heading and go on until the next article starts.
#
from docx import Document
from glob import glob
import pymysql
conn = pymysql.connect(host='127.0.0.1', user='user', passwd='passwd', db='database', port=3307,charset='utf8')
cur = conn.cursor()
#cur.execute("SELECT * FROM artcls")

#
# file names were similar to
# 123-OCR.docx
#

filelist = glob("*.docx")
for file in filelist:
	print(file)
	filename, stuff =  file.split("-")
	document = Document(file)
	
	lstart = list()
	for paragraph in document.paragraphs:
		if paragraph.style.name=='Heading 2':
			lstart.append(list())
			lstart[-1].append(paragraph.text)
		elif paragraph.style.name=='Heading 1':
			pass
		else:
			if len(lstart) > 0:
				lstart[-1].append(paragraph.text)

	jlist = list()
	for artcl in lstart:
		jlist.append("\n".join(artcl))


## This puts everything in a database: filename, an arbitrary item id, the text and the heading
	for i in range(0,len(lstart)):
#		print(lstart[i][0])
		itemid = filename + "-"+str(i)
		cur.execute('''INSERT INTO artcls(volume,itemid,text,heading) VALUES (%s,%s, %s,%s)''',(filename,itemid,jlist[i],lstart[i][0]))

conn.commit()
cur.close()
conn.close()
