from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pdfminer.layout import LAParams
from pdfminer.converter import TextConverter
from io import StringIO
from pdfminer.pdfpage import PDFPage
from openpyxl import Workbook

import re
# -*- coding: UTF-8 -*- 


	
def get_pdf_file_content(path_to_pdf):
	
	'''
	path_to_pdf: is the parameter that will give access to the PDF File 
	we want to extract the content.
	'''
	
	'''
	PDFResourceManager is used to store shared resources such as fonts or images that 
	we might encounter in the files. 
	'''
	
	resource_manager = PDFResourceManager(caching=True)
	
	'''
	create a string object that will contain the final text the representation of the pdf. 
	'''
	out_text = StringIO()
	
	'''
	UTF-8 is one of the most commonly used encodings, and Python often defaults to using it.
	In our case, we are going to specify in order to avoid some encoding errors.
	'''
	codec = 'utf-8'
	
	"""
	LAParams is the object containing the Layout parameters with a certain default value. 
	"""
	laParams = LAParams()
	
	'''
	Create a TextConverter Object, taking :
	- ressource_manager,
	- out_text 
	- layout parameters.
	'''
	text_converter = TextConverter(resource_manager, out_text, laparams=laParams)
	fp = open(path_to_pdf, 'rb')
	
	'''
	Create a PDF interpreter object taking: 
	- ressource_manager 
	- text_converter
	'''
	interpreter = PDFPageInterpreter(resource_manager, text_converter)

	'''
	We are going to process the content of each page of the original PDF File
	'''
	for page in PDFPage.get_pages(fp, pagenos=set(), maxpages=0, password="", caching=True, check_extractable=True):
		interpreter.process_page(page)


	'''
	Retrieve the entire contents of the “file” at any time 
	before the StringIO object’s close() method is called.
	'''
	text = out_text.getvalue()

	'''
	Closing all the ressources we previously opened
	'''
	fp.close()
	text_converter.close()
	out_text.close()
	
	'''
	Return the final variable containing all the text of the PDF
	'''
	return text

new_re = re.compile("(?<=￥)\d+.\d+|(?<=¥)\d+.\d+|(?<=￥ )\d+.\d+|(?<=¥ )\d+.\d+")

path_to_pdf = "/Users/ZhangPeng/Desktop/电子发票/2022年2月/2yueMerged_merged.pdf"

pdftext = get_pdf_file_content(path_to_pdf).rstrip('')

newpdf = pdftext.split('')

# def storeToExcel(count,Num,total):
# 	wb = Workbook() #创建文件对象
# 	ws = wb.active  #获取默认sheet
# 	b = '第' + str(c) + '张'
# 	ws.append([b,'is','your','!'])
# 	wb.save("sample1.xlsx")
# 	pass

def findMaxNum(arr):
	maxNum = 0
	for x in arr:
		b = float(x)
		if maxNum < b:
			maxNum = b
			pass
	return maxNum

total = 0.0
count = 0

wb = Workbook() #创建文件对象
ws = wb.active  #获取默认sheet


for single in newpdf:
	# print(single)
	count = count + 1
	matchObj = new_re.findall(single)
	maxNum = findMaxNum(matchObj)
	title = '第' + str(count) + '张:'
	ws.append([title,maxNum])
	# print("第",count,"张:",maxNum)
	total = total + maxNum

heji = '=sum(B1:B' + str(count) + ')'
ws.append(['合计：',heji])

wb.save("heji.xlsx")

print(total)