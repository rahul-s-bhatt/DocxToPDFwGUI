import sys
import os
import comtypes.client

wdFormatPDF = 17
infolder="C:\\Users\\BHATT\\Desktop\\ConvertedPDF"
out_folder ="C:\\Users\\BHATT\\Desktop\\ConvertedPDF"

word = comtypes.client.CreateObject('Word.Application')
for in_file_name in os.listdir(infolder):
	print(in_file_name)
	in_file = infolder + '\\' + in_file_name
	doc = word.Documents.Open(in_file)
	print("\n"+in_file+" opened")
	
	outfile_name = in_file_name.replace("docx","pdf")
	out_file = out_folder + '\\' + outfile_name
	doc.SaveAs(out_file, FileFormat=wdFormatPDF)
	doc.Close()
	print("successfully converted"+outfile_name)
word.Quit()
	
