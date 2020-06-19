from docxtpl import DocxTemplate
import comtypes.client
import os
import ast

########################################### - Reads the context.txt and convert to dictionary
def read_txt_file(input_string):
	if(input_string.endswith('.txt')):
		text_file = open(input_string, 'r')
		text = text_file.read()
		text_file.close()
		return text
	else:
		return input_string

file = open('context.txt', 'r')
content = file.read()
context = ast.literal_eval(content)

for key in context.keys():
	input_string = input (f'{key}: ')  # - User input to update context values
	context[key] = read_txt_file(input_string)

file.close()

############################################ - Opens the template and use context to update the document
doc = DocxTemplate('new-Coverletter.docx')
title = input('new file name: ')
doc.render(context)
doc.save(title + '.docx')

############################################# - Save the updated document to PDF
word = comtypes.client.CreateObject('Word.Application')
word_doc = word.Documents.Open(os.getcwd() + '\\' + title + '.docx')
word_doc.SaveAs(os.getcwd() + '\\' + title + '.pdf', FileFormat=17)
word_doc.Close()
word.Quit()