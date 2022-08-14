import os, docx

current = 0
chapters = ["Chapter One","Chapter Two","Chapter Three","Chapter Four","Chapter Five","Chapter Six","Chapter Seven","Chapter Eight","Chapter Nine","Chapter Ten","Chapter Eleven","Chapter Twelve","Chapter Thirteen","Chapter Fourteen","Chapter Fifteen","Chapter Sixteen","Chapter Seventeen","Chapter Eighteen", "Glossary"]
directory = os.getcwd()
seperator = os.path.sep
appending_starts = False
book = docx.Document('name.docx') #NAME
for i in book.paragraphs:
	if i.text.find(chapters[current+1]) != -1:
		newDoc.save(directory+seperator+'chapter_'+str(current+1)+'.docx')

		if current+2  == len(chapters):
			break
		current += 1

	if i.text.find(chapters[current]) != -1:
		newDoc = docx.Document()
		appending_starts = True

	if appending_starts == True:
		newDoc.add_paragraph(i.text)


		

