import docx
import xlsxwriter
import os
import requests

dict1 = {}

for filename in os.scandir():
	doc = docx.Document(filename)
	for para in doc.paragraphs:
		for word in para.strip(",.!?").split():
			if word in dict1:
				dict1[word] += 1
			else:
				dict1[word] = 1



url = "https://google-translate1.p.rapidapi.com/language/translate/v2"

payload = "q=Hello%2C%20world!&source=en&target=es"
headers = {
    'content-type': "application/x-www-form-urlencoded",
    'accept-encoding': "application/gzip",
    'x-rapidapi-key': "2da75cc372msh80223adaecf7041p10bee7jsnb0f534bca746",
    'x-rapidapi-host': "google-translate1.p.rapidapi.com"
    }

response = requests.request("POST", url, data=payload, headers=headers)

print(response.text)