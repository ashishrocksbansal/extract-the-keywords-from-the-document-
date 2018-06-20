import PyPDF2
from rake_nltk import Rake
import xlsxwriter

#first we will extract the text from the pdf.To do this PyPDF2 is used/

#creating a pdf file object
pdfFileObj = open(r'C:\Users\ashu\Downloads\JavaBasics-notes.pdf', 'rb')
 
# creating a pdf reader object
pdfReader = PyPDF2.PdfFileReader(pdfFileObj)


#total number of pages
total_pages=pdfReader.numPages

#storing all the pages text in one variable mytext
mytext=""
for page_number in range(total_pages):
    mytext+=pdfReader.getPage(page_number).extractText()


#now,we will use this text to extract keywords from it using RAKE(rapid automation keyword extraction)

#creating a Rake object
r=Rake()

#extract keywords from text
r.extract_keywords_from_text(mytext)

#get keywords from highest to lowest order with score
keywords_with_scores=r.get_ranked_phrases_with_scores()

#now,we will write it to the excel sheet

#creating the excel file 
workbook = xlsxwriter.Workbook('keywords_extraction.xlsx')
worksheet = workbook.add_worksheet()

# Start from the first cell. Rows and columns are zero indexed.
row = 0
col = 0

#declaring the names of the two rows
worksheet.write(row, col,     'keywords')
worksheet.write(row, col + 1, 'score')
row += 1

# Iterate over the data and write it out row by row.
for score,keyword in (keywords_with_scores):
    worksheet.write(row, col,     keyword)
    worksheet.write(row, col + 1, score)
    row += 1
workbook.close()


