from sys import argv
import os.path
from string import punctuation
from collections import *
from docx import Document
from pptx import Presentation
import csv
import re
import pandas as pd


keyWords = ['God', 'national','power', 'superpower', 'countries', 'Council', 'Nations', 'nation', 'USA', 'Creater', 'creater', 'Country', 'Almighty',
             'country', 'People', 'people', 'Liberty', 'liberty', 'America', 'Independence', 
             'honor', 'brave', 'Freedom', 'freedom', 'Courage', 'courage', 'Proclamation',
             'proclamation', 'United States', 'Emancipation', 'emancipation', 'Constitution',
             'constitution', 'Government', 'Citizens', 'citizens']

 # Setting of the regex pattern : spaces and/or punctuation between and after the word
def word_matches(word, sentence):
        sentence = "." + sentence + "." 
        pattern = re.compile(r'.*(\s|\W)+' + re.escape(word) + r'(\s|\W)+.*')
        return re.match(pattern, sentence)

file_counter = 0

# Processing file by file
for one_filename in argv[1:]:
        print ("Text file to import and read:", one_filename)
        print ("\nReading file...\n")
        
        file_counter += 1
        # Get the extension of the file
        ext = os.path.splitext(one_filename)[1][1:]

        # Extracting the sentences of the file 

        # Text file
        if ext == "txt" or ext == "rtf":
                text_file = open(one_filename, 'r')
                all_lines = text_file.readlines()
                text_file.close()
                all_sentences = all_lines.split(".")
                print ("\nFile read finished!")
        
        # CSV file
        elif ext == "csv":
                text_file = open(one_filename, 'r')
                all_lines = text_file.readlines()
                text_file.close()
                all_sentences = []
                for value in all_lines:
                        for sentence in value.split("."):
                                all_sentences.append(sentence)
                print ("\nFile read finished!")

        # Word file
        elif ext == "doc" or ext == "docx":
                text_file = Document(one_filename)
                all_paragraphs = text_file.paragraphs
                all_sentences = []
                for paragraph in all_paragraphs:
                        for sentence in paragraph.text.split("."):
                                all_sentences.append(sentence)
                print ("\nFile read finished!")

        # Excel file
        elif ext == "xls" or ext == "xlsx":
                data = pd.read_excel(one_filename, encoding='utf-8')
                df = data.dropna(how="all")
                df = df.replace(u'\xa0',' ')
                # From Dataframe to Numpy array
                data_array = df.get_values()
                cells = []
                cleanCells = []
                # Extracting raw data from the dataframe
                for value in data_array:
                        cells.append(value)
                # Filtering the useful data (text from the Excel)
                for cell in cells:
                        for element in cell:
                                if isinstance(element, str):
                                        element = element.replace(u'\xa0',' ')
                                        cleanCells.append(element)
                # --------
                # For testing purpose
                # data_file = open("DataExcel.txt", "w", encoding="utf-8")
                # for cell in cleanCells:
                #         data_file.write(str(cell) + "\n\n")
                # data_file.close()
                # --------
                all_sentences = []
                for cell in cleanCells:
                        for sentence in str(cell).split("."):
                                all_sentences.append(sentence)
                print ("\nFile read finished!")


        # Powerpoint file       
        elif ext == "ppt" or ext =="pptx":
                prs = Presentation(one_filename)
                all_sentences = []
                for slide in prs.slides:
                        for shape in slide.shapes:
                                if shape.has_text_frame:
                                        for sentence in shape.text.split("."):
                                                sentence = sentence.replace(u'\xa0',' ')
                                                all_sentences.append(sentence)
                print ("\nFile read finished!")

        print ("Write results: List_of_words" + str(file_counter) + ".txt")

        # Adding the matching sentences to a file
        output_file = open("List_of_words" + str(file_counter) + ".txt", "w")
        output_file.write( "-------- %s -------- \n\n\n" % one_filename )
        for word in keyWords:
                output_file.write( "- %s : \n\n" % word)
                # Case insensitive matching
                lowerWord = word.lower()
                for sentence in all_sentences: 
                        if word_matches(lowerWord, sentence.lower()):
                                output_file.write( "\t--> %s.\n\n" % sentence)
                       

# Write the sentence in the new document if it matches the Keyword list -- IT WORKS 
# for word in keyWords:
#     for paragraph in all_paragraphs: 
#             for sentence in paragraph.text.split("."):
#                  if word in sentence:
#                          output_file.write( "- %s.\n\n" % (sentence) )

        output_file.close()

