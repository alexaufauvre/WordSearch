from sys import argv
import os.path
from string import punctuation
from collections import *
from docx import Document
from pptx import Presentation
import csv
import re
import docx2txt
import pandas as pd
from tkinter import *
from tkinter import ttk
from tkinter.filedialog import askdirectory, askopenfilename, askopenfilenames, asksaveasfilename
from PIL import Image, ImageTk
from itertools import chain

# Globals
filenames = []
keywords_filename = ""
sentences = []
keywords = []
success_msg = ""
labels = {}
labels_file = {}

# Get the filename with the keywords
def get_filename():
        # Tk().withdraw()
        # Filetypes allowed
        filetypes = [("Excel files", "*.xlsx *.xls")]
        print("Initializing Dialogue... \nPlease select a file.")
        filename = askopenfilename(initialdir=os.getcwd(), title='Sélectionnez un fichier', filetypes=filetypes)
        if filename:
                global keywords_filename
                keywords_filename = filename
                return filename 
        else:
                pass

# Returns a String containing the list of the names of the imported files  
def preview_filenames(filenames):
        filenameslabel_str = ""
        global filenameslabel
        for i in range (len(filenames)):
                filenameslabel_str = filenameslabel_str + "\n" + filenames[i].split("/")[-1]
        filenameslabel = filenameslabel_str
        return filenameslabel

# Returns the list of the names of the imported files 
def get_filenames():
        # Tk().withdraw()
        # Filetypes allowed
        filetypes = [("Excel files","*.xlsx *.xls"),("CSV files", "*.csv"), ("Text files", "*.txt *.rtf"), ("Word files", "*.docx"), ("Powerpoint files", "*.pptx")]
        print("Initializing Dialogue... \nPlease select a file.")
        tk_filenames = askopenfilenames(initialdir=os.getcwd(), title='Sélectionnez un ou plusieurs fichiers', filetypes=filetypes)
        if tk_filenames:
                filenames_list = list(tk_filenames)
                global filenames
                filenames = filenames_list
                preview_filenames(filenames)
                return filenames_list 
        else:
                pass


# Get the keywords from the file Excel format
def get_keywords():
        keywords_file = get_filename()
        print(keywords_file)
        ext = os.path.splitext(keywords_file)[1][1:]
        print(ext)
        if ext:
                if ext == "xls" or ext == "xlsx":
                        # Retrieving the data
                        data = pd.read_excel(keywords_file, header = None, encoding="utf-8")
                        # Cleaning the data
                        df = data.dropna(how="all")
                        df = df.replace(u'\xa0',' ')
                        # From Dataframe to Numpy array
                        data_array = df.get_values()
                        cells = []
                        cleanCells = []
                        # Extracting raw data from the dataframe
                        for value in data_array:
                                cells.append(value)
                        print("Cells")
                        print(cells)
                        # Filtering the useful data (text from the Excel)
                        for cell in cells:
                                for element in cell:
                                        if isinstance(element, str):
                                                element = element.replace(u'\xa0',' ')
                                                cleanCells.append(element)
                        error_msg.set("")
                        global keywords
                        keywords = cleanCells
                        print ("\nFile read finished!")
                        print (cleanCells)
                        return cleanCells
                else:
                        print ("Please enter a valid file")
                        error_msg.set("Veuillez entrer un fichier au format Excel")
                        return 0
        else:
                pass

# Retrieve all the sentences from the input files
def get_all_sentences():
        filenames_list = get_filenames()
        all_sentences = []
        # Processing file by file
        for one_filename in filenames_list:
                print ("Text file to import and read:", one_filename)
                print ("\nReading file...\n")
                
                # Get the extension of the file
                ext = os.path.splitext(one_filename)[1][1:]

                # Extracting the sentences of the file 

                # Text file
                if ext == "txt" or ext == "rtf":
                        text_file = open(one_filename, 'r')
                        all_lines = text_file.readlines()
                        text_file.close()
                        for sentence in all_lines.split("."):
                            all_sentences.append(sentence)
                        print ("\nFile read finished!")
                
                # CSV file
                elif ext == "csv":
                        text_file = open(one_filename, 'r')
                        all_lines = text_file.readlines()
                        text_file.close()
                        for value in all_lines:
                                for sentence in value.split("."):
                                        all_sentences.append(sentence)
                        print ("\nFile read finished!")

                # Word file
                elif ext == "docx": 
                        # Retrieve all the text
                        text_data = docx2txt.process(one_filename)
                        for sentence in text_data.split("."):
                                all_sentences.append(sentence)
                        print ("\nFile read finished!")

                # Excel file
                elif ext == "xls" or ext == "xlsx":
                        # Retrieving the data
                        data = pd.read_excel(one_filename, header = None, encoding='utf-8')
                        # Cleaning the data
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
                        for cell in cleanCells:
                                for sentence in str(cell).split("."):
                                        all_sentences.append(sentence)
                        print ("\nFile read finished!")


                # Powerpoint file       
                elif ext =="pptx":
                        prs = Presentation(one_filename)
                        for slide in prs.slides:
                                for shape in slide.shapes:
                                        if shape.has_text_frame:
                                                for sentence in shape.text.split("."):
                                                        sentence = sentence.replace(u'\xa0',' ')
                                                        all_sentences.append(sentence)
                        print ("\nFile read finished!")
        print ("All sentences retrieved!")
        # Optimization : Splitting sentences with no point
        all_sentences = list(flatmap(lambda x: x.splitlines(), all_sentences))
        global sentences
        sentences = all_sentences
        return all_sentences


# Setting of the regex pattern : spaces and/or punctuation between and after the word
def word_matches(word, sentence):
        sentence = "." + sentence + "." 
        pattern = re.compile(r'.*(\s|\W)+' + re.escape(word) + r'(\s|\W)+.*')
        return re.match(pattern, sentence)


# Create the output document and process the matching
def keywords_matching(keywords_list, sentences_list, label_error):
        global filenames, keywords_filename, labels, error_msg
        if keywords_filename == "" or len(filenames) == 0:
                error_msg.set("Veuillez choisir au moins deux fichiers")
                return label_error.update()
        else:
                error_msg.set("")
                label_error.update()
                # Adding the matching sentences to a file
                output_file = asksaveasfilename(defaultextension=".docx", title="Choisissez où vous voulez enregistrer votre fichier de réponses")
                document = Document()
                document.add_heading('Résultats', 0)
                paragraphs = {}
                for idx, word in enumerate(keywords_list):
                        cpt = 0
                        idx = str(idx)
                        paragraphs["key" + idx] = document.add_paragraph('')
                        paragraphs["key" + idx].add_run( "- %s : \n\n" % word).bold = True
                        paragraphs["p" + idx] = document.add_paragraph('')
                        # Case insensitive matching
                        lower_word = word.lower()
                        for sentence in sentences_list: 
                                if word_matches(lower_word, sentence.lower()):
                                        paragraphs["p" + idx].add_run( "\t--> %s.\n\n" % sentence)
                                        cpt +=1
                        paragraphs["key" + idx].add_run(str(cpt) + " occurrences").italic = True
                global success_msg
                success_msg = 'Fichier de résultats ' + os.path.basename(output_file) + ' téléchargé !'
                document.save(output_file)
                print (success_msg)
                print ("\nLabels -->\n")
                print (labels)
                return output_file

# # Convert .doc file to .docx
# def convert_to_docx(filename):
#         doc_file = filename
#         docx_file = filename + 'x'
#         if not os.path.exists(docx_file):
#                 os.system('mv ' + doc_file + ' ' + docx_file)
#         else:
#           # already a file with same name as doc exists having docx extension, 
#           # which means it is a different file, so we cant read it
#                 print('Info : file with same name of doc exists having docx extension, so we cant read it')
#         return docx_file

# Homemade flatmap function
def flatmap(f, items):
        return chain.from_iterable(list(map(f, items)))

# Set text to a label
def set_text(label, text):
        label["text"] = text
        return label.update()

# Set multi text to labels --> Set text[1] to label[1], etc.
def set_text_multi(labels, text):
        print(labels_file)
        for i in range(len(filenames)): 
                labels["label_file" + str(i)]["text"] = filenames[i].split("/")[-1]
                labels["label_file" + str(i)].update
        return labels

# Delete all the data created by the user
def reset(mw):
        global sentences, filenames, keywords_filename, keywords, error_msg, success_msg, labels_file, labels
        # Reset global variables
        filenames = []
        keywords_filename = ""
        sentences = []
        keywords = []
        success_msg = ""
        error_msg.set("")
        # Reset keyword file imported
        labels["label_keyword_file"]["text"] = keywords_filename
        # Reset UI messages
        labels["label_error"]["text"] = error_msg
        labels["label_success"]["text"] = success_msg
        # Reset files imported
        reset_labels(labels_file)
        return labels_file

# Delete all labels in the input table
def reset_labels(labels):
        for item in labels.values():
                item.destroy()
        labels.clear()
        return labels

# Create a label for each text in the list
def create_labels(mw, text_list, fg):
        global labels_file
        # Reset files imported
        reset_labels(labels_file)     
        if isinstance(text_list, list):
                for i in range (len(text_list)):
                        id_label = "label_file" +  str(i)
                        new_label = Label(mw, text = text_list[i].split("/")[-1], fg = fg)
                        new_label.pack()
                        labels_file[id_label] = new_label
                return labels_file
        else:
                pass

# Create a label with textvariable
def create_label_var(mw, id_label, textvar, fg):
        global labels
        new_label = Label(mw, textvariable = textvar, fg = fg)
        new_label.pack()
        labels[id_label] = new_label
        return new_label

# Create a label with static text
def create_label_static(mw, id_label, text, fg):
        global labels
        new_label = Label(mw, text = text, fg = fg)
        new_label.pack()
        labels[id_label] = new_label
        return new_label

# Update the labels
def update_labels(labels):
        print(labels)
        for label in labels.values():
                label.update()
        return labels



# GUI
def main():
        # Init window
        root = Tk()
        root.title('Welcome to the Kanbios Parser')
        root.geometry("850x650")
        ttk.Style().configure("TButton", padding=6, relief="flat", background="#ccc")
        # Init frame
        mw = Frame(root)
        mw.pack(padx=20, pady=20)
        global labels, error_msg, filenameslabel
        error_msg = StringVar()
        filenameslabel = StringVar()

        # Logo Kanbios
        logo_canvas = Canvas(mw, width=371, height=105)
        logo=PhotoImage(file="logo-kanbios-resized.png")
        logo_canvas.create_image(5, 5,image=logo, anchor="nw")
        logo_canvas.pack(side=TOP)

        # Step 1
        create_label_static(mw, 'label_title1', '\n1 - Importer vos mots-clés (fichier Excel)', 'black')
        # Get the name of the keywords file
        keywords_filename_btn = ttk.Button(mw, text = 'Parcourir...', command = (lambda : get_keywords() and set_text(labels["label_keyword_file"], keywords_filename.split("/")[-1])))
        keywords_filename_btn.pack()
        # Preview of the file name
        create_label_var(mw, 'label_keyword_file', keywords_filename, 'blue')
        labels["label_keyword_file"]["text"] = keywords_filename

        # Step 2
        create_label_static(mw, 'label_title2', '\n2 - Importer vos fichiers à analyser', 'black')
        # Get the names of the files to analyse
        get_filenames_btn = ttk.Button(mw, text = 'Parcourir...', command = lambda : get_all_sentences() and create_labels(mw, filenames, 'blue') and set_text_multi(labels_file, filenames))
        get_filenames_btn.pack()
        # Preview of the file name [If one label containing all file names]
        # create_label_var(mw, 'label_importfile', filenameslabel, 'blue')
        
        # Step 3
        create_label_static(mw, 'label_title3', '\n\n3 - Télécharger votre fichier de correspondance', 'black')
        # Process the matching
        process_btn = ttk.Button(mw, text = 'Télécharger', command = lambda : keywords_matching(keywords, sentences, labels["label_error"]) and set_text(labels["label_success"], success_msg))
        process_btn.pack()

        # Success message
        create_label_var(mw, 'label_success', success_msg, 'green')

        # Error message
        create_label_var(mw, 'label_error', error_msg, 'red')

        # Reset the data
        reset_btn = ttk.Button(mw, text = 'Réinitialiser', command = lambda : reset(mw) and update_labels(labels))
        reset_btn.pack(side=BOTTOM)
       

        # Close the window
        close_btn = ttk.Button(mw, text = 'Fermer', command = root.destroy)
        close_btn.pack(side=BOTTOM)

        mw.mainloop()

# Main
main()

