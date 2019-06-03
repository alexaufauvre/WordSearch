from sys import argv
import os.path
from string import punctuation
from collections import *
from docx import Document
from pptx import Presentation
import csv
import re
import pandas as pd
from tkinter import *
from tkinter import ttk
from tkinter.filedialog import askdirectory, askopenfilename, askopenfilenames, asksaveasfilename
from PIL import Image, ImageTk

# Globals
filenames = []
keywords_filename = ""
sentences = []
keywords = []
success_msg = ""
labels = {}
labels_file = {}

def get_filename():
        # Tk().withdraw()
        print("Initializing Dialogue... \nPlease select a file.")
        filename = askopenfilename(initialdir=os.getcwd(), title='Sélectionnez un fichier')
        if filename:
                global keywords_filename
                keywords_filename = filename
                return filename 
        else:
                pass

def preview_filenames(filenames):
        filenameslabel_str = ""
        global filenameslabel
        for i in range (len(filenames)):
                filenameslabel_str = filenameslabel_str + "\n" + filenames[i].split("/")[-1]
        filenameslabel = filenameslabel_str
        return filenameslabel

def get_filenames():
        # Tk().withdraw()
        print("Initializing Dialogue... \nPlease select a file.")
        tk_filenames = askopenfilenames(initialdir=os.getcwd(), title='Sélectionnez un ou plusieurs fichiers')
        if tk_filenames:
                filenames_list = list(tk_filenames)
                global filenames
                filenames = filenames_list
                preview_filenames(filenames)
                return filenames_list 
        else:
                pass


def get_keywords():
        keywords_file = get_filename()
        print(keywords_file)
        ext = os.path.splitext(keywords_file)[1][1:]
        print(ext)
        if ext:
                if ext == "csv":
                        error_msg.set("")
                        text_file = open(keywords_file, 'r')
                        all_lines = text_file.readlines()
                        text_file.close()
                        all_keywords = []
                        for value in all_lines:
                                if value:
                                        clean_value = re.sub('[^a-zA-Z0-9]+', '', value)
                                        all_keywords.append(clean_value)
                        global keywords
                        keywords = all_keywords
                        print ("\nFile read finished!")
                        print (all_keywords)
                        return all_keywords
                else:
                        print ("Please enter a valid file")
                        error_msg.set("Veuillez entrer un fichier au format CSV")
                        return 0
        else:
                pass

def get_all_sentences():
        filenames_list = get_filenames()
        # Processing file by file
        all_sentences = []
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
                elif ext == "doc" or ext == "docx":
                        text_file = Document(one_filename)
                        all_paragraphs = text_file.paragraphs
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
                        for cell in cleanCells:
                                for sentence in str(cell).split("."):
                                        all_sentences.append(sentence)
                        print ("\nFile read finished!")


                # Powerpoint file       
                elif ext == "ppt" or ext =="pptx":
                        prs = Presentation(one_filename)
                        for slide in prs.slides:
                                for shape in slide.shapes:
                                        if shape.has_text_frame:
                                                for sentence in shape.text.split("."):
                                                        sentence = sentence.replace(u'\xa0',' ')
                                                        all_sentences.append(sentence)
                        print ("\nFile read finished!")

        print ("All sentences retrieved!")
        global sentences
        sentences = all_sentences
        return all_sentences


 # Setting of the regex pattern : spaces and/or punctuation between and after the word
def word_matches(word, sentence):
        sentence = "." + sentence + "." 
        pattern = re.compile(r'.*(\s|\W)+' + re.escape(word) + r'(\s|\W)+.*')
        return re.match(pattern, sentence)


def keywords_matching(keywords_list, sentences_list, label_error):
        global filenames, keywords_filename, labels, error_msg
        if keywords_filename == "" or len(filenames) == 0:
                error_msg.set("Veuillez choisir au moins deux fichiers")
                return label_error.update()
        else:
                error_msg.set("")
                label_error.update()
                # Adding the matching sentences to a file
                output_file = open(asksaveasfilename(title='Choisissez où vous voulez enregistrer votre fichier de réponses', defaultextension=".txt"), "w")
                for word in keywords_list:
                        output_file.write( "- %s : \n\n" % word)
                        # Case insensitive matching
                        lowerWord = word.lower()
                        for sentence in sentences_list: 
                                if word_matches(lowerWord, sentence.lower()):
                                        output_file.write( "\t--> %s.\n\n" % sentence)
                global success_msg
                success_msg = 'Fichier de résultats ' + os.path.basename(output_file.name) + ' téléchargé !'
                output_file.close()
                print (success_msg)
                print ("\nLabels -->\n")
                print (labels)
                return success_msg

# Set text to a label
def set_text(label, text):
        label["text"] = text
        return label.update()

# Set multi text to labels
def set_text_multi(labels, text):
        print(labels_file)
        for i in range(len(filenames)): 
                labels["label_file" + str(i)]["text"] = filenames[i].split("/")[-1]
                labels["label_file" + str(i)].update
        return labels

def reset(mw):
        global sentences, filenames, keywords_filename, keywords, error_msg, success_msg, labels_file, labels
        filenames = []
        keywords_filename = ""
        sentences = []
        keywords = []
        success_msg = ""
        error_msg.set("")
        # Reset keyword file imported
        labels["label_keyword_file"]["text"] = keywords_filename
        labels["label_error"]["text"] = error_msg
        labels["label_success"]["text"] = success_msg

        # Reset files imported
        reset_labels(labels_file)
        return labels_file


def reset_labels(labels):
        for item in labels.values():
                item.destroy()
        labels.clear()
        return labels


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

def create_label_var(mw, id_label, textvar, fg):
        global labels
        new_label = Label(mw, textvariable = textvar, fg = fg)
        new_label.pack()
        labels[id_label] = new_label
        return new_label

def create_label_static(mw, id_label, text, fg):
        global labels
        new_label = Label(mw, text = text, fg = fg)
        new_label.pack()
        labels[id_label] = new_label
        return new_label

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
        create_label_static(mw, 'label_title1', '\n1 - Importer vos mots-clés', 'black')
        # Get the name of the keywords file
        keywords_filename_btn = ttk.Button(mw, text = 'Parcourir... (.csv)', command = (lambda : get_keywords() and set_text(labels["label_keyword_file"], keywords_filename.split("/")[-1])))
        keywords_filename_btn.pack()
        # Preview of the file name
        create_label_var(mw, 'label_keyword_file', keywords_filename, 'blue')
        labels["label_keyword_file"]["text"] = keywords_filename

        # Step 2
        create_label_static(mw, 'label_title2', '\n2 - Importer vos fichiers à analyser', 'black')
        # Get the names of the files to analyse
        get_filenames_btn = ttk.Button(mw, text = 'Parcourir...', command = lambda : get_all_sentences() and set_text(labels["label_importfile"], filenameslabel))
        get_filenames_btn.pack()
        # Preview of the file name
        create_label_var(mw, 'label_importfile', filenameslabel, 'blue')
        
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
        reset_btn.pack()

        # Close the window
        close_btn = ttk.Button(mw, text = 'Fermer', command = root.destroy)
        close_btn.pack(side=BOTTOM)

        mw.mainloop()

# Main
main()

