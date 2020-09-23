"""A script that will copy the sentences containing a specific word
Made by: Hans-Kristian Norum Eidesen
E-mail: hans.kristian.eidesen@dnvgl.com
"""

import PyPDF2 # read the PDF
from pprint import pprint
import pandas as pd
import numpy as np

from pdfminer3.pdfparser import PDFParser
from pdfminer3.pdfdocument import PDFDocument
from pdfminer3.pdfinterp import PDFResourceManager, PDFPageInterpreter # process_pdf
from pdfminer3.pdfpage import PDFPage
from pdfminer3.converter import TextConverter
from pdfminer3.layout import LAParams

from io import StringIO

import csv
import re
from string import digits

def __init__(self, ToC, output, paragraph_one_line):
    self.ToC = ToC
    self.output = output
    self.paragraph_one_line = paragraph_one_line

# Function to convert list to string   
def listToString(s):  
    
    # initialize an empty string 
    str1 = ""  
    
    # traverse in the string   
    for ele in s:  
        str1 += ele   
    
    # return string   
    return str1  

def copy_words(file, word1, word2):
    
    """copy the specific sentences containing search words

    Args:
        file (the pdf-file that will be read): reads and extract the specific words from the string
        word1 (string): word1 that the function will look for, e.g. "shall"
        word2 (string): word2 that the funciton will look for, e.g. "should"
    """
    print("Looking for the specified words, \"", word1,"\"",  "and", "\"",word2,"\"") 
    # Using PyPDF2 to convert the .pdf to text
    filename = file
    read_pdf = open(filename, 'rb')
    # Reading the .pdf, outputting text 
    pdfReader = PyPDF2.PdfFileReader(read_pdf)
    count = pdfReader.numPages
    # output = is where the result is stored at
    output = []

    # converting the output to one sting
    modified_text_list = []    
    modified_text_list = pd.DataFrame(columns=["Text"])
    
    for i in range(count):
        page = pdfReader.getPage(i)
        output.append(page.extractText())  
        

    # print(modified_text_list)
    # modified_text_list.to_excel("test_text.xlsx")

    output = listToString(output)

    f = open('output1.csv', 'w', encoding='utf-8')
    f.write(output)
    f.close()
    """
    The block will take the result from the pdf-to-string conversion, search for the words (word1 and word2), and return those lines containing the words.
    """
    # Store paragraph in a variable.
    paragraph = output
    # print(paragraph)
    # Split each sentence by the "." (period)
    sentences_list = paragraph.split(".")
    # empty list that will contain the results
    sentences_with_word = []
    # A list that stores the words we want to look for.    
    words_search = [word1, word2]
    # # to make sure that no lowercase and uppercase errors are found, every character in paragraph and words_search is converted to lowercase
    # paragraph = paragraph.lower()
    # words_search = [word1.lower(), word2.lower()]
    # Convertint them to a dictiorairy
    word_sentence_dictionary = {word1:[], word2:[]}

    # Now, we are ready to start our search and store the sentences that contain the word.

    for word in words_search:
        for sentence in sentences_list:
            if sentence.count(word)>0:
                sentences_with_word.append(sentence)
                word_sentence_dictionary[word] = sentences_with_word
    return sentences_with_word


def remove_digits_in_list(lst):
    pattern = '[0-9]'
    lst = [re.sub(pattern, '', i) for i in lst]
    pattern2 = '[.]'
    lst = [re.sub(pattern2, '', j) for j in lst]
    return lst

def get_ToC(file):
    """This funciton will locate the Table of Content, and return a dataframe with the corresponding ToC-number and name 

    Args:
        file (the pdf-file that will be read): reads and extract the specific words from the string
    """
    # Open a PDF document.
    fp = open(file, 'rb')
    parser = PDFParser(fp)
    document = PDFDocument(parser)

    # Get the outlines of the document.
    # Append result to dataframe
    ToC = [] # Table of Content
    outlines = document.get_outlines()
    for (level,title,dest,a,se) in outlines:
        # ToC.append(level)
        ToC.append(title)
        # print (level,title,dest,a,se )
    # print(ToC)
    # For some reason when the .PDF is being transformed to text, the corresponding chapter-number is removed
    # This section will thus remove the chapter number, and only the chapter name is being used to lookup the 
    # chapter text
    text_only = []
    df_ToC = pd.DataFrame(ToC, columns=['ToC'])
    df_ToC['Chapter text'] = (df_ToC['ToC'].str.replace('[0-9.]', '')).str.lstrip() #removing numbers and first space, leaving the text only
    df_ToC['Chapter number'] = (df_ToC['ToC'].str.replace('[a-zA-Z (),–/]', '')).str.replace(r"^(\d+)\.$", r"\1")# .str.lstrip()
    df_ToC.to_excel('ToC.xlsx')
    # print(ToC)
    return df_ToC

def get_text_between_string(string, first, last):
    try:
        start = string.index(first) + len(first)
        end = string.index(last, start)
        return string[start:end]
    except ValueError:
        return ""


def get_text_in_chapter(file, word1, word2):

    filename = file
    read_pdf = open(filename, 'rb')
    # Reading the .pdf, outputting text 
    pdfReader = PyPDF2.PdfFileReader(read_pdf)
    count = pdfReader.numPages
    # output = is where the result is stored at
    output = []

    # converting the output to one sting
    for i in range(count):
        page = pdfReader.getPage(i)
        output.append(page.extractText())
    output = listToString(output)
    paragraph_one_line = output.replace("\n", " ")

    modified_text = paragraph_one_line
    
    # f = open('paragraph_as_one_line1.csv', 'w', encoding='utf-8')
    # f.write('modified_text')
    # f.close()

    chapters = get_ToC(file) # a df with all the chapter names
    # find the pages where the text is after the ToC (e.g. where the actual text starts)
    start_page = []
    for i in range(0, count):
        PageObj = pdfReader.getPage(i)
        # print("this is page " + str(i)) 
        Text = PageObj.extractText() 
        # print(Text)
        ResSearch = re.search('1 Scope', Text) 
        if ResSearch != None:
            # print(ResSearch)
            # print("The page where the text is ", i)
            start_page.append(i)
            
    page_of_interest = start_page[1] # the page of interesert (where the looking should begin) is usually the second time the first entry in ToC (e.g. "Scope" etc.) comes in the document
    print("The page of intereset is: ", page_of_interest)
    pdf_to_one_string = ""
    output = []

    for n in range(page_of_interest, count):
        page = pdfReader.getPage(n)
        output.append(page.extractText())

    output = listToString(output)
    paragraph_one_line = output.replace("\n", "") # making the entire document to one line
    paragraph_one_line = " ".join(paragraph_one_line.split())
    paragraph_one_line = paragraph_one_line.replace('NORSOK S-001:2018 12 NORSOK © 2018','') # removing stupid watermarks
    paragraph_one_line = paragraph_one_line.replace("NORSOK S-001:2018 provided by Standard Online AS for DNV GL Group Companies 2018-06-21", "") # removing stupid watermark

    f = open('paragraph_as_one_line.csv', 'w', encoding='utf-8')
    f.write(paragraph_one_line)
    
    number_of_chapters = len(chapters) # total number of chapters and sub-chapters
    list_with_all_chapters = []
    remove_digits = str.maketrans('','', digits) # removing digts from the search string
   
    df_all_chapters = pd.DataFrame(columns=['Result'])
    sentences_containing_the_search_words = []
    words_search = [word1, word2]
    word_sentence_dictionary = {word1:[], word2:[]}

    logic0 = 0
    logic1 = 0
    logic2 = 0
    logic3 = 0
    logic4 = 0
    step = 0
    
    for word in words_search:
        for sentence in paragraph_one_line:
            if sentence.count(word) > 0:
                sentences_containing_the_search_words.append(sentence)
                word_sentence_dictionary[word] = sentences_containing_the_search_words
    
    for n in range(100): #number of chapter
        # print(chapters[n])
        current_chapter = ''
        next_chapter = ''
        try:
            print("\n")
            print('step', step)
            # current_chapter = ''
            # next_chapter = ''
            # get_text_between_string = ''
            # print((re.sub('[!@#$.]', "", chapters[n]).translate(remove_digits)))
            """Logic 0
               Getting the chapter text based on what is between each chapter heading (chapter number and chapter text)
            """
            list_with_all_chapters.append(chapters['ToC'][n])
            current_chapter = chapters['ToC'][n]
            next_chapter = chapters['ToC'][n+1]
            get_text_between_string(paragraph_one_line, current_chapter, next_chapter)
           
            if get_text_between_string(paragraph_one_line, current_chapter, next_chapter) != '':
                current_chapter = chapters['ToC'][n]
                next_chapter = chapters['ToC'][n+1]
                logic0 += 1
            """Logic 1:
               Getting the chapter text based on what is between each chapter heading (chapter text only)
            """
            if get_text_between_string(paragraph_one_line, current_chapter, next_chapter) == '':
                print('Logic 0 failed!')
                print('Switching to logic 1')
                current_chapter = chapters['Chapter text'][n]
                next_chapter = chapters['Chapter text'][n+1]

                list_with_all_chapters.append(get_text_between_string(paragraph_one_line, current_chapter, next_chapter)) # works. 
                logic1 += 1
            """Logic 2:
                Getting the chapter text based on what is between each chapter heading (chapter numbers only)
                This logic us mostly used when there are 4x-sub chapters (e.g. a.b.c.d. )

                NOTE: Since this logic looks for numbers, it is likely that the heading text is printed as chapter text.
                    This beahviour needs to be resolved
            """
            if get_text_between_string(paragraph_one_line, current_chapter, next_chapter) == '':
                print('Logic 1 failed!')
                print('Switching to logic 2')

                # print(current_chapter)
                current_chapter = ''
                next_chapter = ''           
                size = 0                
                # These rules looks for if there is a period '.' after the last number, removes it and calls the "get_text_between_string"-function

                if chapters['Chapter number'][n][-1] == '.':
                    # print(current_chapter)
                    current_chapter = chapters['Chapter number'][n][:-1]

                if chapters['Chapter number'][n+1][-1] == '.':
                    # print(current_chapter)
                    next_chapter = chapters['Chapter number'][n+1][:-1]

                # print(current_chapter)
                list_with_all_chapters.append(get_text_between_string(paragraph_one_line, current_chapter, next_chapter)) # works. The search term removes digits and special characters
                print(get_text_between_string(paragraph_one_line, current_chapter, next_chapter))
                logic2 += 1
                """Logic 3:
                    Mostly used if ther are any annexes or appendices, or any chapters starting with a uppercase letter.
                """
            if get_text_between_string(paragraph_one_line, current_chapter, next_chapter) == '':
                print('Logic 2 failed!')
                print('Switching to logic 3')
                
                current_chapter = chapters['Chapter number'][n]
                next_chapter = chapters['Chapter number'][n+1]
                
                print("Switching to logic 3")
                print('Logic 3 reports that current chapter is: ', current_chapter)
                print('Logic 3 reports that next chapter is: ', next_chapter)
                
                list_with_all_chapters.append(get_text_between_string(paragraph_one_line, current_chapter, next_chapter))         
                print("the text is now: ", get_text_between_string(paragraph_one_line, current_chapter, next_chapter))
                logic3 += 1

            elif get_text_between_string(paragraph_one_line, current_chapter, next_chapter) == '':
                print("Looks like we need a 4th logic bois")
                logic4 += 1
            #print(get_text_between_string)         
        except IndexError as err:
            print("An unexpected error occured,  {0}".format(err))
            print("The script will continue anyway") # Print "done" instead.

        text_to_copy = get_text_between_string(paragraph_one_line, current_chapter, next_chapter)
        # These lines will remove the current entry in the paragraph_one_line, which will ensure that chapters with similar names will have unique entries
        # paragraph_one_line = paragraph_one_line.replace(text_to_copy, '')
        paragraph_one_line = paragraph_one_line.replace(next_chapter, '')
        paragraph_one_line = paragraph_one_line.replace(current_chapter, '')
        step += 1
    # f = open('result.csv', 'w', encoding='utf-8')
    # f.write(listToString(list_with_all_chapters))

    #print(list_with_all_chapters)
    df = pd.DataFrame(list_with_all_chapters, columns=['Chapter and text'])
    print(df)
    df.to_excel('chapters.xlsx')
    print('logic0: ',logic0, 'logic1: ',logic1, 'logic2: ', logic2, 'logic3 :',logic3, 'logic4 :',logic4)

get_text_in_chapter('S-001_2018E.pdf', 'shall', 'should')
# copy_words('S-001_2018E.pdf', 'shall', 'should')

# print(get_text_between_string( 'Perform risk analyses and evaluations Risk and safety analyses / studies shall be performed to establish suf ficiently detailed information about the risk associated with the identified hazards and accidental events. The information will be used to evaluate the risk and to decide which solutions (barriers), and related requirements that are needed for preventing, controlling and mitigating the hazards in addition to generic requirements give n by the context. Evaluation of risk includes an assessment of: compliance with predefined evaluation criteria (e.g. minimum requirements and acceptance criteria in context) ; necessary ALARP processes to demonstrate that risk has been reduced to a level as low as reasonably practicable ; uncertainties associated with the hazards, accidental events and their consequences as well as the risk reducing effect of the barriers. The r esults from risk analyses are used for many purposes. One is to provide information and decision support related to the need for and role of risk reducing measures (barriers) and their required performance. Another is to provide decision support regarding the risk level assessed and if this is considered acceptable for the facility. Reference is made to sub clause 5.10 for examples of studies and evaluations. For a development project, the degree of details in the risk and safety analyses / studies will increase as the project matures through different phases. Thus, solutions, assumptions and conservative estimates typically used in early stages may be verified or changed due to more detailed analyses. Identify and define barrier functions, systems and elements (risk treatment)','Perform risk analyses and evaluations', 'Identify and define barrier functions, systems and elements (risk treatment)'))