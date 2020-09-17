"""A script that will copy the sentences containing a specific word
Made by: Hans-Kristian Norum Eidesen
E-mail: hans.kristian.eidesen@dnvgl.com
"""

import PyPDF2 # read the PDF
from pprint import pprint
import pandas as pd

from pdfminer3.pdfparser import PDFParser
from pdfminer3.pdfdocument import PDFDocument

    
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
    for i in range(count):
        page = pdfReader.getPage(i)
        output.append(page.extractText())
    output = listToString(output)
  
    """
    The block will take the result from the pdf-to-string conversion, search for the words (word1 and word2), and return those lines containing the words.
    """
    # Store paragraph in a variable.
    paragraph = output
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

    # The sentences containing the words are stored as lists in Pandas DataFrame
    columns = ['Results']
    df = pd.DataFrame(sentences_with_word, columns = columns)
    df.to_excel('result.xlsx')
    print(df.head())
    
copy_words('S-001_2018E.pdf', 'shall', 'should')

def get_font(obj, fnt, emb):

    '''
    If there is a key called 'BaseFont', that is a font that is used in the document.
    If there is a key called 'FontName' and another key in the same dictionary object
    that is called 'FontFilex' (where x is null, 2, or 3), then that fontname is 
    embedded.
    
    We create and add to two sets, fnt = fonts used and emb = fonts embedded.
    '''
    if not hasattr(obj, 'keys'):
        return None, None
    fontkeys = set(['/FontFile', '/FontFile2', '/FontFile3'])
    if '/BaseFont' in obj:
        fnt.add(obj['/BaseFont'])
    if '/FontName' in obj:
        if [x for x in fontkeys if x in obj]:# test to see if there is FontFile
            emb.add(obj['/FontName'])

    for k in obj.keys():
        get_font(obj[k], fnt, emb)

    return fnt, emb# return the sets for each page

if __name__ == '__main__':
    fname = 'S-001_2018E.pdf'
    read_pdf = open(fname, 'rb')
    pdf = PyPDF2.PdfFileReader(fname)
    fonts = set()
    embedded = set()
    for page in pdf.pages:
        obj = page.getObject()
        # updated via this answer:
        # https://stackoverflow.com/questions/60876103/use-pypdf2-to-detect-non-embedded-fonts-in-pdf-file-generated-by-google-docs/60895334#60895334 
        if type(obj) == PyPDF2.generic.ArrayObject:  # You can also do ducktyping here
            for i in obj:
                if hasattr(i, 'keys'):
                    f, e = get_font(i, fonts, embedded_fonts)
                    fonts = fonts.union(f)
                    embedded = embedded.union(e)
        else:
            f, e = get_font(obj['/Resources'], fonts, embedded)
            fonts = fonts.union(f)
            embedded = embedded.union(e)

    unembedded = fonts - embedded
    print('Font List')
    pprint(sorted(list(fonts)))
    if unembedded:
        print('\nUnembedded Fonts')
        pprint(unembedded)

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
    outlines = document.get_outlines()
    for (level,title,dest,a,se) in outlines:
        print (level, title)

get_ToC('S-001_2018E.pdf')