from flask import Flask, jsonify
import nltk
import re
from pdfminer.high_level import extract_text
import docx2txt
import win32com.client as win32
import os
import json
from win32com.client import constants
# nltk.download('punkt')
# nltk.download('averaged_perceptron_tagger')
# nltk.download('maxent_ne_chunker')
# nltk.download('words')
# nltk.download('stopwords')

app = Flask(__name__)

skills_db = [
    'java',
    'javascript',
    'python',
    'html',
    'css',
    'ajax',
    'flask',
    'sql',
    'machine learning',
    'data analytics',
    'c',
    'php',
    'web api',
    'jquery',
    'word',
    'excel',
    'power point','c sharp',
    'mercurial', 'tortoisehg','dapper','asp dot net', 'vb dot net', 'ado dot net','dot net', 
    'dot net mvc', 'asp dot net mvc','mvc', 'c sharp dot net'
]
ctx = app.app_context()
ctx.push()
@app.route("/")
def hello():
  return "Hello World!" 

@app.route("/info/<string:path>")

def get_info(path):
    
    def SaveAsDocx (path) :
        #Opening MS Word
        word = win32.gencache.EnsureDispatch('Word.Application')
        doc = word.Documents.Open(path)
        doc.Activate()
        #Rename path with .docx
        new_file_abs = os.path.abspath(path)
        new_file_abs = re.sub(r'\.\w+$','.docx', new_file_abs)
        #Save and Close
        word.ActiveDocument.SaveAs (
        new_file_abs , FileFormat = constants.wdFormatXMLDocument
        )
        doc.Close ( False )
        print ('done')
        return new_file_abs

    def take_input(path):
        extension = path[-1:-5:-1][::-1]
        if extension == '.pdf':
            text = extract_text(path)
            #print(pdf_text)
            return text
        elif extension == 'docx':
            text = docx2txt.process(path)
            return text
        elif extension == '.doc':
            doc_to_docx_path = SaveAsDocx(path)
            text = docx2txt.process(doc_to_docx_path)
            return text
        else:
            return None
    
    def get_name(text):
        name_regex = re.compile(r'(([A-Z][a-z]+|[A-Z]+)[\s]([A-Z][a-z]+|[A-Z]+|[a-z]+|([A-Z]?\.))[\s]?([A-Z][a-z]+|[A-Z]+|[a-z]+))')
        #cv_regex = re.compile(r'((C[A-Z]+)[\s](V[A-Z]+))')
        cv_regex = re.compile(r'([C|c]\w+[m]|[M])[\s]([V|v]\w+[ae|AE])')
        name = re.findall(name_regex, text)
        #print(name)
        cv = re.findall(cv_regex,text)
        if re.search(cv_regex, text):
            return name[1][0]
        else:
            return name[0][0]

    def find_phone(text):
        re1 = re.compile(r'\(?\d{3}\)?[-.\s]?\d{3}[-.\s]?\d{4}') 
        re2 = re.compile(r'\(?\d{3}\)?\(?\d{3}\)?\(?\d{4}\)?')
        re3 = re.compile(r'(\+[0-9]{1,3}[\s]?\-?[\s]?[0-9]{10})')
        number1 = re.findall(re1, text)
        if number1:
            phone_no = number1[0]

            if text.find(phone_no):
                return phone_no
            else:
                return None
        number2 = re.findall(re2, text)
        print(number2)
        if number2:
            phone_no = number2[0]

            if text.find(phone_no):
                return phone_no
            else:
                return None
        number3 = re.findall(re3, text)
        print(number3)
        if number3:
            phone_no = number3[0]

            if text.find(phone_no):
                return phone_no
            else:
                return None
    
    def find_email(text):
        reg_ex_email = re.compile(r'[a-z 0-9]+[\._]?[a-z 0-9]+[@]\w+[.]\w{2,3}')
        email = re.findall(reg_ex_email, text)
        if email:
            email_str = ''.join(email[0])
            return email_str
        else:
            return None
    
    def extract_skills(input_text):
        input_list = input_text.split(" ")
        input_list = (map(lambda x: x.lower(), input_list))
        input_list = list(input_list)
        for text in input_list:
            if ".net" in text:
                input_list.append(text.replace(".net"," dot net"))
            if "#" in text:
                input_list.append(text.replace("#"," sharp"))
        input_text = ' '.join(input_list)
        stop_words = set(nltk.corpus.stopwords.words('english'))
        word_tokens = nltk.tokenize.word_tokenize(input_text)

        # remove the stop words
        filtered_tokens = [w for w in word_tokens if w not in stop_words]

        # remove the punctuation
        filtered_tokens = [w for w in word_tokens if w.isalpha()]

        # generate bigrams and trigrams (such as artificial intelligence)
        bigrams_trigrams = list(map(' '.join, nltk.everygrams(filtered_tokens, 2, 4)))
        #print(bigrams_trigrams)
        # we create a set to keep the results in.
        found_skills = list()

        # we search for each token in our skills database
        for token in filtered_tokens:
            if token.lower() in skills_db:
                found_skills.append(token)
    #             found_skills.add(token)


        # we search for each bigram and trigram in our skills database
        for ngram in bigrams_trigrams:
            if ngram.lower() in skills_db:
                found_skills.append(ngram)
    #             found_skills.add(ngram)
        found_skills_occurance = {}
        for i in found_skills:
            if i not in found_skills_occurance:
                found_skills_occurance[i]=1
            else:
                found_skills_occurance[i]+=1
        print(found_skills_occurance)

        return found_skills_occurance

    text = take_input(path)
    name = get_name(text)
    phone = find_phone(text)
    email = find_email(text)
    skills = extract_skills(text)
    
    result = {
        'Name': name,
        'Phone Number': phone,
        'Email': email,
        'Skills': skills
    }
    return jsonify(result)
path = "Add path where resume is stored" #Add path of where resume is stored
print(get_info(path))
ctx.pop()
if __name__ == "__main__":
   app.run(debug=True)