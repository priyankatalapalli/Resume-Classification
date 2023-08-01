import pickle
import streamlit as st
import re
import string
import nltk
import pandas as pd
from nltk.corpus import stopwords
from nltk.stem import PorterStemmer
import spacy
from nltk.tokenize import word_tokenize
from sklearn.feature_extraction.text import TfidfVectorizer
import streamlit.components.v1 as components
import docx2txt
import PyPDF2
import win32com.client as win32
import os
import tempfile

pickle_in = open('classifier.pkl', 'rb')
classifier = pickle.load(pickle_in)
nltk.download('punkt')
nltk.download('stopwords')
stop = stopwords.words('english')
result = string.punctuation
ps = PorterStemmer()
nlp = spacy.load('en_core_web_sm')


# Cleaning the data
def resumecleaning(resume):
    resume = re.sub(r'http\S+', '', resume)
    resume = word_tokenize(resume.lower())
    resume = [word for word in resume if word not in stop and word not in result]
    resume = [ps.stem(word) for word in resume]
    resume = nlp(' '.join(resume))
    resume = [token.lemma_ for token in resume]
    cleaned_text = ' '.join(resume)
    return cleaned_text

def display():
    st.title("Resume - Classifier")

    # File upload section
    uploaded_file = st.file_uploader("Upload your resume", type=["pdf", "docx", "doc"])

    # Checking the type of file was uploaded and extracting it to for further processing
    if uploaded_file:
        with tempfile.NamedTemporaryFile(delete=False) as temp_file:
            temp_file.write(uploaded_file.read())

        if uploaded_file.name.endswith('.docx'):
            resume_text = docx2txt.process(temp_file.name)
        elif uploaded_file.name.endswith('.doc'):
            try:
                word = win32.gencache.EnsureDispatch('Word.Application')
                doc = word.Documents.Open(os.path.abspath(temp_file.name))
                resume_text = doc.Content.Text
                doc.Close()
                word.Quit()
            except Exception as e:
                st.error(f"Error: {e}")
                return
        elif uploaded_file.name.endswith('.pdf'):
            reader = PyPDF2.PdfReader(temp_file.name)
            resume_text = ""
            for page in reader.pages:
                resume_text += page.extract_text()
        else:
            st.warning("Invalid file format. Please upload a PDF or Word document.")
            return
        
        # Load the trained vectorizer
        with open('vectorizer.pkl', 'rb') as vectorizer_file:
            vectorizer = pickle.load(vectorizer_file)
        
        # Cleaning the uploaded resume text
        cleaned = resumecleaning(resume_text)

        # Vectorize the cleaned resume text using the loaded vectorizer
        cleaned_resume_vectorized = vectorizer.transform([cleaned])

        # Make predictions
        predicted_probabilities = classifier.predict_proba(cleaned_resume_vectorized)[0]

        # Define the output classes
        output_classes = ['Peoplesoft', 'React JS Developer', 'SQL', 'Workday']

        # Display the predicted classes with probability greater than 0.5
        # st.write("The classified resume belong to the category of  ")
        for class_name, prediction_value in zip(output_classes, predicted_probabilities):
            if prediction_value > 0.5:
                st.markdown(f"<h3 style='font-size: 24px;'>The uploaded Resume belongs to the category of {class_name} Resumes</h3>", unsafe_allow_html=True)

    else:
        st.warning("Please upload a file.")

if __name__ == '__main__':
    display()
