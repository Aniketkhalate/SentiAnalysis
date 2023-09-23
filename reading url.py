import requests
import os
import re
import urllib
from urllib.request import urlopen
from urllib.error import HTTPError
import pandas as pd
import lxml
from bs4 import BeautifulSoup
import nltk
from nltk.stem import WordNetLemmatizer
from nltk.tokenize import word_tokenize
from nltk.tokenize import sent_tokenize
from nltk.corpus import wordnet
from nltk.corpus import stopwords
from nltk.sentiment.vader import SentimentIntensityAnalyzer
import openpyxl
from openpyxl import load_workbook

#Get the list of websites in an excel file
data1 = pd.read_excel('Input.xlsx')

#Iterate over the list
for i in range(len(data1)):
    f1 = open("{j}.txt".format(j = 'web_content'), "w", encoding="utf-8")
    
    url = data1.loc[i,'URL']
    headers = {
      'User-Agent': 'Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/63.0.3239.132 Safari/537.36 QIHU 360SE'
    }
    
    try:
        f = requests.get(url, headers = headers)
        soup = BeautifulSoup(f.content,"lxml")
        out_title = soup.find('div',{'class':'td-post-header'}).find('h1')
        out_content = soup.find('div',{'class':'td-post-content'}).find_all('p')
    
    except HTTPError as e :
        f1.write("Couldn't find server")
        f1.close()
    
    except AttributeError as er:
       f1.write("Page not found")
       f1.close()
    
    else:
        content = out_title.text.strip() + ' '
                
        for m in out_content:
            content = content + m.text.strip() 
        f1.write(content)
        f1.close()
        
    f1 = open("{j}.txt".format(j = 'web_content'), "r+", encoding="utf-8") 
    
    #Count Syllable function
    def count_syllables(word):
        return len(
            re.findall('(?!e$)[aeiouy]+', word, re.I) +
            re.findall('^[^aeiouy]*e$', word, re.I)
        )

    #Personal Pronoun count function
    def count_personal_pronouns(word):
        return len(
            re.findall('I|we|my|ours|us', word, re.I)
        )
    
    #Sentiment Analyzer function
    def sentiment_analyse(sentiment_text):
        score = SentimentIntensityAnalyzer().polarity_scores(sentiment_text)
        scoress = [score['neg'], score['pos']]
        return scoress
    
    text = f1.read()
    
    workbook = load_workbook(filename="Output.xlsx")
    sheet = workbook.active

    cell = ['C','D','E','F','G','H','I','J','K','L','M','N','O']
    
    
    if text == "Couldn't find server" or text == "Page not found":
        for j in cell:
            sheet[j+'{i}'.format(i=i+2)] = 0
    else:
        #normalization
        text = text.lower()
            
        #removing unicode characters
        clean_text = re.sub(r"(@\[A-Za-z0-9]+)|([^0-9A-Za-z \t])|(\w+:\/\/\S+)|^rt|http.+?", "", text)
                
        #remove stopwords
        filtered_text = []
        stop_words = open('StopWords.txt', 'r', encoding="utf-8")
        for w in clean_text:
            if w not in stop_words:
                filtered_text.append(w)
                
                
        #Tokenization of word of clean text
        tokenized_text = word_tokenize(clean_text, "english")
            
        cal_score = sentiment_analyse(clean_text)
        sheet[cell[0]+'{i}'.format(i=i+2)] = cal_score[1] #positive score
        sheet[cell[1]+'{i}'.format(i=i+2)] = cal_score[0] #negative score
        sheet[cell[2]+'{i}'.format(i=i+2)] = ((cal_score[1]-cal_score[0])/(cal_score[1]-cal_score[0]+0.000001)) #polarity score
        sheet[cell[3]+'{i}'.format(i=i+2)] = ((cal_score[1]+cal_score[0])/(len(tokenized_text)+0.000001)) #subjectivity score
            
        #Tokenization of sentences in text
        tokenized_sentences = sent_tokenize(clean_text,"english")
            
        #Average Sentence Length
        avg_sent_len = len(tokenized_text)/len(tokenized_sentences)
        sheet[cell[4]+'{i}'.format(i=i+2)] = avg_sent_len  
            
        #Count complex words
        complex_words = 0
        for word in tokenized_text:
            g = count_syllables(word)
            if g>2:
                complex_words +=1
            
        #Percentage of complex words
        percent_complex = complex_words/len(tokenized_text)
        sheet[cell[5]+'{i}'.format(i=i+2)] = percent_complex
            
        #FOG Index
        fog = 0.4 *(percent_complex + avg_sent_len)
        sheet[cell[6]+'{i}'.format(i=i+2)] = fog
            
            
        #Average number of words per sentence
        sheet[cell[7]+'{i}'.format(i=i+2)] = len(tokenized_text)/len(tokenized_sentences)
            
            
        sheet[cell[8]+'{i}'.format(i=i+2)] = complex_words

        #Word Count
        sheet[cell[9]+'{i}'.format(i=i+2)] = len(tokenized_text)
        
        #Perform Lemmatization of words
        lemmatizer = WordNetLemmatizer()

        def pos_tagger(nltk_tag):
            if nltk_tag.startswith('J'):
                return wordnet.ADJ
            elif nltk_tag.startswith('V'):
                return wordnet.VERB
            elif nltk_tag.startswith('N'):
                return wordnet.NOUN
            elif nltk_tag.startswith('R'):
                return wordnet.ADV
            else:         
                return None

        #Find the POS tag for each token
        pos_tagged = nltk.pos_tag(filtered_text)

        #We use our own pos_tagger function to make things simpler to understand.
        wordnet_tagged = list(map(lambda x: (x[0], pos_tagger(x[1])), pos_tagged))

        lemmatized_sentence = []
        for word, tag in wordnet_tagged:
            if tag is None:
                # if there is no available tag, append the token as is
                lemmatized_sentence.append(word)
            else:       
                # else use the tag to lemmatize the token
                lemmatized_sentence.append(lemmatizer.lemmatize(word, tag))
        lemmatized_sentence = "".join(lemmatized_sentence)

        

        lemma_tokenize = word_tokenize(lemmatized_sentence,'english')
        
        syllables_per_word = 0
        for wor in lemma_tokenize:
            syllables_per_word += count_syllables(wor)
        
        #Syllable Count
        sheet[cell[10]+'{i}'.format(i=i+2)] = syllables_per_word
        
        #Personal Pronoun
        per_pro = 0
        per_pro += count_personal_pronouns(clean_text)
        sheet[cell[11]+'{i}'.format(i=i+2)] = per_pro
        
        #Average Word Length
        char_count = 0
        for q in tokenized_text:
            char_count = char_count + len(re.findall('\w', q, re.I))
            
        avg_word_length = char_count/len(tokenized_text)
        sheet[cell[12]+'{i}'.format(i= i+2)] = avg_word_length
            
    workbook.save(filename="Output.xlsx")
    f1.close()