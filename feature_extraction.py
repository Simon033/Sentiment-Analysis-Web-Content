#!/usr/bin/env python
# coding: utf-8

# Importing necessary packages

import pandas as pd
import requests
from bs4 import BeautifulSoup
import nltk
from nltk.tokenize import word_tokenize
import re
import os

# Downloading NLTK data package
nltk.download('punkt')


# Read the Input file into a pandas DataFrame
input_df = pd.read_excel('Input.xlsx')

# Sending requests to all URL
error_urls = []

for index, row in input_df.iterrows():
    url = row['URL']
    url_id = row['URL_ID']

    # Making a request to URL with user agent
    header = {'User-Agent': "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/108.0.0.0 Safari/537.36"}
    try:
        response = requests.get(url,headers=header)
    except:
        error_urls.append(url_id)
        print("Error getting response of {}".format(url_id))
        continue
        
    try:
        #Create soup
        soup = BeautifulSoup(response.content, 'html.parser')
        
        #get title
        title = soup.find('h1').get_text()
        
        #get the article text
        article = ""
        for p in soup.find_all('p'):
            article += p.get_text()
    except Exception as e:
        error_urls.append(url_id)
        print(f"Error processing {url_id}: {str(e)}")
        continue
    
    # Write title and text to a file
    file_name = 'C:\\Users\\simon\\Python Data\\projects\\interview\\Blackcoffer\\URL Text\\' + str(url_id) + '.txt'
    with open(file_name, 'w',encoding='utf-8') as file:
        file.write(title + '\n' + article)


# directories
text_dir = "C:\\Users\\simon\\Python Data\\projects\\interview\\Blackcoffer\\URL Text"
stopwords_dir = "C:\\Users\\simon\\Python Data\\projects\\interview\\Blackcoffer\\StopWords"
sentiment_dir = "C:\\Users\\simon\\Python Data\\projects\\interview\\Blackcoffer\\MasterDictionary"

# Load all stop words into a set variable
stop_words = set()
for files in os.listdir(stopwords_dir):
    with open(os.path.join(stopwords_dir, files), 'r', encoding='ISO-8859-1') as f:
        stop_words.update(set(f.read().splitlines()))


# Process and load all text files into a list
docs = []
for text_file in os.listdir(text_dir):
    with open(os.path.join(text_dir, text_file), 'r',encoding='ISO-8859-1') as f:
        text = f.read()
        # Tokenize
        words = word_tokenize(text)
        # Remove stop words
        filtered_text = [word for word in words if word.lower() not in stop_words]
        # Append to list
        docs.append(filtered_text)

        
# Read positive and negative files from directory
pos = set()
neg = set()
for files in os.listdir(sentiment_dir):
    if files == 'positive-words.txt':
        with open(os.path.join(sentiment_dir, files), 'r', encoding='ISO-8859-1') as f:
            pos.update(f.read().splitlines())
    elif files == 'negative-words.txt':
        with open(os.path.join(sentiment_dir, files), 'r', encoding='ISO-8859-1') as f:
            neg.update(f.read().splitlines())


# Extracting Variables

# calculate positive and negative score
positive_words = [[word.lower() for word in doc if word.lower() in pos] for doc in docs]
negative_words = [[word.lower() for word in doc if word.lower() in neg] for doc in docs]

positive_score = [len(words) for words in positive_words]
negative_score = [len(words) for words in negative_words]

# calculate polarity and subjectivity score
polarity_score = [(pos_score - neg_score) / (pos_score + neg_score + 0.000001) for pos_score, neg_score in zip(positive_score, negative_score)]
subjectivity_score = [(pos_score + neg_score) / (len(doc) + 0.000001) for doc, pos_score, neg_score in zip(docs, positive_score, negative_score)]


# Fucntion to Calculate other NLP metrics
def calculate_metric(file):
    with open(os.path.join(text_dir, file), 'r',encoding='ISO-8859-1') as f:
        text = f.read()
        
        # Remove punctuations and split into sentences
        text = re.sub(r'[^\w\s.]', '', text)
        sentences = text.split('.')
        
        num_sentences = len(sentences)
        
        # Total words
        words = [word for word in text.split() if word.lower() not in stop_words]
        num_words = len(words)
        
        # complex words
        vowels = 'aeiou'
        complex_words = [word for word in words if sum(1 for letter in word if letter.lower() in vowels) > 2]
        
        # syllable count
        vowels = 'aeiou'
        syllable_count = 0
        syllable_words = []
        
        for word in words:
            # Skip the entire word if it ends with "es" or "ed"
            if word.endswith(('es', 'ed')):
                continue  
                
            syllable_count_word = sum(1 for letter in word if letter.lower() in vowels)
            
            if syllable_count_word >= 1:
                syllable_words.append(word)
                syllable_count += syllable_count_word
                
        # sentence length
        avg_sentence_len = num_words / num_sentences
        
        # syllable count per word
        avg_syllable_word_count = syllable_count / len(syllable_words)
        
        # percent complex words
        percent_complex_words = len(complex_words) / num_words
        
        # fog index
        fog_index = 0.4 * (avg_sentence_len + percent_complex_words)
        
        # Word Count and Average Word Length
        word_count = len(words)
        average_word_length = sum(len(word) for word in words) / word_count
        
        return avg_sentence_len, percent_complex_words, fog_index, len(complex_words), avg_syllable_word_count, word_count, average_word_length


# calculating metrics
metrics = [calculate_metric(file) for file in os.listdir(text_dir)]
avg_sentence_length, percentage_of_complex_words, fog_index, complex_word_count, avg_syllable_word_count, word_count, average_word_length = zip(*metrics)


# Since both Avg Sentence Length and Number of Words per sentence evaluate the same
avg_words_per_sentence=avg_sentence_length


# Count Personal Pronouns
personal_pronouns = set(["I", "we", "my", "ours", "us"])  # Convert to a set for faster membership checks

pronoun_counts = []

for file in os.listdir(text_dir):
    with open(os.path.join(text_dir, file), 'r', encoding='ISO-8859-1') as f:
        text = f.read()
        file_pronoun_count = sum(len(re.findall(r"\b" + pronoun + r"\b", text.lower()))for pronoun in personal_pronouns)
        
        # Exclude "US" as it is refering to the country
        file_pronoun_count -= len(re.findall(r'\bUS\b', text))
        
        pronoun_counts.append(file_pronoun_count)


# Read the output data structure
output_df = pd.read_excel('Output Data Structure.xlsx')


# drop rows that got an error
if error_urls:
    # Filter rows to drop based on URL_IDs in error_urls
    rows_to_drop = output_df[output_df['URL_ID'].isin(error_urls)].index
    
    # Drop the rows from the output_df dataframe
    output_df.drop(rows_to_drop, axis=0, inplace=True)


# output variables
variables = [
    positive_score,
    negative_score,
    polarity_score,
    subjectivity_score,
    avg_sentence_length,  
    percentage_of_complex_words,
    fog_index,
    avg_words_per_sentence,
    complex_word_count,
    word_count,
    avg_syllable_word_count,
    pronoun_counts,
    average_word_length
]

# Write the values into DataFrame
for i, var in enumerate(variables):
    output_df[output_df.columns[i + 2]] = var
    
# Save the DataFrame to a CSV file
output_df.to_csv('Output.csv', index=False) 

if error_urls:
    print("Process Completed with Errors")
else:
    print("Process Completed Sucessfully")