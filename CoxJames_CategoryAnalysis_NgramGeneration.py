import openpyxl
import nltk
from nltk.tokenize import word_tokenize
from nltk.probability import FreqDist
from nltk.util import ngrams
import string
import requests
import os

# Get the directory path of the Python script
dir_path = os.path.dirname(os.path.realpath(__file__))

# Load the Excel file in the same directory as the Python script
wb = openpyxl.load_workbook(os.path.join(dir_path, 'INSERT EXCEL FILE NAME'))

# Get the sheet named "Categories"
ws_categories = wb['Categories']

# Get the sheet named "Dictionary"
ws_dictionary = wb['Dictionary']

# Create a dictionary to store the categories and their keywords
categories = {}

# Define the set of stop words
with open(os.path.join(dir_path, 'Stopwords List Expanded.txt'), 'r') as f:
    stopwords_list = f.readlines()
stop_words = set([word.strip() for word in stopwords_list])

# Loop through the rows of the sheet "Categories"
feedback_counts = {}
for row in ws_categories.rows:

    # Get the feedback and the category
    feedback = row[0].value
    category = row[1].value

    # Increment the feedback count for the category
    if category not in feedback_counts:
        feedback_counts[category] = 1
    else:
        feedback_counts[category] += 1

    # If the category does not exist in the dictionary, create it
    if category not in categories:
        categories[category] = []

    # Tokenise the feedback
    tokens = [word.lower() for word in word_tokenize(feedback) if word.isalnum() and not word[0].isdigit()]

    # Remove stop words from the list of tokens
    tokens = [word for word in tokens if word not in stop_words and word.lower() not in stop_words]

    # Remove duplicate words from the list of tokens
    tokens = list(set(tokens))

    # Create 3-grams and 4-grams from the list of tokens
    n = 4
    n_grams = list(ngrams(tokens, n))
    for gram in n_grams:
        categories[category].append(' '.join(gram))
    n = 3
    n_grams = list(ngrams(tokens, n))
    for gram in n_grams:
        categories[category].append(' '.join(gram))

    # Tag the tokens with part-of-speech tags
    tagged_tokens = nltk.pos_tag(tokens)

    # Add the adjectives, adverbs, verbs, and nouns to the list of keywords for the category
    for word, tag in tagged_tokens:
        if tag.startswith('J') or tag.startswith('R') or tag.startswith('V') or tag.startswith('N'):
            categories[category].append(word)

# Loop through the categories
for category in categories:

    # Get the list of keywords for the category
    keywords = categories[category]

    # Create a frequency distribution of the keywords
    fdist = FreqDist(keywords)

    # Get the total number of feedbacks assigned to the category
    total_feedbacks = feedback_counts[category]

    # Get the list of unique keywords sorted by frequency
    unique_keywords = [word for word, freq in sorted(fdist.items(), key=lambda x: x[1], reverse=True) if word not in string.punctuation and word not in stop_words and (len(word.split()) == 3 or len(word.split()) == 4)]

    # Initialize a list to store the common n-grams
    common_ngrams = []

    # Loop through all possible n-grams for the category
    for n in range(2, 5):
        ngrams_list = list(ngrams(keywords, n))
        ngrams_freq = FreqDist(ngrams_list)
        for ngram, freq in ngrams_freq.items():
            if freq/total_feedbacks > 0:
                if n == 5:
                    common_ngrams.append(' '.join(ngram))
                else:
                    common_ngrams.append(' '.join(ngram[:-1]))

    # Modify the list of unique keywords to only include the common n-grams
    unique_keywords = [word for word in unique_keywords if word in common_ngrams]

    # Write the list of unique keywords to the sheet "Dictionary"
    keywords_str = ', '.join(unique_keywords)
    ws_dictionary.append([category, keywords_str])

# Save and close the Excel file
wb.save(os.path.join(dir_path, 'Dictionary_Output.xlsx'))
wb.close()