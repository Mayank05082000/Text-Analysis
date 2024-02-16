##DATA ANALYSIS

import os
import pandas as pd
import nltk
from nltk.tokenize import word_tokenize, sent_tokenize
from nltk.corpus import stopwords
from nltk.probability import FreqDist
from nltk.tag import pos_tag
from nltk.sentiment import SentimentIntensityAnalyzer
from nltk.tokenize.treebank import TreebankWordDetokenizer

nltk.download('punkt')
nltk.download('stopwords')
nltk.download('averaged_perceptron_tagger')
nltk.download('vader_lexicon')

input_excel_path = 'Input.xlsx'
df_input = pd.read_excel(input_excel_path)

stop_words = set()
stopwords_folder = 'StopWords'

for file_name in os.listdir(stopwords_folder):
    with open(os.path.join(stopwords_folder, file_name), 'rb') as file:
        content = file.read()

    try:
        decoded_content = content.decode('utf-8')
    except UnicodeDecodeError:

        decoded_content = content.decode('ansi')  

    stop_words.update(decoded_content.splitlines())

    positive_words_path = 'MasterDictionary/positive-words.txt'
negative_words_path = 'MasterDictionary/negative-words.txt'

# Read the positive words 
with open(positive_words_path, 'r', encoding='utf-8-sig') as file:
    positive_words = set(file.read().splitlines())

# Read the negative words 
with open(negative_words_path, 'r', encoding='latin-1') as file:
    negative_words = set(file.read().splitlines())


def syllable_count(word):
    count = 0
    vowels = "aeiou"
    word = word.lower()

    # Check if the word is a single-letter word
    if len(word) == 1:
        return 1

    # Check the first letter of the word
    if word[0] in vowels:
        count += 1

    # Loop through the rest of the letters in the word
    for i in range(1, len(word)):
        # Check if the current letter is a vowel and the previous letter is not a vowel
        if word[i] in vowels and word[i - 1] not in vowels:
            count += 1

    # Exclude syllable count for words ending with 'e'
    if word.endswith("e") and word[-2] not in vowels:
        count -= 1

    return count


# Function to calculate derived variables
def calculate_derived_variables(text):
    words = word_tokenize(text)
    cleaned_words = [word.lower() for word in words if word.isalpha() and word.lower() not in stop_words]

    # Positive Score
    positive_score = sum(1 for word in cleaned_words if word in positive_words)

    # Negative Score
    negative_score = sum(1 for word in cleaned_words if word in negative_words)

    # Polarity Score
    polarity_score = (positive_score - negative_score) / ((positive_score + negative_score) + 0.000001)

    # Subjectivity Score
    subjectivity_score = (positive_score + negative_score) / (len(cleaned_words) + 0.000001)

    # Average Sentence Length
    sentences = sent_tokenize(text)
    avg_sentence_length = len(words) / len(sentences)

    # Percentage of Complex Words
    complex_word_count = sum(1 for word in cleaned_words if syllable_count(word) > 2)
    percentage_complex_words = complex_word_count / len(cleaned_words)

    # Fog Index
    fog_index = 0.4 * (avg_sentence_length + percentage_complex_words)

    # Average Number of Words Per Sentence
    avg_words_per_sentence = len(words) / len(sentences)

    # Complex Word Count
    complex_word_count = sum(1 for word in cleaned_words if syllable_count(word) > 2)

    # Word Count
    word_count = len(cleaned_words)

    # Syllable Per Word
    syllable_per_word = sum(syllable_count(word) for word in cleaned_words) / word_count

    # Personal Pronouns
    personal_pronouns = sum(1 for word in cleaned_words if word.lower() in ['i', 'we', 'my', 'ours', 'us'])

    # Average Word Length
    avg_word_length = sum(len(word) for word in cleaned_words) / word_count

    return [positive_score, negative_score, polarity_score, subjectivity_score, avg_sentence_length,
            percentage_complex_words, fog_index, avg_words_per_sentence, complex_word_count, word_count,
            syllable_per_word, personal_pronouns, avg_word_length]


# Function to perform analysis on each article text
def analyze_text(folder_path):
    # Create an empty DataFrame to store the results
    result_df = pd.DataFrame(columns=[
        'POSITIVE SCORE', 'NEGATIVE SCORE', 'POLARITY SCORE', 'SUBJECTIVITY SCORE',
        'AVG SENTENCE LENGTH', 'PERCENTAGE OF COMPLEX WORDS', 'FOG INDEX',
        'AVG NUMBER OF WORDS PER SENTENCE', 'COMPLEX WORD COUNT', 'WORD COUNT',
        'SYLLABLE PER WORD', 'PERSONAL PRONOUNS', 'AVG WORD LENGTH'
    ])
    
    # Iterate over all files in the folder
    for filename in os.listdir(folder_path):
        file_path = os.path.join(folder_path, filename)

        # Check if the item is a file
        if os.path.isfile(file_path):
            with open(file_path, 'r', encoding='utf-8') as file:
                text = file.read()

            derived_variables = calculate_derived_variables(text)

            # Append the results to the DataFrame
            result_df.loc[len(result_df)] = derived_variables


    return result_df


# Call the analyze_text function and store the result
output_folder_path = 'output_file'
result_df = analyze_text(output_folder_path)

# Merge the result_df with the original input DataFrame
df_final = pd.merge(df_input, result_df, left_index=True, right_index=True)

# Save the final DataFrame to an output Excel file
output_excel_path = 'Output_Data.xlsx'
df_final.to_excel(output_excel_path, index=False)
print(f"Output saved to {output_excel_path}")