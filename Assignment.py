from bs4 import BeautifulSoup
import requests
import pandas as pd
import os
import re
import nltk
from nltk.tokenize import word_tokenize, sent_tokenize

data = pd.read_excel("G:\Blackcoffer Assignment\Input.xlsx")

output_folder = "Extracted_Text"

# Creating a Extracted_Text Folder
if not os.path.exists(output_folder):
    os.makedirs(output_folder)

# Function to extract article text from URL
def extract_article_text(url):
    try:
        response = requests.get(url)
        soup = BeautifulSoup(response.text, 'html.parser')

        # Extracting title and article text
        title = soup.title.text if soup.title else ''
        title = title.replace('-', '')
        paragraphs = soup.find_all('p')
        article_text = ' '.join([p.get_text() for p in paragraphs])

        # Removing extra spaces and special characters
        article_text = re.sub(r'\s+', ' ', article_text).strip()

        return title, article_text
    except Exception as e:
        print(f"Error extracting text from {url}: {str(e)}")
        return None, None
    

# Main loop for data extraction
for index, row in data.iterrows():
    url_id = row['URL_ID']
    url = row['URL']

    # Extracting article text and title for all the websites
    title, article_text = extract_article_text(url)
    if article_text is not None:
        # Saving extracted content to a text file
        file_name = os.path.join(output_folder, f"{url_id}.txt")
        with open(file_name, 'w', encoding='utf-8') as file:
            file.write(f"{title}\n")
            file.write(f"{article_text}")

print("Data extraction completed. Text files are saved in the 'Extracted_Text' folder.")

# Downloading nltk resources
try:
    nltk.data.find('tokenizers/punkt')
except LookupError:
    nltk.download('punkt')

# Postive and negative word files
positive_words_file = "positive-words.txt"
negative_words_file = "negative-words.txt"

# stop words files
stop_words_files = [
    "StopWords_Auditor.txt",
    "StopWords_Currencies.txt",
    "StopWords_DatesandNumbers.txt",
    "StopWords_Generic.txt",
    "StopWords_GenericLong.txt",
    "StopWords_Geographic.txt",
    "StopWords_Names.txt"
]

# Reading positive and negative words

positive_words = set()
negative_words = set()

masterdictionary= 'MasterDictionary'

positive_words_file_path = os.path.join(masterdictionary, positive_words_file)
with open(positive_words_file_path, 'r') as file:
    positive_words.update(file.read().splitlines())

negative_words_file_path = os.path.join(masterdictionary, negative_words_file)
with open(negative_words_file_path, 'r') as file:
    negative_words.update(file.read().splitlines())

# Reading stop words
additional_stop_words = set()
stop_words_folder = "StopWords"

for stop_words_file in stop_words_files:
    stop_words_file_path = os.path.join(stop_words_folder, stop_words_file)

    with open(stop_words_file_path, 'r', encoding='latin-1') as file:
        additional_stop_words.update(file.read().splitlines())

# Creating output dataframe
output_columns = ["URL_ID", "URL", "Positive Score", "Negative Score", "Polarity Score",
                  "Subjectivity Score", "Average Sentence Length", "Percentage of Complex Words",
                  "Fog Index", "Average Number of Words Per Sentence", "Complex Word Count",
                  "Word Count", "Syllable Per Word", "Personal Pronouns", "Average Word Length"]

output_data = pd.DataFrame(columns=output_columns)


def syllable_count(word):
    # Basic syllable counting
    count = 0
    vowels = "aeiou"
    word = word.lower().strip(".:;?!")

    # Handling exceptions for words ending with "es" or "ed"
    if word.endswith("es") or word.endswith("ed"):
        pass
    else:
        if word[0] in vowels:
            count += 1
        for index in range(1, len(word)):
            if word[index] in vowels and word[index - 1] not in vowels:
                count += 1
        if word.endswith("e"):
            count -= 1
        if count == 0:
            count += 1

    return count


# Text Analysis
def analyze_text(row):
    url_id = row["URL_ID"]
    url = row["URL"]

    output_folder = 'Extracted_Text'
    # Extracted text file path
    text_file_path = os.path.join(output_folder, f"{url_id}.txt")

    # Read the extracted text file
    with open(text_file_path, 'r', encoding='utf-8') as text_file:
        text = text_file.read()

    # Tokenization and Cleaning
    words = [word.lower() for word in word_tokenize(text) if word.isalpha() and word.lower() not in additional_stop_words]
    sentences = sent_tokenize(text)

    # Sentiment Analysis
    positive_score = sum(1 for word in words if word in positive_words)
    negative_score = sum(1 for word in words if word in negative_words)
    polarity_score = (positive_score - negative_score) / ((positive_score + negative_score) + 0.000001)
    subjectivity_score = (positive_score + negative_score) / (len(words) + 0.000001)

    # Readability Analysis
    avg_sentence_length = len(words) / len(sentences)
    complex_word_count = sum(1 for word in words if syllable_count(word) > 2)
    percentage_complex_words = complex_word_count / len(words)
    fog_index = 0.4 * (avg_sentence_length + percentage_complex_words)
    avg_number_of_words_per_sentence = len(words) / len(sentences)
    
    word_count = len(words)

    # Syllable Count Per Word
    syllable_per_word = sum(syllable_count(word) for word in words) / len(words)

    # Personal Pronouns
    personal_pronouns_count = len(re.findall(r'\b(?:I|we|my|ours|us)\b', text, flags=re.IGNORECASE))

    # Average Word Length
    avg_word_length = sum(len(word) for word in words) / len(words)

    return pd.Series({
        "URL_ID": url_id,
        "URL": url,
        "Positive Score": positive_score,
        "Negative Score": negative_score,
        "Polarity Score": polarity_score,
        "Subjectivity Score": subjectivity_score,
        "Average Sentence Length": avg_sentence_length,
        "Percentage of Complex Words": percentage_complex_words,
        "Fog Index": fog_index,
        "Average Number of Words Per Sentence": avg_number_of_words_per_sentence,
        "Complex Word Count": complex_word_count,
        "Word Count": word_count,
        "Syllable Per Word": syllable_per_word,
        "Personal Pronouns": personal_pronouns_count,
        "Average Word Length": avg_word_length
    })


# Applying the text analysis function to each row
output_data = data.apply(analyze_text, axis=1)

# Storing the results to the output file
output_file = "output.xlsx"
output_data.to_excel(output_file, index=False)

print(f"Text analysis results saved to {output_file}")