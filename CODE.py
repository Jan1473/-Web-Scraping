import requests
from bs4 import BeautifulSoup
import nltk
from nltk.corpus import stopwords
import string
import re
import csv
import pandas as pd

file_paths = ["C:\\Users\\janan\\OneDrive\\Desktop\\NLP\\StopWords_Auditor.txt", "C:\\Users\\janan\\OneDrive\\Desktop\\NLP\\StopWords_Currencies.txt", "C:\\Users\\janan\\OneDrive\\Desktop\\NLP\\StopWords_DatesandNumbers.txt","C:\\Users\\janan\\OneDrive\\Desktop\\NLP\\StopWords_Generic.txt","C:\\Users\\janan\\OneDrive\\Desktop\\NLP\\StopWords_GenericLong.txt","C:\\Users\\janan\\OneDrive\\Desktop\\NLP\\StopWords_Geographic.txt","C:\\Users\\janan\\OneDrive\\Desktop\\NLP\\StopWords_Names.txt"]
output_folder = "C:\\Users\\janan\\OneDrive\\Desktop\\NLP"
output_file = "StopWords.txt"
output_path = f"{output_folder}/{output_file}"
with open(output_path, "w") as combined_file:
    for file_path in file_paths:
        # Open each input text file in read mode
        with open(file_path, "r") as input_file:
            # Read the content of the input file and write it to the output file
            content = input_file.read()
            combined_file.write(content)
            combined_file.write("\n")  # Optionally, add a newline between the content of each file



def scrape_text_from_website(url):
    response = requests.get(url)
    if response.status_code == 200:
        soup = BeautifulSoup(response.text, "html.parser")
        p_elements = soup.find_all("p")
        tex = ""
        for p in p_elements:
            tex += p.get_text() + "\n"
        lines=tex.split('\n')
        text='\n'.join(lines[17:len(lines)-4])
        return text
    else:
        return None

content=scrape_text_from_website("https://insights.blackcoffer.com/rise-of-telemedicine-and-its-impact-on-livelihood-by-2040-3-2/")


def process_text(text, custom_stopwords):
    nltk.download('stopwords')
    stop_words = set(stopwords.words('english'))
    #punctuation = set(string.punctuation)
    
    words = nltk.word_tokenize(text)
    filtered_text = [word for word in words if word.lower() not in stop_words]
    
    text_final = ' '.join(filtered_text)
    
    with open(custom_stopwords, 'r') as stopwords_file:
        custom_stopwords = set(word.strip() for word in stopwords_file)

    words = text_final.split()
    text_final = ' '.join(word for word in words if word.lower() not in custom_stopwords)
    
    return text_final

result=process_text(content,'C:\\Users\\janan\\OneDrive\\Desktop\\NLP\\StopWords.txt')

def calculate_polarity_subjectivity(text_final, positive_words_file, negative_words_file):
    with open(positive_words_file, 'r') as positive_words:
        pw = set(word.strip() for word in positive_words)
    w = text_final.split()
    positive = [word for word in w if word.lower() in pw]

    with open(negative_words_file, 'r') as negative_words:
        nw = set(word.strip() for word in negative_words)
    w = text_final.split()
    negative = [word for word in w if word.lower() in nw]

    POSITIVE_SCORE = len(positive)
    NEGATIVE_SCORE = len(negative)
    Polarity_Score = (POSITIVE_SCORE - NEGATIVE_SCORE) / (POSITIVE_SCORE + NEGATIVE_SCORE + 0.000001)
    Subjectivity_Score = (POSITIVE_SCORE + NEGATIVE_SCORE) / (len(text_final.split()) + 0.000001)

    return  POSITIVE_SCORE, NEGATIVE_SCORE,Polarity_Score, Subjectivity_Score

# Function to count complex words, average sentence length, and fog index
def calculate_complexity(text_final):
    def count_syllables(word):
        word = word.lower()
        vowels = "aeiouy"
        count = 0

        if word.endswith("es"):
            word = word[:-2]
        elif word.endswith("ed"):
            word = word[:-2]

        if len(word) == 0:
            return 0

        if word[0] in vowels:
            count += 1
        for i in range(1, len(word)):
            if word[i] in vowels and word[i - 1] not in vowels:
                count += 1

        return count

    words = text_final.split()
    syllable_counts = {}
    complex_words = 0

    for word in words:
        syllable_count = count_syllables(word)
        syllable_counts[word] = syllable_count
        if syllable_count > 2:
            complex_words += 1

    sentence_pattern = r'(?<=[.!?]) +'
    sentences = re.split(sentence_pattern, text_final)
    sentences = [sentence.strip() for sentence in sentences if sentence.strip()]
    length_of_sentences = len(sentences)
    if length_of_sentences == 0:
        avg_sen_len = 0  # Set average sentence length to zero when there are no sentences
    else:
        avg_sen_len = len(text_final.split()) / length_of_sentences
        
    if len(text_final.split()) == 0:
        per_com_words = 0  # Set average sentence length to zero when there are no sentences
    else:
        per_com_words = complex_words / len(text_final.split())
    
    
    fog_index = 0.4 * (avg_sen_len + per_com_words)
    avg_no_of_words = avg_sen_len

    return complex_words, avg_sen_len, fog_index, avg_no_of_words

# Function to calculate average word length
def calculate_average_word_length(text_final):
    text = text_final.translate(str.maketrans('', '', string.punctuation))
    words = text.split()
    total_characters = sum(len(word) for word in words)
    total_words = len(words)
    average_word_length = total_characters / total_words
    return average_word_length



def count_personal_pronouns(text):
    pronoun_pattern = r'\b(?!(US\b))((I|we|my|ours|us)\b)\b'
    personal_pronouns = re.findall(pronoun_pattern, text, flags=re.IGNORECASE)
    pronoun_count = len(personal_pronouns)
    return pronoun_count


# Main function to process the text and save the results to a CSV file
def process_and_save_results(url, custom_stopwords, positive_words_file, negative_words_file):
    text = scrape_text_from_website(url)
    if text is None:
        return "Failed to retrieve the web page."
    text_final = process_text(text, custom_stopwords)

    pos_score, neg_score, polarity, subjectivity = calculate_polarity_subjectivity(text_final, positive_words_file, negative_words_file)
    complex_words, avg_sen_len, fog_index, avg_word_length = calculate_complexity(text_final)
    pronoun_count = count_personal_pronouns(text_final)

    return polarity, subjectivity, pos_score, neg_score, complex_words, avg_sen_len, fog_index, avg_word_length, pronoun_count

input_excel_file = "C:\\Users\\janan\\OneDrive\\Desktop\\NLP\\Input.xlsx"  # Replace with your Excel file path
output_csv_file = "C:\\Users\\janan\\OneDrive\\Desktop\\NLP\\text_analysis_results.csv"
custom_stopwords = "C:\\Users\\janan\\OneDrive\\Desktop\\NLP\\StopWords.txt"
positive_words_file = "C:\\Users\\janan\\OneDrive\\Desktop\\NLP\\positive-words.txt"
negative_words_file = "C:\\Users\\janan\\OneDrive\\Desktop\\NLP\\negative-words.txt"



def process_urls_from_excel(input_excel_file, output_csv_file, custom_stopwords, positive_words_file, negative_words_file):
    # Read URLs from the Excel file using pandas
    data = pd.read_excel(input_excel_file)
    urls = data['URL'].tolist()

    # Create a list to store results for each URL
    results = []

    for url in urls:
        result = process_and_save_results(url, custom_stopwords, positive_words_file, negative_words_file)
        if isinstance(result, tuple):
            results.append([url] + list(result))
        else:
            print(f"Error processing URL: {url}")

    # Save the results to a CSV file
    with open(output_csv_file, mode="w", newline="") as file:
        writer = csv.writer(file)
        writer.writerow(["URL", "Polarity", "Subjectivity", "Positive Score", "Negative Score", "Complex Words",
                         "Avg Sentence Length", "Fog Index", "Average Word Length", "Pronoun Count"])
        for result in results:
            writer.writerow(result)

    print(f"Results saved to {output_csv_file}.")



process_urls_from_excel(input_excel_file, output_csv_file, custom_stopwords, positive_words_file, negative_words_file)























