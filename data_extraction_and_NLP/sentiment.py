import pandas as pd
from newspaper import Article
import os
from textblob import TextBlob
import syllables
import nltk
# nltk.download('punkt')
# nltk.download('averaged_perceptron_tagger')
# nltk.download('brown')
print("downloading cmudict...")
nltk.download('cmudict')
print("DONE")

from nltk.corpus import cmudict
d = cmudict.dict()

    

def download_articles(filename):
    df = pd.read_excel(filename)
    url_ids = df['URL_ID']
    urls  = df['URL']
    os.makedirs("./Articles", exist_ok=True)
    print("Downloading Articles")

    for i in range(len(urls)):
        try:
            a = Article(urls[i])
            a.download()
            a.parse()
            # ./Articles/37.txt 
            filename = "./Articles/" + str(url_ids[i]) +".txt"
            with open(filename, "w+") as f:
                f.write(a.text)
                f.close()
            print(f"Article {i} Done ")
        
        except:
            print(f"Article {i} Failed ")




# function to create a list of custum words from files
def create_wordset(file_list):
    wordset = set()

    for file_name in file_list:
        with open(file_name, 'r') as f:
            for line in f:
                line_list  = line.split("|")
                for word in line_list:
                    wordset.add(word.strip())
    return wordset

#Function to remove stopwords from a text
def remove_stopwords(text, stopwords=nltk.corpus.stopwords.words('english')):
    words = nltk.word_tokenize(text)
    # words = [word for word in words if word not in stopwords]
    clean_words = []
    for word in words:
        if word not in stop_words:
            clean_words.append(clean_words)
            
    return ' '.join(words)

# Function to calculate positive words in a text
def count_positive_words(text, positive_words=[]):
    count = 0
    for word in text.split():
        if word.lower() in positive_words:
            count += 1
    return count

# Function to calculate negative words in a text
def count_negative_words(text, negative_words=[]):
    count = 0
    for word in text.split():
        if word.lower() in negative_words:
            count += 1
    return count

# Function to calculate polarity score of a text
def count_polarity_score(text, positive_score, negative_score):
    return (positive_score - negative_score) / (positive_score + negative_score) + 0.000001

# Function to calculate the subjectivity score of a text
def count_subjectivity_score(text, positive_score, negative_score, number_of_words):
    return (positive_score + negative_score) / number_of_words + 0.000001

# Function to calculate the number of sentences in a text
def count_sentences(text):
    return len(nltk.sent_tokenize(text))

# Function to calculate the number of words in a text
def count_words(text):
    return len(nltk.word_tokenize(text))

# Function to calculate the number of syllables in a word
def count_syllables(word):
    if word.lower() not in d:
        return 0
    elif word[-2:] in ['es', 'ed']:
        return len([syl for syl in d[word.lower()][0] if syl[-1].isdigit()]) - 1
    else:
        return len([syl for syl in d[word.lower()][0] if syl[-1].isdigit()])

# Function to calculate the number of complex words in a text
def count_complex_words(text):
    count = 0
    for word in nltk.word_tokenize(text):
        if count_syllables(word) >= 2:
            count += 1
    return count

# Function to calculate the number of personal pronouns in a text
def count_personal_pronouns(text):
    personal_pronouns = set(["i", "we", "my", "ours", "us"])
    count = 0
    for word, tag in nltk.pos_tag(nltk.word_tokenize(text)):
        if tag == 'PRP' and word.lower() in personal_pronouns:
            count += 1
    return count

# Function to calculate the average word length in a text
def avg_word_length(text):
    return sum(len(word) for word in nltk.word_tokenize(text)) / len(nltk.word_tokenize(text))

# Function to calculate the average number of words per sentence in a text
def avg_words_per_sentence(text):
    return len(nltk.word_tokenize(text)) / len(nltk.sent_tokenize(text))


def count_syllables_in_text(text):
    count = 0
    for word in text.split():
        count += count_syllables(word)
    return count



# Function to perform textual analysis on a given text and return a dictionary of variables
def analyze_text(text,positive_words, negative_words, stopwords=nltk.corpus.stopwords.words('english')):
    print("Removing Stopwords...")
    text = remove_stopwords(text, stopwords)  
    print("Done")
    print("Counting number of sentences...")
    number_of_sentences = count_sentences(text)
    print("Counting number of words...")
    number_of_words = count_words(text)
    print("Counting number of Complex words...")
    complex_word_count = count_complex_words(text)
    print("Calculating positive score...")
    positive_score = count_positive_words(text, positive_words)
    print("Calculating negative score...")
    negative_score = count_negative_words(text, negative_words)
    print("Calculating polarity score...")
    polarity_score = count_polarity_score(text, positive_score, negative_score)
    print("Calculating subjectivity score...")
    subjectivity_score = count_subjectivity_score(text, positive_score, negative_score, number_of_words)
    print("Calculating average sentence length...")
    avg_sentence_length = number_of_words / number_of_sentences
    print("Calculating percentage of complex words...")
    percent_complex_words = complex_word_count / number_of_words
    print("Calculating Fog Index...")
    fog_index = 0.4 * (avg_sentence_length + percent_complex_words)
    print("Counting personal pronouns...")
    personal_pronouns = count_personal_pronouns(text)
    print("Calculating average word length...")
    avg_word_length_value = avg_word_length(text)
    print("Calculating syllables per word...")
    syllables_per_word = count_syllables_in_text(text) / number_of_words
    print("Calculating average number of words per sentence...")
    avg_words_per_sentence = number_of_words / number_of_sentences



    return {'POSITIVE SCORE': positive_score,
            'NEGATIVE SCORE': negative_score,
            'POLARITY SCORE': polarity_score,
            'SUBJECTIVITY SCORE': subjectivity_score,
            'AVG SENTENCE LENGTH': avg_sentence_length,
            'PERCENTAGE OF COMPLEX WORDS': percent_complex_words,
            'FOG INDEX': fog_index,
            'AVG NUMBER OF WORDS PER SENTENCE': avg_words_per_sentence,
            'COMPLEX WORD COUNT': complex_word_count,
            'SYLLABLE PER WORD': syllables_per_word,
            'PERSONAL PRONOUNS': personal_pronouns,
            'AVG WORD LENGTH': avg_word_length_value}








download_articles("Input.xlsx")
df = pd.read_excel("Input.xlsx")
url_ids = df['URL_ID']
df1 = pd.read_excel('Output Data Structure.xlsx')
file_list_stopwords = ["StopWords\\StopWords_Auditor.txt", "StopWords\\StopWords_Currencies.txt", "StopWords\\StopWords_DatesandNumbers.txt", "StopWords\\StopWords_Generic.txt", "StopWords\\StopWords_GenericLong.txt", "StopWords\\StopWords_Geographic.txt", "StopWords\\StopWords_Names.txt"]
file_list_positive = ["MasterDictionary\\positive-words.txt"]
file_list_negative = ["MasterDictionary\\negative-words.txt"]

stop_words = create_wordset(file_list_stopwords)
positive_words = create_wordset(file_list_positive)
negative_words = create_wordset(file_list_negative)

for i in range(len(url_ids)):
    print(f"==================== Summarizing Article {i+1} ====================")
    filename = f"./Articles/{url_ids[i]}.txt"
    try:
        with open(filename, "r") as f:
            result = analyze_text(f.read(),positive_words,negative_words, stop_words)
            f.close()
    except:
        result = {'POSITIVE SCORE': "NaN",
            'NEGATIVE SCORE': "NaN",
            'POLARITY SCORE': "NaN",
            'SUBJECTIVITY SCORE': "NaN",
            'AVG SENTENCE LENGTH': "NaN",
            'PERCENTAGE OF COMPLEX WORDS': "NaN",
            'FOG INDEX': "NaN",
            'AVG NUMBER OF WORDS PER SENTENCE': "NaN",
            'COMPLEX WORD COUNT': "NaN",
            'SYLLABLE PER WORD': "NaN",
            'PERSONAL PRONOUNS': "NaN",
            'AVG WORD LENGTH': "NaN"}
        

    df1.loc[i,result.keys()] = result
    print(f"==================== Done Summarizing Article {i+1} ====================")
df1.to_excel('Output Data Structure.xlsx', index=False)

