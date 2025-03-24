import pandas as pd
import os
import re

# File paths 
excel_path = "...."
files_directory = "..."
dictionary_file = "..."
stop_words_directory = "..."
neg_dictionary_file = "..."

# Load stp words, positive, and negative dictionaries 
def load_stop_words(directory):
    stop_words = set()
    for file in os.listdir(directory):
        file_path = os.path.join(directory, file)
        if os.path.isfile(file_path):
            try:
                with open(file_path, "r", encoding="utf-8", errors="ignore") as f:
                    stop_words.update(word.strip().lower() for word in f.readlines())
            except Exception as e:
                print(f"Error reading stop words from {file}: {e}")
    return stop_words

stop_words = load_stop_words(stop_words_directory)

with open(dictionary_file, "r", encoding="utf-8", errors="ignore") as f:
    dictionary_words = set(word.strip().lower() for word in f.readlines())

with open(neg_dictionary_file, "r", encoding="utf-8", errors="ignore") as f:
    neg_dictionary_words = set(word.strip().lower() for word in f.readlines())

df = pd.read_excel(excel_path, index_col=0)

# ppositive and negative score functions 
def count_positive_words(file_path, dictionary_words, stop_words):
    try:
        with open(file_path, "r", encoding="utf-8", errors="ignore") as f:
            text = f.read().lower()
            words = set(re.findall(r'\b\w+\b', text))
            words -= stop_words
            return sum(1 for word in words if word in dictionary_words)
    except Exception as e:
        print(f"Error reading {file_path}: {e}")
        return 0

def count_negative_words(file_path, neg_dictionary_words, stop_words):
    try:
        with open(file_path, "r", encoding="utf-8", errors="ignore") as f:
            text = f.read().lower()
            words = set(re.findall(r'\b\w+\b', text))
            words -= stop_words
            negative_count = sum(1 for word in words if word in neg_dictionary_words)
            word_count = len(words)
            return negative_count, word_count
    except Exception as e:
        print(f"Error reading {file_path}: {e}")
        return 0, 0

# ensure all required columns exist 
required_columns = ['POSITIVE SCORE', 'NEGATIVE SCORE', 'WORD COUNT', 'AVG SENTENCE LENGTH', 
                    'PERCENTAGE OF COMPLEX WORDS', 'FOG INDEX', 'AVG NUMBER OF WORDS PER SENTENCE', 
                    'COMPLEX WORD COUNT', 'SYLLABLE PER WORD', 'PERSONAL PRONOUNS', 'AVG WORD LENGTH']
for col in required_columns:
    if col not in df.columns:
        df[col] = None

# Process positive scores
for filename in df.index:
    file_path = os.path.join(files_directory, filename + ".txt")
    df.at[filename, 'POSITIVE SCORE'] = count_positive_words(file_path, dictionary_words, stop_words) if os.path.exists(file_path) else None

# Process negative scores and word count
for filename in df.index:
    file_path = os.path.join(files_directory, filename + ".txt")
    if os.path.exists(file_path):
        neg_score, word_count = count_negative_words(file_path, neg_dictionary_words, stop_words)
        df.at[filename, 'NEGATIVE SCORE'] = neg_score
        df.at[filename, 'WORD COUNT'] = word_count
    else:
        df.at[filename, 'NEGATIVE SCORE'] = None
        df.at[filename, 'WORD COUNT'] = None

# Calculate polarity and subjectivity scores 
df['POLARITY SCORE'] = (df['POSITIVE SCORE'] - df['NEGATIVE SCORE']) / (df['POSITIVE SCORE'] + df['NEGATIVE SCORE'] + 0.000001)
df['SUBJECTIVITY SCORE'] = (df['POSITIVE SCORE'] + df['NEGATIVE SCORE']) / (df['WORD COUNT'] + 0.000001)

# Updated syllable counting function with "es" and "ed" exceptions
def count_syllables(word):
    word = word.lower().strip()
    if len(word) <= 3:
        return 1
    # Handle exceptions for words ending with "es" or "ed"
    if word.endswith('es') or word.endswith('ed'):
        word = word[:-2]  # Remove the ending before counting
    vowels = "aeiouy"
    count = 0
    prev_vowel = False
    for char in word:
        if char in vowels:
            if not prev_vowel:  # Count only if not preceded by another vowel
                count += 1
            prev_vowel = True
        else:
            prev_vowel = False
    return max(1, count)  # Every word has at least 1 syllable

# Function to count personal pronouns
def count_personal_pronouns(text):
    pronouns = ["I", "we", "my", "ours", "us"]
    # Regex pattern to match pronouns as whole words
    pattern = r'\b(I|we|my|ours|us)\b'
    matches = re.findall(pattern, text, re.IGNORECASE)
    # Exclude "US" (country name) by checking if "us" is uppercase
    count = sum(1 for match in matches if not (match == "us" and match.isupper()))
    return count

# Updated readability function to include new metrics
def calculate_readability(file_path):
    try:
        with open(file_path, "r", encoding="utf-8", errors="ignore") as f:
            text = f.read()
            sentences = [s.strip() for s in re.split(r'[.!?]+', text) if s.strip()]
            num_sentences = len(sentences)
            words = re.findall(r'\b\w+\b', text.lower())
            total_words = len(words)
            complex_words = 0
            total_syllables = 0
            total_chars = 0

            # Process each word for syllables, complexity, and length
            for word in words:
                syllables = count_syllables(word)
                total_syllables += syllables
                if syllables > 2:
                    complex_words += 1
                total_chars += len(word)

            # Existing readability metrics
            avg_sentence_length = total_words / num_sentences if num_sentences > 0 else 0
            percent_complex = (complex_words / total_words) * 100 if total_words > 0 else 0
            fog_index = 0.4 * (avg_sentence_length + percent_complex)

            # New metrics
            syllable_per_word = total_syllables / total_words if total_words > 0 else 0
            personal_pronouns = count_personal_pronouns(text)
            avg_word_length = total_chars / total_words if total_words > 0 else 0

            return (avg_sentence_length, percent_complex, fog_index, complex_words, 
                    syllable_per_word, personal_pronouns, avg_word_length)
    except Exception as e:
        print(f"Error processing {file_path}: {e}")
        return None, None, None, None, None, None, None

# Process readability metrics and populate DataFrame
for filename in df.index:
    file_path = os.path.join(files_directory, filename + ".txt")
    if os.path.exists(file_path):
        (avg_sentence_length, percent_complex, fog_index, complex_words, 
         syllable_per_word, personal_pronouns, avg_word_length) = calculate_readability(file_path)
        df.at[filename, 'AVG SENTENCE LENGTH'] = avg_sentence_length
        df.at[filename, 'PERCENTAGE OF COMPLEX WORDS'] = percent_complex
        df.at[filename, 'FOG INDEX'] = fog_index
        df.at[filename, 'AVG NUMBER OF WORDS PER SENTENCE'] = avg_sentence_length
        df.at[filename, 'COMPLEX WORD COUNT'] = complex_words
        df.at[filename, 'SYLLABLE PER WORD'] = syllable_per_word
        df.at[filename, 'PERSONAL PRONOUNS'] = personal_pronouns
        df.at[filename, 'AVG WORD LENGTH'] = avg_word_length
    else:
        for col in ['AVG SENTENCE LENGTH', 'PERCENTAGE OF COMPLEX WORDS', 'FOG INDEX', 
                    'AVG NUMBER OF WORDS PER SENTENCE', 'COMPLEX WORD COUNT', 
                    'SYLLABLE PER WORD', 'PERSONAL PRONOUNS', 'AVG WORD LENGTH']:
            df.at[filename, col] = None

# Save results back to Excel
df.to_excel(excel_path)
print("Updated DataFrame with new metrics:")
df.head()
