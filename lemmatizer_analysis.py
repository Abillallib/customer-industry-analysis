import pandas as pd
import nltk
from nltk.stem import WordNetLemmatizer

# Download the necessary NLTK resources
nltk.download('wordnet')

# Initialize the lemmatizer
lemmatizer = WordNetLemmatizer()

# Load the Excel file and read Sheet1
file_path = r'C:\path\termfrequency.xlsx'
df = pd.read_excel(file_path, sheet_name='Sheet1')

# Function to get the root word
def get_root_word_lemma(term):
    return lemmatizer.lemmatize(term)

# Apply the function to the "terms" column
df['Root Word'] = df['Term'].apply(get_root_word_lemma)

# Save the updated dataframe to a new Excel file
output_file_path = r'C:\path\termfrequency_amended2.xlsx'
df.to_excel(output_file_path, index=False)

print(f"File has been updated with root words and saved as {output_file_path}")
