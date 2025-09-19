# Load necessary libraries
library(tm)
library(SnowballC)
library(wordcloud)
library(writexl)

# Step 1: Read the Text File
# Replace '.txt' with the actual file path.
setwd(dirname(rstudioapi::getActiveDocumentContext()$path))
text_data <- readLines("Customer Company Names.txt", warn = FALSE)

# Step 2: Create a Corpus from Text Data
corpus <- Corpus(VectorSource(text_data))
View(corpus)

# Step 3: Preprocess the Text Data
# Convert to lowercase
corpus <- tm_map(corpus, content_transformer(tolower))

# Remove punctuation
corpus <- tm_map(corpus, removePunctuation)

# Remove numbers
corpus <- tm_map(corpus, removeNumbers)

# Remove stopwords (common words that don't add much meaning)
corpus <- tm_map(corpus, removeWords, stopwords("en"))

# Strip any extra whitespace
corpus <- tm_map(corpus, stripWhitespace)

# Step 4: Create a Term-Document Matrix and Calculate Word Frequencies
tdm <- TermDocumentMatrix(corpus)
tdm_matrix <- as.matrix(tdm)

# Calculate the frequency of each term
term_frequency <- rowSums(tdm_matrix)
term_frequency <- sort(term_frequency, decreasing = TRUE)

# Optional: Generate the Word Cloud
# Set minimum frequency for words to appear in the cloud with `min.freq`
# Choose colors with `brewer.pal`
wordcloud(names(term_frequency), term_frequency, min.freq = 1, colors = brewer.pal(8, "Dark2"))


# Optional: Print most common terms to check them
print(head(term_frequency, 10))


# Step 5: Output to Excel
term_frequency_df <- data.frame(
Term = names(term_frequency),
       Frequency = as.numeric(term_frequency),
       stringsAsFactors = FALSE
      )
View(term_frequency_df)
# Write to Excel
write_xlsx(term_frequency_df, "termfrequency.xlsx")
