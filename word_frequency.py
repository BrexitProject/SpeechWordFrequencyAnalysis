# This script takes one file as input, counts the word frequency and writes 
# the result into an Excel file.
# It ignores stopwords and punctuations.
# NLTK and xlsxwriter are required.
# On *nix machine, simply run 'pip install -U nltk' 
# and 'pip install -U xlsxwriter'

# Also run this snippet in your terminal:
# python
# >>> import nltk
# >>> nltk.download('stopwords')
# >>> nltk.download('punkt')
# >>> exit()

import sys
import codecs
import nltk
import xlsxwriter
import operator

# default_stopwords = set(nltk.corpus.stopwords.words('english'))

input_file = sys.argv[1]

fp = codecs.open(input_file, 'r', 'utf-8')

words = nltk.word_tokenize(fp.read())

# Remove single-character tokens (mostly punctuation)
words = [word for word in words if len(word) > 1]

# Remove numbers
words = [word for word in words if not word.isnumeric()]

# Lowercase all words (default_stopwords are lowercase too)
words = [word.lower() for word in words]

# Remove stopwords
# words = [word for word in words if word not in default_stopwords]

# Calculate frequency distribution
freq_dist = nltk.FreqDist(words)
sorted_freqDist = sorted(
    freq_dist.items(), key=operator.itemgetter(1), reverse=True)

workbook = xlsxwriter.Workbook('%s_word_frequency.xlsx' % input_file)
worksheet = workbook.add_worksheet()

d = {'a': ['e1', 'e2', 'e3'], 'b': ['e1', 'e2'], 'c': ['e1']}
row = 0
col = 0

for pair in sorted_freqDist:
    row += 1
    worksheet.write(row, col, pair[0])
    worksheet.write(row, col + 1, pair[1])

workbook.close()

print(sorted_freqDist)

fp.close()
