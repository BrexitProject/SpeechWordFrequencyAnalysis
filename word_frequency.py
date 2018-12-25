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
import os
import glob
from collections import Counter

default_stopwords = set(nltk.corpus.stopwords.words('english'))

input_file = sys.argv[1]

path = sys.argv[1]

freq_dict = Counter()

if os.path.isdir(path):
  for filename in glob.glob(os.path.join(path, '*.txt')):
    print(filename)
    fp = codecs.open(filename, 'r', 'utf-8')
    words = nltk.word_tokenize(fp.read()) 
    words = [word for word in words if len(word) > 1]
    words = [word for word in words if not word.isnumeric()]
    words = [word.lower() for word in words]
    file_freq_dict = Counter(nltk.FreqDist(words))
    freq_dict = freq_dict + file_freq_dict

if os.path.isfile(path):
  fp = codecs.open(path, 'r', 'utf-8')
  words = nltk.word_tokenize(fp.read()) 
  words = [word for word in words if len(word) > 1]
  words = [word for word in words if not word.isnumeric()]
  words = [word.lower() for word in words]
  file_freq_dict = Counter(nltk.FreqDist(words))
  freq_dict = freq_dict + file_freq_dict

sorted_freqDist = sorted(
    freq_dict.items(), key=operator.itemgetter(1), reverse=True)

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

fp.close()
