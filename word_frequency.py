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

import re
import sys
import codecs
import nltk
import xlsxwriter
import operator
import os
import glob
from collections import Counter

default_stopwords = set(nltk.corpus.stopwords.words('english'))

custom_filter = {'while', 'been', 'too', "hadn't", 'once', "weren't", 'hadn', "didn't", 'who',
                 'am', 'upon', 'have', "isn't", 'after', 'where', 'doesn', 'at', 'to', 'this', 'about', 'from', 'under',
                 'its',  'has', 'having', 'y',  'which', 'any', 'other', 'such', 'were', 'but', 'because', 'o',
                 've', 're', 'being', 'these', 'or', 'ma', 'will', 'the', 'how', 'whom', 'was', "aren't",
                 'hasn', 'below',  'those', 'is', 'more',  'until', 'had', 'in', 'again', 'before',
                 'during', 'up', 'an', 'on', 'm', 'for', 'by', "hasn't", 'off', 'if', 'all', 'out', 'wasn', 'it',
                 'does', 'itself', 'not', 'own', 'shan', 'between', 'then', 's', 'of', 'did', 'be',
                 'there', 'that', 'than', 'some', "it's", 'do', 'doing', 'most', 'isn', "wasn't",
                 'here', 'with', 'nor', 'just', 'each', 'are', 'a', 'as', 'into', 'weren', 'and',
                 'didn', 'when', 'through', 'aren', 'what', 'ain', 'so'}

input_file = sys.argv[1]

path = sys.argv[1]

freq_dict = Counter()

first_person = {'we', 'our', 'us', 'ours', 'ourselves'}
second_person = {'you', 'your', 'yours', 'yourself', 'yourselves'}
third_person = {'he', 'she', 'his', 'her', 'hers',
                'they', 'them', 'theirs', 'theirselves'}


def count_one_file(filename):
    global freq_dict
    fp = codecs.open(filename, 'r', 'utf-8')
    raw = fp.read()
    words = nltk.word_tokenize(raw)
    words = [word for word in words if len(word) > 1]
    words = [word for word in words if not word.isnumeric()]
    words = [word.lower() for word in words]
    words = [word for word in words if word not in custom_filter]
    file_freq_dict = Counter(nltk.FreqDist(words))
    file_freq_dict["have to"] = len(re.findall("have to", raw))
    file_freq_dict["first person"] = len(
        [word for word in words if word in first_person])
    file_freq_dict["second person"] = len(
        [word for word in words if word in second_person])
    file_freq_dict["third person"] = len(
        [word for word in words if word in third_person])
    freq_dict = freq_dict + file_freq_dict
    fp.close()


if os.path.isdir(path):
    for filename in glob.glob(os.path.join(path, '*.txt')):
        print(filename)
        count_one_file(filename)

if os.path.isfile(path):
    count_one_file(path)


sorted_freqDist = sorted(
    freq_dict.items(), key=operator.itemgetter(1), reverse=True)

workbook = xlsxwriter.Workbook('%s_word_frequency.xlsx' % input_file)
worksheet = workbook.add_worksheet()

d = {'a': ['e1', 'e2', 'e3'], 'b': ['e1', 'e2'], 'c': ['e1']}
row = 0
col = 0

for pair in sorted_freqDist:
    
    row += 1
    if pair[0] == "first person" or pair[0] == "second person" or pair[0] == "third person":
      format = workbook.add_format()
      format.set_bg_color('red')
      worksheet.write(row, col, pair[0], format)
      worksheet.write(row, col + 1, pair[1], format)
    else:
      worksheet.write(row, col, pair[0])
      worksheet.write(row, col + 1, pair[1])

workbook.close()
