# SpeechWordFrequencyAnalysis
A simple script to count word frequency in a txt file and outputs the result to an Excel file.

This script takes one file as input, counts the word frequency and writes 
the result into an Excel file.
It ignores stopwords and punctuations.
NLTK and xlsxwriter are required.
On *nix machine, simply run 'pip install -U nltk' 
and 'pip install -U xlsxwriter'

Also run this snippet in your terminal:
python
>>> import nltk
>>> nltk.download('stopwords')
>>> nltk.download('punkt')
>>> exit()
