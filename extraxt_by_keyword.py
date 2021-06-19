import pandas as pd
import codecs
from bs4 import BeautifulSoup

# Read the spreadsheet with filenames
df = pd.read_excel("input/Filename.xlsx")

# Empty list for storing results
results = []

# Set the keyword
keyword = ''

# Repeat for each filename
for filename in df['FileName']:
    # Open the file using codecs
    path = filename # If the file is in the same folder, otherwise modify: path = 'folder/' + filename
    html = codecs.open(path)

    # Remove all HTML tags
    soup = BeautifulSoup(html)
    for script in soup(["script", "style"]):
        script.decompose()
    strips = list(soup.stripped_strings)

    # Look for the keyword between quotation marks and save the next paragraph
    list_of_strings = []
    counter = 0
    for i in range(len(strips)):
        if strips[i] == keyword or strips[i] == '/"' + keyword + '/"':
            list_of_strings.append(strips[i + 1])
            counter += 1

    # If nothing found in document, create a row with filename and 0
    if counter == 0:
        results.append([filename, counter, ''])

    # Else, get rows with filename, frequency and paragraph
    else:
        for string in list_of_strings:
            results.append([filename, counter, string])

# Put results into a dataframe (columns)
df = pd.DataFrame(results, columns=['Filename', 'Frequency', 'Paragraph'])

# Export to CSV and Excel
df.to_csv('output/keyword_results.csv')
df.to_excel('output/keyword_results.xlsx')
