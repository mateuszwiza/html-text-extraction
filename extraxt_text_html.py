import pandas as pd
import codecs
from bs4 import BeautifulSoup

# Read the spreadsheet with filenames
df = pd.read_excel("input/Filename.xlsx")

# Empty list for storing results
results = []

# Define cell limit
limit = 32767

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

    # Connect all pieces of textxsinto a single string
    string = ''
    for s in strips:
        string += ' ' + s

    # Divide the string into slices of at most 'limit' characters
    list_of_strings = [filename]
    index = 0
    while index < len(string):
        end = index + limit
        list_of_strings.append(string[index:end])
        index += limit + 1

    # Attach the list of strings for 1 file to the global list 'results'
    results.append(list_of_strings)

# Put results into a dataframe (columns)
df = pd.DataFrame(results)
df = df.rename(columns={0: 'Filename'})

# Export to CSV and Excel
df.to_csv('output/results.csv')
df.to_excel('output/results.xlsx')
