import os
import pandas as pd
from docx import Document

def read_table_from_docx(file_path):
    document = Document(file_path)
    table_data = []

    for table in document.tables:
        for i, row in enumerate(table.rows):
            if i == 0:  # Ignore the first row (header row) of each table
                continue
            row_data = []
            for cell in row.cells:
                row_data.append(cell.text)
            table_data.append(row_data)
    return table_data

input_directory = ''
output_csv = ''

all_data = []

for file_name in os.listdir(input_directory):
    if file_name.startswith('~'):
        continue

    file_path = os.path.join(input_directory, file_name)
    if file_path.endswith('.docx'):
        print(f'Processing file: {file_path}')
        table_data = read_table_from_docx(file_path)
        print(f'Table data: {table_data}')
        all_data.extend(table_data)

headers = ['Nr.', 'Betrag', 'Verletzung', 'Dauer und Umfang der Behandlung, Arbeitsunfähigkeit', 'Person des Verletzten', 'Dauerschaden', 'Besondere Umstände', 'Gericht, Datum der Entscheidung, Az., Veröffentlichung bzw. Einsender']

df = pd.DataFrame(all_data, columns=headers)

# Remove rows with incorrect number of columns
df = df[df.apply(lambda x: len(x) == len(headers), axis=1)]

# Export the DataFrame to a CSV file with the specified encoding
df.to_csv(output_csv, index=False, encoding='utf-8-sig')
