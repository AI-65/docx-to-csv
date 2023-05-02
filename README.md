# Docx to CSV Converter

This Python script extracts tabular data from multiple Microsoft Word documents (.docx) and combines them into a single CSV file. The script is designed to work with a specific table structure, containing information about court decisions on compensation for pain and suffering (Schmerzensgeld) in Germany. You can customize the column names according to your needs.

## Prerequisites

- Python 3.6 or higher
- Install the required packages with the following command: 

  pip install pandas python-docx


## Usage

1. Place all the .docx files you want to process in a folder (e.g., `schmerzensgeldbackup`).
2. Update the `input_directory` variable in the script with the path to the folder containing the .docx files.
3. Update the `output_csv` variable in the script with the path where you want to save the resulting CSV file.
4. If necessary, customize the column names by modifying the `headers` list in the script.
5. Run the script with the following command:

  python docx_to_csv.py


The script will process all the .docx files in the input directory and create a CSV file containing the combined data.

## Table Structure

The script is designed to work with .docx files containing tables with the following columns by default:

- Nr.
- Betrag
- Verletzung
- Dauer und Umfang der Behandlung, Arbeitsunfähigkeit
- Person des Verletzten
- Dauerschaden
- Besondere Umstände
- Gericht, Datum der Entscheidung, Az., Veröffentlichung bzw. Einsender

You can customize the column names by modifying the `headers` list in the script. The header row of each table is ignored during the processing to avoid duplicating the headers in the final CSV file.

## License

This project is licensed under the terms of the MIT License.




