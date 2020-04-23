import csv
import subprocess
from xlsxwriter.workbook import Workbook

# Cerca la directory corrente in cui è eseguita l'applicazione
def directory():
    current_directory = subprocess.check_output('echo %cd%', shell = True)
    current_directory_list = list(str(current_directory))

    if "b" in current_directory_list and "'" in current_directory_list:
        current_directory_list.remove("b")
        current_directory_list.remove("'")

    for _ in range(0, 4):
        current_directory_list.pop()
    
    current_directory_list.pop(len(current_directory_list) - 1)

    current_directory_string = ''.join(current_directory_list)
    current_directory_list = current_directory_string.split("\\\\")
    current_directory_string = '\\'.join(current_directory_list)

    return current_directory_string

# Convertitore
def converter(csvfile):
    workbook = Workbook('output.xlsx')
    worksheet = workbook.add_worksheet()

    with open(csvfile, 'rt', encoding = 'utf-8') as file_:
        reader = csv.reader(file_, delimiter = ';')

        for r, row in enumerate(reader):
            for c, col in enumerate(row):
                worksheet.write(r, c, col)

    workbook.close()

def main():
    pathCSV = directory() + "\\" + "input.CSV"

    converter(pathCSV)

main()