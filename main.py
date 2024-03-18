import pandas as pd
import re
import camelot
import openpyxl

def extract_numbers(text):
    ind = re.findall(r'(^\d+)\.', text)
    if(len(ind) == 0):
        ind.append(-1)
    return ind[0]

def pdf2df(pdf_file_path: str, init_pag: int = 3, end_pag: int = 9):
    list_pag = []

    for pag in range(init_pag, end_pag):
        # Read tables from all pages using flavor='stream'
        tables = camelot.read_pdf(pdf_file_path, flavor='stream', edge_tol=500, pages=str(pag))

        # Access the first table's DataFrame
        df_Dwelling = tables[0].df
        df_Dwelling.rename(columns={0: 'DESCRIPTION'}, inplace=True)
        df_Dwelling.rename(columns={1: 'QUANTITY'}, inplace=True)
        df_Dwelling.rename(columns={2: 'UNIT PRICE'}, inplace=True)
        df_Dwelling.rename(columns={3: 'TAX'}, inplace=True)
        df_Dwelling.rename(columns={4: 'RCV'}, inplace=True)
        df_Dwelling.rename(columns={5: 'ACV'}, inplace=True)
        df_Dwelling['index_pdf'] = df_Dwelling['DESCRIPTION'].apply(extract_numbers)
        list_pag.append(df_Dwelling)
    return list_pag

def editExcel(path: str, sheet_name: str, start_line: int, end_line: int, list_table: list, save_path: str):
    # Load the workbook
    workbook = openpyxl.load_workbook(path)
    sheet = workbook[sheet_name]

    for index in range(start_line, end_line):
        value = sheet['A' + str(index)].value
        if(isinstance(value, (int,float))):
            for table in list_table:
                text_value = str(int(value))
                if(text_value in table['index_pdf'].values):
                    value_pdf = table.query('index_pdf == @text_value')
                    sheet['B' + str(index)] = value_pdf['DESCRIPTION'].item()
                    sheet['C' + str(index)] = value_pdf['QUANTITY'].item()
                    sheet['D' + str(index)] = value_pdf['RCV'].item()
                    break
    workbook.save(save_path)


#main                
pdf_file_path = r'C:\Users\Ruan Lucas Donino\Documents\Projects_Python\ScrapPDF_Upwork\herrmann_updated_claim.pdf'
excel_file_path = r'herrmann_sheet.xlsx'
sheet = 'Sheet1'

list_data_pdf = pdf2df(pdf_file_path, 3, 9)
editExcel(excel_file_path, sheet, 1, 90, list_data_pdf, 'herrmann_updated_claim.xlsx')
