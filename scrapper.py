#!/bin/python3
banner = ''' 
 ____  _____ _____ _____ ____  ____  
|  _ \| ____|_   _| ____|  _ \/ ___| 
| |_) |  _|   | | |  _| | |_) \___ \ 
|  __/| |___  | | | |___|  _ < ___) |
|_|   |_____| |_| |_____|_| \_\____/ 
                                     
 '''

print(banner)
import pandas as pd
import argparse

try:
    doc = ''' Please be sure that you provide absolute path to your csv file\nThat's the file wey the extracts wey 
    you wan extract the inside E.G /location/to/the/file\n
    Also be sure to enter the correct inputs\nEnjoy! '''
    parser = argparse.ArgumentParser(description=doc)
    args = parser.parse_args()

    def sheet_Extracts(filepath):
        print("Please wait...")
        df = pd.read_excel(filepath)       
        uExtract =  input(f"Here are all the available columns which one do you want to extarct \n {df.columns}\n")
        if not uExtract.istitle():
            uExtract = uExtract.title() 
        try:
            print("Please wait...")
            mail_range = input("Enter the range of things wey you for like extract:)\n something like 1-100\n")
            mstart, mend = map(int, mail_range.split("-"))
            extracts = df[uExtract].dropna().tolist()[mstart:mend]
            return extracts
        except Exception as e:                        
            print(f"make sure watin you enter the something like a-z or 20-90\nAnd make sure the name match the Column you entered")  

    def sheet_tofile(extracts, output_file):
        with open (output_file,'a') as file:
            for extract in extracts:
                extract = str(extract)
                extract = extract.replace(".0", "")
                file.write(extract + '\n')
        file.close()        

    filepath = "10kssn.xlsx"
    output_file = "Extracts.txt" 
    extracts = sheet_Extracts(filepath)       
    sheet_tofile(extracts, output_file)
    print("Extracted successfully...")
except Exception as e:
    print(f"There was an error {e}")
