# Data Reporting Script
# Generates example test data files then crates the following:
# Excel file with csv data in each tab, a PDF with graphed data, and a text file with pass/fail results

import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.backends.backend_pdf import PdfPages
import csv
import random
import xlsxwriter
import os
import glob
import time

path = os.path.dirname(os.path.realpath(__file__))
all_files = glob.glob(os.path.join(path, "*.csv"))

# Genrate some test files
def generate_csvs():
    print('Generating test files')
    for u in range (1,11):
        header = ['Temp (C)','Humidity (%)','Voltage (V)']
        filename = 'UnitSN_'+ str(u) +'.csv'
        with open(filename, 'a', newline='') as csvfile:
            writer = csv.writer(csvfile)
            writer.writerow(header)
            for i in range(1, 1001):
                if filename == 'UnitSN_4.csv':  # include some failures
                    Temp = 35 + i/200    
                elif filename == 'UnitSN_6.csv':
                    Humidity = random.choice(range(15,30))
                elif filename == 'UnitSN_8.csv':
                    Volts = 3 - i/500
                else:
                    Temp = 25 + i/200
                    Humidity = random.choice(range(10,15))
                    Volts = 5 - i/500
                df = [Temp, Humidity, Volts]
                writer.writerow(df)
                                    

# Write each csv to a labelled tab in a new Excel file
def data_to_excel(all_files):
    print("Creating Excel Workbook")
    all_files = glob.glob(os.path.join(path, "*.csv"))
    writer = pd.ExcelWriter('All_Units.xlsx', engine='xlsxwriter')

    for f in all_files:
        df = pd.read_csv(f)
        df.to_excel(writer, sheet_name=os.path.basename(f))
    writer.save()

# Generate a PDF will labelled graphs
def plot_csv_data(all_files):
    print("Generating PDF")
    pdf_file = os.path.join(path, 'Unit_Test_Graphs.pdf')
    
    with PdfPages(pdf_file) as pdf:
        for f in all_files:
            # Read the csv file and plot data
            unitSN = (f.split('/')[-1]).split('.csv')[0]
            df = pd.read_csv(f)
            
            # Plot the data from each column in the csv file.
            for i in range(df.shape[1]):
                plt.plot(df.iloc[:, i], label=df.columns[i])
                plt.legend()
                plt.title(unitSN)
                print("      Running...")

            # Save and closethe pdf file.
            pdf.savefig()
            plt.close()


# Append a summary of results to a test log file
def text_summary(all_files):
    txt_file = os.path.join(path, 'Test_Summary.txt')
    for file in all_files:
        # Read the csv file into a Pandas dataframe. Split into columns.
        unitSN = (file.split('/')[-1]).split('.csv')[0]
        df = pd.read_csv(file)
        df_temp = df['Temp (C)']
        df_hum = df['Humidity (%)']
        df_volt = df['Voltage (V)']
        
        max_temp = df_temp.max()
        max_humidity = df_hum.max()
        min_voltage = df_volt.min()
        
# Check results and print failures to text file    
        with open(txt_file, 'a') as f:
            if int(max_temp) > 30:
                f.write(unitSN + ' FAIL: Temperature = ' + str(max_temp) + 'C\n')
                print('\nFAIL ' + unitSN)
            elif int(max_humidity) > 15:
                f.write(unitSN + ' FAIL: Humifidy = ' + str(max_humidity)+ '%\n')
                print('\nFAIL ' + unitSN)
            elif int(min_voltage) < 3:
                f.write(unitSN + ' FAIL: Voltage = ' + str(min_voltage)+ 'V\n')
                print('\nFAIL ' + unitSN)
            else:
                pass
        f.close()


# Main
# Known Bug: Must generate the test files first, then run the reporting functions
# If ran consecutively, there are issues with PDF and TXT trying to generate before data is present
# Workaround is a user menu to make sure the files are present before generating reports

prompt = input("Press 1 to generate test files.\nPress 2 to generate report from esitsting files: ")

if prompt == '1':
    generate_csvs()
    prompt = input("\nTest Files Compelte. Create Reports? y/n ")
    if prompt == 'y':    
        data_to_excel(all_files)
        plot_csv_data(all_files)
        text_summary(all_files)
        print("\nDone")
    elif prompt == "n":
        print("Exiting...")
    else:
        pass
elif prompt == '2':
        data_to_excel(all_files)
        plot_csv_data(all_files)
        text_summary(all_files)
        print("\nDone")
    
else:
    pass
