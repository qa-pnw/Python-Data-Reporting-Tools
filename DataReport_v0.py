# Data Reporting Script
# Generates example test data files then crates the following:
# Excel file with csv data in each tab, a PDF with graphed data

import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.backends.backend_pdf import PdfPages
import csv
import random
import xlsxwriter
import os
import glob

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
                    Temp = 31 + i/200    
                if filename == 'UnitSN_6.csv':
                    Humidity = random.choice(range(15,30))
                if filename == 'UnitSN_8.csv':
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
            
            for i in range(df.shape[1]):
                plt.plot(df.iloc[:, i], label=df.columns[i])
                plt.legend()
                plt.title(unitSN)
                print("      Running...")
            # Save the plot to the pdf file.
            pdf.savefig()
            plt.close()


# Function calls
generate_csvs() 
plot_csv_data(all_files)
data_to_excel(all_files)
print("Done")
