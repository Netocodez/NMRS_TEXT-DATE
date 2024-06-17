from tkinter import ttk
import time
import pandas as pd
import tkinter as tk
from tkinter import filedialog
from tkinter import *

def Convert():
    start_time = time.time()  # Start time measurement

    # Process
    status_label.config(text=f"Attempting to read file...")
    input_file_path = filedialog.askopenfilename(initialdir = '/Desktop', 
                                                        title = 'Select a excel file', 
                                                        filetypes = (('excel file','*.xls'), 
                                                                     ('excel file','*.xlsx'),
                                                                     ('excel file','*.xlsb')))
    
    if not input_file_path:  # Check if the file selection was cancelled
        status_label.config(text="No file selected. Please select a file to convert.")
        return  # Exit the function

    #Add Date of Current Status for inactive and active clients
    df = pd.read_excel(input_file_path,sheet_name=0, dtype=object)
    df['IIT DATE'] = pd.to_datetime(df['LastPickupDate']) + pd.to_timedelta(df['DaysOfARVRefil'], unit='D') + pd.Timedelta(days=29)
    df.loc[(df['CurrentARTStatusWithPillBalance'] == 'InActive') & (df['PatientOutcomeDate'].isna()), 'PatientOutcomeDate'] = df['IIT DATE']
    df.loc[(df['CurrentARTStatusWithPillBalance'] == 'Active') & (df['PatientOutcomeDate'].isna()), 'PatientOutcomeDate'] = df['LastPickupDate']
        
            
    #Convert Date Objects to Date
    dfDates = ['DateOfBirth','ARTStartDate','LastPickupDate','LastINHDispensedDate','ViralLoadSampleCollectionDate','ViralLoadEncounterDate','EnrollmentDate','PatientOutcomeDatePreviousQuarter','PatientOutcomeDate','BiometricCaptureDate','HIVConfirmedDate','LastVisitDate','InitialCD4CountDate','CurrentCD4CountDate','LastEACDate','PregnancyStatusDate','LastDeliveryDate','ViralLoadReportedDate','ResultDate','AssayDate','ApprovalDate','DateOfTermination','PharmacyNextAppointment','ClinicalNextAppointment','MarkAsDeseasedDeathDate','CurrentWeightDate','TBStatusDate','BaselineINHStartDate','BaselineINHStopDate','CurrentINHStartDate','CurrentINHOutcomeDate','BaselineTBTreatmentStartDate','BaselineTBTreatmentStopDate','LastViralLoadSampleCollectionFormDate','LastSampleTakenDate','OTZEnrollmentDate','OTZOutcomeDate','InitialFirstLineRegimenDate','InitialSecondLineRegimenDate','LastPickupDatePreviousQuarter','RecaptureDate']
    for col in dfDates:
        df[col] = pd.to_datetime(df[col],errors='coerce').dt.date

    # Initialize the progress bar maximum value
    progress_bar['maximum'] = len(df)  # Set to 100 for percentage completion

    for index, row in df.iterrows():
                
        # Update the progress bar value
        progress_bar['value'] = index + 1
        
        # Calculate the percentage of completion
        percentage = ((index + 1) / len(df)) * 100
        
        # Update the status label with the current percentage
        status_label.config(text=f"Converting Date Text to Date: {index + 1}/{len(df)} ({percentage:.2f}%)")
        
        # Update the GUI to reflect changes
        root.update_idletasks()
        
        # Simulate time-consuming task
        time.sleep(0.00001)
    

    #format and export
    df = df.drop('IIT DATE', axis=1)
    status_label2.config(text=f"Just a moment! Formating and Saving Converted File...")
    output_file_name = input_file_path.split("/")[-1][:-4]
    output_file_path = filedialog.asksaveasfilename(initialdir = '/Desktop', 
                                                    title = 'Select a excel file', 
                                                    filetypes = (('excel file','*.xls'), 
                                                                 ('excel file','*.xlsx')),defaultextension=".xlsx", initialfile=output_file_name)

    if not output_file_path:  # Check if the file save was cancelled
        status_label.config(text="File conversion was cancelled. No file was saved.")
        progress_bar['value'] = 0
        return  # Exit the function
        
    writer = pd.ExcelWriter(output_file_path, engine="xlsxwriter")
    df.to_excel(writer, sheet_name="NMRS-RADET", startrow=1, header=False, index=False)
    
    workbook = writer.book
    worksheet = writer.sheets["NMRS-RADET"]
    
    # Add a header format.
    header_format = workbook.add_format(
        {
            "bold": True,
            "text_wrap": True,
            "valign": "bottom",
            "fg_color": "#D7E4BC",
            "border": 1,
        }
    )
    
    # Write the column headers with the defined format.
    for col_num, value in enumerate(df.columns.values):
        worksheet.write(0, col_num + 0, value, header_format)
    
    # Close the Pandas Excel writer and output the Excel file.
    writer.close()
    
    # Update the status label upon completion
    end_time = time.time()  # End time measurement
    total_time = end_time - start_time  # Calculate total time taken
    status_label.config(text=f"Conversion Complete! Time taken: {total_time:.2f} seconds")
    status_label2.config(text=f" Converted File Location: {output_file_path}")

# Creating Main Window
root = tk.Tk()
root.title("NETO's NMRS DATE TEXT TO DATE CONVERTER V001")
root.geometry("600x350")
root.config(bg="#f0f0f0")

# Adding a Button to the Window
convert_button = tk.Button(root, text="SELECT FILE & CONVERT", command=Convert, font=("Helvetica", 14), bg="blue", fg="#ffffff")
convert_button.pack(pady=10)
#text1 = tk.Label(root, text="Output's ")
#text1.pack(pady=1)
convert_button1 = tk.Button(root, text="EXIT CONVERTER", command=root.destroy, font=("Helvetica", 14), bg="red", fg="#ffffff")
convert_button1.pack(pady=20)
      
# Progress bar widget
progress_bar = ttk.Progressbar(root, orient='horizontal', length=300, mode='determinate')
progress_bar.pack(pady=10)

# Label for percentage and progress messages
status_label = tk.Label(root, text="0%", font=('Helvetica', 12))
status_label.pack(pady=1)

status_label2 = tk.Label(root, text='Welcome to NMRS DATE CONVERTER ', bg="#D3D3D3", font=('Helvetica', 9))
status_label2.pack(pady=15)

#text = tk.Label(root, text="Welcome! to NMRS TEXT DATE to DATE Converter!")
#text.pack(pady=1)
text3 = tk.Label(root, text="you will be prompted to select a file and the location you want to save the converted file")
text3.pack(pady=1)

text2 = tk.Label(root, text="Contacts: email: chinedum.pius@gmail.com, phone: +2348134453841")
text2.pack(pady=15)


# Adding File Dialog
filedialog = tk.filedialog 

# Running the GUI
root.mainloop()

#pyinstaller NMRS_TEXT-DATE.py --onefile --windowed --upx-dir "C:\upx-4.2.4-win64"