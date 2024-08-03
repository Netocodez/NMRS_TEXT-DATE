from tkinter import ttk
import time
import pandas as pd
import tkinter as tk
from tkinter import filedialog
from tkinter import *
import string
import tkinter.font as font

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
    df.to_excel(writer, sheet_name="NMRS", startrow=1, header=False, index=False)
    
    workbook = writer.book
    worksheet = writer.sheets["NMRS"]
    
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
    

def AddRadet():
    start_time = time.time()  # Start time measurement

    # Process
    status_label.config(text=f"Attempting to read Basleing Radet...")
    input_Baseline_path = filedialog.askopenfilename(initialdir = '/Desktop', 
                                                        title = 'Select a excel file', 
                                                        filetypes = (('excel file','*.xls'), 
                                                                     ('excel file','*.xlsx'),
                                                                     ('excel file','*.xlsb')))
    
    if not input_Baseline_path:  # Check if the file selection was cancelled
        status_label.config(text="No Radet file selected. Please select a Radet file to convert.")
        return  # Exit the function
    
    status_label.config(text=f"Attempting to read NMRS file...")
    input_file_path = filedialog.askopenfilename(initialdir = '/Desktop', 
                                                        title = 'Select a excel file', 
                                                        filetypes = (('excel file','*.xls'), 
                                                                     ('excel file','*.xlsx'),
                                                                     ('excel file','*.xlsb')))
   
    if not input_file_path:  # Check if the file selection was cancelled
        status_label.config(text="No NMRS file selected. Please select a NMRS file to convert.")
        return  # Exit the function 
        

    dfRadet = pd.read_excel(input_Baseline_path,sheet_name=0, dtype=object)
    df = pd.read_excel(input_file_path,sheet_name=0, dtype=object)

    #Adding Unique Identifiers
    dfRadet['unique identifiers'] = dfRadet["LGA"].astype(str) + dfRadet["Facility"].astype(str) + dfRadet["Hospital Number"].astype(str) + dfRadet["Unique ID"].astype(str)
    df['unique identifiers'] = df["LGA"].astype(str) + df["FacilityName"].astype(str) + df["PatientHospitalNo"].astype(str) + df["PatientUniqueID"].astype(str)

    #remove duplicates
    dfRadet = dfRadet.drop_duplicates(subset=['unique identifiers'], keep=False)
    d = dict(enumerate(string.ascii_uppercase))
    m = df.duplicated(['unique identifiers'], keep=False)
    df.loc[m, 'unique identifiers'] += '_' + df[m].groupby(['unique identifiers']).cumcount().map(d)

    #Adding New Columns
    df['Patient id']=df['unique identifiers'].map(dfRadet.set_index('unique identifiers')['Patient ID'])
    df['Patient id'] = df['Patient id'].fillna(df['unique identifiers'])
    df['Case Manager']=df['unique identifiers'].map(dfRadet.set_index('unique identifiers')['Case Manager'])

    
    #Add Date of Current Status for inactive and active clients
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
        time.sleep(0.000001)
    

    #format and export
    df = df.drop('IIT DATE', axis=1)
    df = df.drop('unique identifiers', axis=1)
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
    

#Creating A tooltip Class
class ToolTip(object):

    def __init__(self, widget):
        self.widget = widget
        self.tipwindow = None
        self.id = None
        self.x = self.y = 0

    def showtip(self, text):
        "Display text in tooltip window"
        self.text = text
        if self.tipwindow or not self.text:
            return
        x, y, cx, cy = self.widget.bbox("insert")
        x = x + self.widget.winfo_rootx() + 320
        y = y + cy + self.widget.winfo_rooty() +15
        self.tipwindow = tw = Toplevel(self.widget)
        tw.wm_overrideredirect(1)
        tw.wm_geometry("+%d+%d" % (x, y))
        label = Label(tw, text=self.text, justify=LEFT,
                      background="#ffffe0", relief=SOLID, borderwidth=1,
                      font=("tahoma", "8", "normal"))
        label.pack(ipadx=1)

    def hidetip(self):
        tw = self.tipwindow
        self.tipwindow = None
        if tw:
            tw.destroy()

def CreateToolTip(widget, text):
    toolTip = ToolTip(widget)
    def enter(event):
        toolTip.showtip(text)
    def leave(event):
        toolTip.hidetip()
    widget.bind('<Enter>', enter)
    widget.bind('<Leave>', leave)
    
#Creating Rounded Button Class
class RoundedButton(tk.Canvas):
  def __init__(self, parent, border_radius, padding, color, text='', command=None):
    tk.Canvas.__init__(self, parent, borderwidth=0,
                       relief="raised", highlightthickness=0, bg=parent["bg"])
    self.command = command
    font_size = 15
    self.font = font.Font(size=font_size, family='Helvetica')
    self.id = None
    height = font_size + (5 * padding)
    width = self.font.measure(text)+(5*padding)

    width = width if width >= 80 else 80

    if border_radius > 0.5*width:
      print("Error: border_radius is greater than width.")
      return None

    if border_radius > 0.5*height:
      print("Error: border_radius is greater than height.")
      return None

    rad = 2*border_radius

    def shape():
      self.create_arc((0, rad, rad, 0),
                      start=90, extent=90, fill=color, outline=color)
      self.create_arc((width-rad, 0, width,
                        rad), start=0, extent=90, fill=color, outline=color)
      self.create_arc((width, height-rad, width-rad,
                        height), start=270, extent=90, fill=color, outline=color)
      self.create_arc((0, height-rad, rad, height), start=180, extent=90, fill=color, outline=color)
      return self.create_polygon((0, height-border_radius, 0, border_radius, border_radius, 0, width-border_radius, 0, width,
                           border_radius, width, height-border_radius, width-border_radius, height, border_radius, height),
                                 fill=color, outline=color)

    id = shape()
    (x0, y0, x1, y1) = self.bbox("all")
    width = (x1-x0)
    height = (y1-y0)
    self.configure(width=width, height=height)
    self.create_text(width/2, height/2,text=text, fill='white', font= self.font)
    self.bind("<ButtonPress-1>", self._on_press)
    self.bind("<ButtonRelease-1>", self._on_release)

  def _on_press(self, event):
      self.configure(relief="sunken")

  def _on_release(self, event):
      self.configure(relief="raised")
      if self.command is not None:
          self.command()


# Creating Main Window
root = tk.Tk()
root.title("NETO's NMRS DATE TEXT TO DATE CONVERTER V001")
root.geometry("600x350")
root.config(bg="#f0f0f0")

# Adding a Button to the Window
convert_button = tk.Button(root, text="JUST NMRS DATE CONVERTER..", command=Convert, font=("Helvetica", 14), bg="blue", fg="#ffffff")
convert_button.pack(pady=0.7)
tooltip = CreateToolTip(convert_button, "This will prompt you to select\n NMRS line List, it only convert'S Date text to date\n It will also prompt you to save converted file\n No New Column is added")
convert_button2 = tk.Button(root, text="NMRS DATE CONVERTER (CMG)", command=AddRadet, font=("Helvetica", 14), bg="blue", fg="#ffffff")
convert_button2.pack(pady=0.5)
tooltip = CreateToolTip(convert_button2, "This will prompt you to select:\n baseline Radet, NMRS line List\n and location to save the file after conversion\n it appends case manager column to the converted file")
#text1 = tk.Label(root, text="Output's ")
#text1.pack(pady=1)
#convert_button1 = tk.Button(root, text="EXIT CONVERTER", command=root.destroy, font=("Helvetica", 14), bg="red", fg="#ffffff")
#convert_button1.pack(pady=20)

convert_button1 = RoundedButton(root, text="EXIT CONVERTER", border_radius=10, padding=4, command=root.destroy, color="red")
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