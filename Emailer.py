import smtplib
import datetime
from address_reader import AddressReader
import pandas as pd
import openpyxl
import os
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from time import sleep



'''
class Emailer:
Has an AddressReader object, a list of JobGroup objects, a from address, a support address, and the path of the main input file
'''
class Emailer:
    def __init__(self, input_file):
        self.address_reader = AddressReader()
        self.groups = self.address_reader.make_groups()
        self.from_address = "PM_JOB_ALERTS" 
        self.support_address = "pmcompsupport@byu.edu"
        self.input_file = input_file

#TODO: Change send_emails and save_files to boolean values since we aren't using command line arguments anymore 
    def assemble_excel_files(self, send_emails, save_files):
        excel_data = {}
        found = False
        df = pd.read_excel(self.input_file)
        df = df.fillna("nan")
        for index, row in df.iterrows():
            for group in self.groups:
                if row['Next Task'] in group.jobs:
                    if group.name in excel_data:
                        excel_data[group.name].append(row)
                        found = True
                    else:
                        excel_data[group.name] = [row]
                        found = True
                    break
            if not found and row['Next Task'] != "nan":
                print(f"No group found for {row['Next Task']}")
            found = False

        
        for group in self.groups:
            if group.name not in excel_data:
                print(f"No jobs found for {group.name}. No Excel file created.")
            else:
                num_jobs = len(excel_data[group.name])
                group.file = self.create_excel(excel_data[group.name], group.name, save_files)
                if send_emails:
                    self.send_email(group, group.file, num_jobs)
                    if save_files:
                        print(f"Excel file created for {group.name} with {num_jobs} jobs. Saved in Downloads folder. Email sent to {group.email}.")
                    else:
                        print(f"Excel file created for {group.name} with {num_jobs} jobs. Email sent to {group.email}.")
                else:
                    if save_files:
                        print(f"Excel file created for {group.name} with {num_jobs} jobs. Test File Saved in Downloads folder.")
                    else:
                        print(f"Test run for {group.name} with {num_jobs} jobs. Nothing saved.")
        
        if send_emails:
            num_jobs = len(df)
            self.send_log(num_jobs)


        if not save_files:
            self.clear_excel_folder()


#This method sends a log with the number of unfinished jobs to the support address
    def send_log(self, num_jobs):
        print("Sending log")
        server = smtplib.SMTP('128.187.16.31', 25)
        
        msg = MIMEMultipart()
        msg['From'] = self.from_address
        msg['To'] = self.support_address
        msg['Subject'] = f"Job Alert Log for {datetime.datetime.now().strftime('%Y-%m-%d %H-%M-%S')}"
        body = f"Script run at {datetime.datetime.now().strftime('%Y-%m-%d %H-%M-%S')}. {num_jobs} jobs found."
        msg.attach(MIMEText(body, 'plain'))

        server.starttls()
        server.sendmail(self.from_address, self.support_address, msg.as_string())



#Takes in the data, group name, and whether or not to save the file.  Creates the file and returns the path
    def create_excel(self, data, group_name, save_files):
        df = pd.DataFrame(data)
        
        if save_files:
            #Get path to downloads folder and create name for new folder
            downloads_folder = os.path.join(os.path.expanduser('~'), 'Downloads')
            folder_name = f'job_alert_test_files {datetime.datetime.now().strftime("%Y-%m-%d %H-%M")}'

            #Create folder if it doesn't exist
            if not os.path.exists(f'{downloads_folder}/{folder_name}'):
                os.makedirs(f'{downloads_folder}/{folder_name}')
            
            #Create file name and use it to create path
            file_name = f'{group_name} {datetime.datetime.now().strftime("%Y-%m-%d %H-%M-%S")}.xlsx'
            path_name = f'{downloads_folder}/{folder_name}/{file_name}'

        else:
            #If we aren't saving the file, put it in the excel folder
            #The excel folder is a temporary folder that is cleared after the program runs
            excel_folder_path = os.path.join(os.path.dirname(__file__), "excel")
            path_name = f'{excel_folder_path}/{group_name} {datetime.datetime.now().strftime("%Y-%m-%d %H-%M-%S")}.xlsx'

        #Create the excel file and save it to the path
        df.to_excel(path_name, index=False)

        #Format the excel file so that the columns are wide enough
        wb = openpyxl.load_workbook(path_name)
        sheet = wb.active
        sheet.column_dimensions['B'].width = 40
        sheet.column_dimensions['C'].width = 25
        sheet.column_dimensions['D'].width = 25
        sheet.column_dimensions['E'].width = 30
        sheet.column_dimensions['F'].width = 30
        sheet.column_dimensions['G'].width = 30
        sheet.column_dimensions['H'].width = 15
        

        wb.save(path_name)
        wb.close()
        return path_name
    

    #We can use this function if we don't want to save the files.  It clears the excel folder
    #May not even be necessary, but it doesn't hurt to have it
    def clear_excel_folder(self):
        #clear excel folder
        excel_folder = os.path.join(os.path.dirname(__file__), "excel")
        for file in os.listdir(excel_folder):
            file_path = os.path.join(excel_folder, file)
            try:
                if os.path.isfile(file_path):
                    os.unlink(file_path)

            #Windows likes to complain that files are being used by another process, not sure why so we just ignore it
            except Exception as e:
                print("")
    

    '''
    Sends an email to the group leader with the excel file attached
    '''
    def send_email(self, group, filename, num_jobs):
        #Start server
        #This runs on same server as IPPPing, shouldn't cause any issues but good to know just in case
        server = smtplib.SMTP('128.187.16.31', 25)


        #If no email is found for the group leader, we can't send an email.  Print a warning and return
        if group.email == "nan" or group.email == '' or group.email == None:
            print(f"WARNING: No email found for {group.name}. Email not sent.")
            return
        

        #Create the email by building the message and attaching the file
        msg = MIMEMultipart()
        msg['From'] = self.from_address
        msg['To'] = group.email
        msg['Subject'] = f"{num_jobs} Unfinished Jobs in Your Area"

        #Get the first name of the group leader because if we use the full name it looks really formal and weird
        first_name = group.name.split()[0]
        body = f"""Hello {first_name},\n\nYou currently have {num_jobs} unfinished jobs in your area. 
Please see the attached Excel file for more information.\n\n\n This message was generated automatically.  For any questions/problems, please contact Print and Mail IT at {self.support_address}"""
        msg.attach(MIMEText(body, 'plain'))


        #Attach the file using the filename
        attachment = open(filename, "rb")
        part = MIMEBase('application', 'octet-stream')
        part.set_payload(attachment.read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', f'attachment; filename= {group.name} {datetime.datetime.now().strftime("%Y-%m-%d %H-%M-%S")}.xlsx')
        msg.attach(part)
 
        server.starttls()
        server.sendmail(self.from_address, group.email, msg.as_string())

