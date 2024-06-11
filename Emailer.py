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
This program analyzes a spreadsheet of incomplete jobs in the Print and Mail department.
It sends an email to whoever is responsible for the job, notifying them of all the incomplete jobs they are responsible for.
For any questions/problems, please contact Print and Mail IT at pmcompsupport@byu.edu
'''

class Emailer:
    def __init__(self, input_file):
        self.address_reader = AddressReader()
        self.groups = self.address_reader.make_groups()
        self.from_address = "PM_JOB_ALERTS"
        self.support_address = "pmcompsupport@byu.edu"
        self.input_file = input_file

    def assemble_excel_files(self, send_emails, save_files="yes"):
        send = False
        save = True
        if save_files == "no":
            save = False
        print(send_emails)
        if send_emails == "send_emails":
            send = True
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
                group.file = self.create_excel(excel_data[group.name], group.name)
                if send:
                    self.send_email(group, group.file, num_jobs)
                    print(f"Excel file created for {group.name} with {num_jobs} jobs. Saved in Downloads folder. Email sent to {group.email}.")
                else:
                    print(f"Excel file created for {group.name} with {num_jobs} jobs. Test File Saved in Downloads folder.")

        sleep(1)

        if not save:
            folder_name = group.file.split('/')[-2]
            #remove folder contents
            for group in self.groups:
                if group.file:
                    if group.name == "Brett Hodge":
                        continue
                    print(group.file)
                    os.remove(group.file)

            os.rmdir(f'{os.path.expanduser("~")}/Downloads/{folder_name}')
            print(f"Test files deleted.")




    def create_excel(self, data, group_name):
        df = pd.DataFrame(data)
        #excel_folder_path = os.path.join(os.path.dirname(__file__), "excel")
        downloads_folder = os.path.join(os.path.expanduser('~'), 'Downloads')
        folder_name = f'job_alert_test_files {datetime.datetime.now().strftime("%Y-%m-%d %H-%M")}'

        if not os.path.exists(f'{downloads_folder}/{folder_name}'):
            os.makedirs(f'{downloads_folder}/{folder_name}')
        

        file_name = f'{group_name} {datetime.datetime.now().strftime("%Y-%m-%d %H-%M-%S")}.xlsx'
        path_name = f'{downloads_folder}/{folder_name}/{file_name}'
        writer = pd.ExcelWriter(path_name)
        df.to_excel(writer, index=False)
        writer._save()

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
        writer.close()
        print(f"File Closed: {path_name}")
        return path_name

    def send_email(self, group, filename, num_jobs):
        server = smtplib.SMTP('128.187.16.31', 25)
        
        msg = MIMEMultipart()
        msg['From'] = self.from_address
        msg['To'] = group.email
        msg['Subject'] = f"Print and Mail Job Alert: {group.name} ({num_jobs} Unfinished Jobs)"
        body = f"""Hello {group.name},\n\nYou have {num_jobs} unfinished jobs in your area. 
Please see the attached Excel file for more information.\n\n\n This message was generated automatically.  For any questions/problems, please contact Print and Mail IT at {self.support_address}"""
        msg.attach(MIMEText(body, 'plain'))

        attachment = open(filename, "rb")
        part = MIMEBase('application', 'octet-stream')
        part.set_payload(attachment.read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', f'attachment; filename= {group.name} {datetime.datetime.now().strftime("%Y-%m-%d %H-%M-%S")}.xlsx')
        msg.attach(part)

        server.starttls()
        server.sendmail(self.from_address, group.email, msg.as_string())




        







#TEST#
#emailer = Emailer("Excel Data.xlsx")

#emailer.assemble_excel_files()