import sys
from Emailer import Emailer

'''
This program analyzes a spreadsheet of incomplete jobs in the Print and Mail department.
It sends an email to whoever is responsible for the job, notifying them of all the incomplete jobs they are responsible for.
For any questions/problems, please contact Print and Mail IT at pmcompsupport@byu.edu
'''


#This method creates an emailer object and calls the assemble_excel_files method, which creates the excel files and sends the emails
def run(filename, send_emails, save_files=True):
    try:
        emailer = Emailer(filename)
        emailer.assemble_excel_files(send_emails, save_files)
    except Exception as e:
        print(f"Error: {e}")
        sys.exit(1)


def main_function(filename, send_emails, save_files=True):
    run(filename, send_emails, save_files)


