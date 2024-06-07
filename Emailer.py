import smtplib
import datetime
from address_reader import AddressReader
import pandas as pd


'''
This program analyzes a spreadsheet of incomplete jobs in the Print and Mail department.
It sends an email to whoever is responsible for the job, notifying them of all the incomplete jobs they are responsible for.
For any questions/problems, please contact Print and Mail IT at pmcompsupport@byu.edu
'''

class Emailer:
    def __init__(self, input_file):
        self.address_reader = AddressReader()
        self.groups = self.address_reader.make_groups()
        self.from_address = "pm_jobalerts@byu.edu"
        self.support_address = "pmcompsupport@byu.edu"
        self.input_file = input_file

    def assemble_excel_files(self):
        excel_data = {}
        df = pd.read_excel(self.input_file)
        for index, row in df.iterrows():
            for group in self.groups:
                if row['Next Task'] in group.jobs:
                    if group.name in excel_data:
                        excel_data[group.name].append(row)
                    else:
                        excel_data[group.name] = [row]

        for group in excel_data:
            self.create_excel(excel_data[group], group)



    def create_excel(self, data, group_name):
        df = pd.DataFrame(data)
        writer = pd.ExcelWriter(f'{group_name} {datetime.datetime.now().strftime("%Y-%m-%d %H-%M-%S")}.xlsx')
        df.to_excel(writer, index=False)
        writer._save()






#TEST#
emailer = Emailer("Excel Data.xlsx")

emailer.assemble_excel_files()