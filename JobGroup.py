'''
class JobGroup:
contains name of the group, list of tasks that fall into that group, email address of the group leader, and the filename of the excel sheet 
containing the report for that group
'''


class JobGroup:
    def __init__(self, name, jobs, email):
        self.name = name
        self.jobs = jobs
        self.email = email
        self.file = None

    
    def add_job(self, job):
        self.jobs.append(job)

    def __str__(self):
        return f"Name: {self.name}\nJobs: {self.jobs}\nEmail: {self.email}"
