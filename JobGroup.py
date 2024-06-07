class JobGroup:
    def __init__(self, name, jobs, email):
        self.name = name
        self.jobs = jobs
        self.email = email
    
    def add_job(self, job):
        self.jobs.append(job)

    def __str__(self):
        return f"Name: {self.name}\nJobs: {self.jobs}\nEmail: {self.email}"