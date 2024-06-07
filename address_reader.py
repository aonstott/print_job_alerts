from JobGroup import JobGroup

#open the file
class AddressReader:
    def read_email_addresses(self, groups):
        with open('group_leaders.cfg') as file:
            for line in file:
                (name, email) = line.split(":")
                #remove leading and trailing whitespace from key and value
                name = name.strip()
                email = email.strip()
                for group in groups:
                    if group.name == name:
                        group.email = email
        return groups
    
    def read_job_groups(self):
        groups = []
        with open('groups.cfg') as file:
            for line in file:
                if self.is_new_group(line):
                    group_name = line.split(":")[1].strip()
                    current_group = JobGroup(group_name, [], "")
                    groups.append(current_group)
                else:
                    #if line not empty, add job to current group
                    if line.strip():   
                        current_group.add_job(line.strip())
        return groups 


    def is_new_group(self, line):
        return line.startswith("NAME")  
    
    def make_groups(self):
        groups = self.read_job_groups()
        groups = self.read_email_addresses(groups)
        return groups

    

###TEST###
'''address_reader = AddressReader()
groups = address_reader.read_job_groups()
groups = address_reader.read_email_addresses(groups)
for group in groups:
    print(group)'''

