import yaml
import os
import datetime
import win32com.client
import codecs
from shutil import copyfile

path = "c:\\Users\\bergelvn\\Downloads\\"
date_today = datetime.datetime.now().date()

signature_path = os.path.join((os.environ['USERPROFILE']),'AppData\\Roaming\\Microsoft\\Signatures\Work_files\\') # Finds the path to Outlook signature files with signature name "Work"
html_doc = os.path.join((os.environ['USERPROFILE']),'AppData\\Roaming\\Microsoft\\Signatures\\Niklas standard.htm')     #Specifies the name of the HTML version of the stored signature
html_doc = html_doc.replace('\\\\', '\\') #Removes escape backslashes from path string

html_file = codecs.open(html_doc, 'r', 'utf-8', errors='ignore') #Opens HTML file and ignores errors
signature_code = html_file.read()               #Writes contents of HTML signature file to a string
signature_code = signature_code.replace('Work_files/', signature_path)      #Replaces local directory with full directory path
html_file.close()



if __name__ == '__main__':

    try:
        stream = open("email_addresses.yaml", 'r')
        emailDict = yaml.load(stream, Loader=yaml.FullLoader)

        stream = open("bi-weekly_reports.yaml", 'r')
        projects = yaml.load_all(stream, Loader=yaml.FullLoader)

        os.chdir(path)

        for project in projects:
            try:
                project_name = project["project name"]
                local_dir = project["local dir"]
                jira_id = project["jira id"]

                current_report_name = ("{}.pdf".format(jira_id))
                new_report_name = ("{}_{}_{}.pdf".format(date_today, project_name, jira_id))

                try:
                    try:
                        os.rename(current_report_name, new_report_name)
                    except FileExistsError as err:
                        print("File exists: {}, {}" . format(new_report_name, err))
                    new_destination = ("{}{}\{}".format(local_dir, "bi-weekly reports", new_report_name))
                    copyfile(new_report_name, new_destination)

                    send_email_to = project["send emails to"]
                    to_field = "DL_TCD_Global_PMO_Portfolio_Mgmt"

                    for function, name in send_email_to.items():
                        try:
                            email_address = emailDict[name]
                            #print("{} : {} : {}".format(function, name, email_address))
                            if to_field == (""):
                                to_field = email_address
                            else:
                                to_field = "{};{}" . format(email_address, to_field)
                        except KeyError as err:
                            print("Email Address not found: {}" . format(err))

                    outlook = win32com.client.Dispatch('outlook.application')
                    mail = outlook.CreateItem(0)
                    mail.Subject = ("[PR] {} {}".format(jira_id, project_name))
                    mail.To = to_field
                    mail.HTMLBody = ("{}{}{}".format("<p>Hi,</p>",
                        "<p>biweekly report attached if any questions, please let me know.</p>",
                        signature_code))
                    print("new report name: {}" . format(new_report_name))
                    mail.Display()
                    try:
                        if os.path.isfile(new_report_name):
                            print("File is there")
                            file_path = "{}{}" . format(path, new_report_name)
                            mail.Attachments.Add(file_path)
                            #mail.Attachments.Add("c:\\Users\\bergelvn\Downloads\\2021-09-13_GDGB H3G SRoam LDAP V3 Upgrade_SC-352.pdf")
                        else:
                            print("file {} do not exist" . format(new_report_name))
                    except:
                        print("report not found: {}" . format(new_report_name))

                    #mail.Display()

                except FileNotFoundError as err:
                    print("bi-weekly report for project {} not get generated".format(project_name))
                    print("File Not Found Error: {}".format(err))
            except KeyError as err:
                print("parameters required: {}" . format(err))

    except FileNotFoundError as err:
        print("Configuration files; bi-weekly_reports.yaml and email_addresses.yaml are required")
        print("File Not Found Error: {}".format(err))
