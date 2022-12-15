
from Modules.draftEmail import TimesheetEmail
from Modules.ExcelAdapter import ExcelAdapter
from Modules.ExcelTracker import ExcelTracker
from Modules.disciplines_details import DisciplineLead

# required inputs
filenames = {
    "email_tracker": r'C:\Users\ukaea001\Documents\PythonPrograms\PLMB\Reminders\Email_listing_tracker.xlsx',
    "sheetname": 'Timesheet',
    "contacts_list": r'C:\Users\ukaea001\Documents\PythonPrograms\PLMB\Reminders\List_of_contacts.xlsx',
}
contacts= {  
    "mailCC": 'carolina.morales@wsp.com', # add correct email address
    "mailTo": ['Nelson']
}

def send_timesheet_email(tracker: str, sheetname: str, contact_list: str, required_contacts: str, cc_contacts: str):
    for lead in ExcelAdapter(contact_list).df_to_list():
        leads = DisciplineLead(first_name=lead[0], last_name=lead[1], email=lead[2], discipline=lead[3])
        if leads.first_name in required_contacts:
            outcome, alert = TimesheetEmail.send_timesheet_reminder(first_name= leads.first_name, discipline=leads.discipline, email_to= leads.email, cc_contacts=cc_contacts)
            print(outcome, alert)
            if outcome is True:
                ExcelTracker(tracker, sheetname).add_to_tracker(first_name= leads.first_name, last_name=leads.last_name, discipline=leads.discipline, email= leads.email, comment=sheetname + ' ' + alert, cc_contacts=cc_contacts)
                print('Timesheet email alert activated - message sent')

if __name__ == "__main__":
    send_timesheet_email(filenames['email_tracker'], filenames['sheetname'], filenames['contacts_list'], contacts['mailTo'], contacts['mailCC'])


