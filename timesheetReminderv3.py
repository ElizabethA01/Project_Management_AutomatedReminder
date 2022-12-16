from Modules.draft_email import TimesheetEmail
from Modules.excel_adapter import ExcelAdapter
from Modules.excel_tracker import ExcelTracker
from Modules.disciplines_details import DisciplineLead

# required inputs
filenames = {
    "email_tracker": '',
    "sheetname": 'Timesheet',
    "contacts_list": ''
}
contacts= {  
    "mailCC": "",
    "mailTo": []
}

def send_timesheet_email(tracker: str, sheetname: str, contact_list: str, required_contacts: str, cc_contacts: str):
    df = ExcelAdapter(contact_list).extract_data_to_df()
    for rowNum in range(len(df)):
        if df['First name'][rowNum] in required_contacts:
            leads = DisciplineLead(first_name=df['First name'][rowNum], last_name=df['Last name'][rowNum], email=df['Email address'][rowNum], discipline=df['Discipline'][rowNum])
            outcome, alert = TimesheetEmail.send_timesheet_reminder(first_name= leads.first_name, discipline=leads.discipline, email_to= leads.email, cc_contacts=cc_contacts)
            if outcome is True:
                print('Timesheet email alert activated - message sent')
                ExcelTracker(tracker, sheetname).add_to_tracker(first_name= leads.first_name, last_name=leads.last_name, discipline=leads.discipline, email= leads.email, comment=sheetname + ' ' + alert, cc_contacts=cc_contacts)
            else:
                print(outcome)


if __name__ == "__main__":
    send_timesheet_email(filenames['email_tracker'], filenames['sheetname'], filenames['contacts_list'], contacts['mailTo'], contacts['mailCC'])



