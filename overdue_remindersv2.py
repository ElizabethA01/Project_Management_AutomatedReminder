from Modules.excel_adapter import ExcelAdapter
from Modules.disciplines_details import DisciplineLead, MDLTracker
from Modules.ml1_tracker import ML1Tracker
from Modules.draft_email import OverdueReminders

filenames = {
    "ml1 sheetname": 'SAT MDL Tracker',
    "contacts_list": r'C:\Users\ukaea001\Documents\PythonPrograms\PLMB\Reminders\List_of_contacts.xlsx',
    "tracker_link": 'https://wsponline.sharepoint.com/:x:/r/sites/CO-P1808/2_Wip/2_2_TEC/5.7_SYS_EQU/02_GEN/RAS_Tracker/V3/WSP-GEN-FOR-RAS-TRA-001_V3.xlsm?d=w64cbca852ffc461291e16bef2f99478a&csf=1&web=1&e=i8F16T'

}
contacts= {  
    "mailCC": 'carolina.morales@wsp.com', # add correct email address
}

def send_overdue_reminders(sheetname: str, contact_list: str, cc_contacts: str, tracker: str):
    df = ExcelAdapter(contact_list).extract_data_to_df() 
    for rowNum in range(len(df)):
        if df['MDL tracker start row'][rowNum] != "":
            leads = DisciplineLead(first_name=df['First name'][rowNum], last_name=df['Last name'][rowNum], email=df['Email address'][rowNum], discipline=df['Discipline'][rowNum], spreadsheet_work= MDLTracker(start_row=df['MDL tracker start row'][rowNum], end_row= df['MDL tracker end row'][rowNum]))
            overdue = ML1Tracker.check_overdue_items(leads.spreadsheet_work.start_row, leads.spreadsheet_work.end_row)
            if overdue != 0:
                OverdueReminders.send_overdue_reminder(first_name= leads.first_name, discipline=leads.discipline, email_to= leads.email, overdue=overdue, tracker_link=tracker, cc_contacts=cc_contacts)
            else:
                print(f'no overdue items for {leads.discipline} - {leads.first_name} {leads.last_name}')



if __name__ == "__main__":
    send_overdue_reminders(filenames['ml1 sheetname'], filenames['contacts_list'], contacts['mailCC'], filenames['tracker_link'])
