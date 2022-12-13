
from Modules.draftEmail import InvoiceEmail
from Modules.ExcelAdapter import ExcelAdapter
from Modules.ExcelTracker import ExcelTracker
from Modules.disciplines_details import DisciplineLead

# email tracking list excel file location
email_tracker_filename = r'C:\Users\ukaea001\Documents\PythonPrograms\PLMB\Reminders\Email_listing_tracker.xlsx'
sheetname = 'Invoice'
contacts_list_filename = r'C:\Users\ukaea001\Documents\PythonPrograms\PLMB\Reminders\List_of_contacts.xlsx'
mailCC_contacts = 'carolina.morales@wsp.com'
required_contacts = ['David', 'Benny', 'Nelson']

def send_invoice_email(tracker: str, sheetname: str, contact_list: str, required_contacts: str, cc_contacts: str):
    for lead in ExcelAdapter(contact_list).df_to_list():
        leads = DisciplineLead(first_name=lead[0], last_name=lead[1], email=lead[2], discipline=lead[3])
        if leads.first_name in required_contacts:
            outcome = InvoiceEmail.send_invoice_reminder(first_name= leads.first_name, discipline=leads.discipline, email_to= leads.email, cc_contacts=cc_contacts)
            if outcome == True:
                print('Invoice email alert activated - message sent')
                ExcelTracker(tracker, sheetname).add_to_tracker(first_name= leads.first_name, last_name=leads.last_name, discipline=leads.discipline, email= leads.email, cc_contacts=cc_contacts)

if __name__ == "__main__":
    send_invoice_email(email_tracker_filename, sheetname, contacts_list_filename, required_contacts, mailCC_contacts)



