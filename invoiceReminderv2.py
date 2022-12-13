
from Modules.DraftEmail import InvoiceEmail
from Modules.ExcelAdapter import ExcelAdapter
from Modules.ExcelTracker import ExcelTracker
from Modules.disciplines_details import DisciplineLead

# email tracking list excel file location
email_tracker_filename = r'C:\Users\ukaea001\Documents\PythonPrograms\PLMB\Reminders\Email_listing_tracker.xlsx'
sheetname = 'Invoice'
contacts_list_filename = r'C:\Users\ukaea001\Documents\PythonPrograms\PLMB\Reminders\List_of_contacts.xlsx'
# required_contacts = ['David', 'Benny', 'Nelson']
required_contacts = ['Benny']

def send_invoice_email(tracker, sheetname, contacts):
    for lead in ExcelAdapter(contacts).df_to_list():
        leads = DisciplineLead(first_name=lead[0], last_name=lead[1], email=lead[2], discipline=lead[3])
        if leads.first_name in required_contacts:
            InvoiceEmail.send_invoice_reminder(first_name= leads.first_name, discipline=leads.discipline, email_to= leads.email)
            ExcelTracker(tracker, sheetname).add_to_tracker(first_name= leads.first_name, last_name=leads.last_name, discipline=leads.discipline, email= leads.email)


if __name__ == "__main__":
    send_invoice_email(email_tracker_filename, sheetname, required_contacts)



