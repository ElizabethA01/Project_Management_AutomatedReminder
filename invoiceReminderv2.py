
from Modules.DraftEmail import InvoiceEmail
from Modules.ExcelAdapter import ExcelAdapter
from Modules.disciplines_details import DisciplineLead, MDLTracker

# email tracking list excel file location
email_tracker_filename = r'C:\Users\ukaea001\Documents\PythonPrograms\PLMB\Reminders\Email_listing_tracker.xlsx'
invoice_sheetname = 'Invoice'
contacts_list_filename = r'C:\Users\ukaea001\Documents\PythonPrograms\PLMB\List_of_contacts.xlsx'
required_contacts = ['David', 'Benny', 'Nelson']

for lead in ExcelAdapter(contacts_list_filename).df_to_list():
    leads = DisciplineLead(first_name=lead[0], last_name=lead[1], email=lead[2], discipline=lead[3], spreadsheet_work=MDLTracker(start_row=lead[4], end_row=lead[5]))
    if leads.first_name in required_contacts:
        InvoiceEmail.send_invoice_reminder(first_name= leads.first_name, discipline=leads.discipline, email_to= leads.email)


# send_invoice_email()



