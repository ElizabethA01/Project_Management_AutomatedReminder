from ExcelSheets import ExcelAdapter

# define class for discipline leads
class MDLTracker:
    # start_row = None
    # end_row = None
    def __init__(self, start_row: int, end_row: int) -> None:
        self.start_row = start_row
        self.end_row = end_row

class DisciplineLead:
    # first_name = None
    # last_name = None
    # email= None
    # discipline= None
    spreadsheet_work= MDLTracker

    def __init__(self, first_name: str, last_name: str, email: str, discipline: str, spreadsheet_work: MDLTracker) -> None:
        self.first_name = first_name
        self.last_name = last_name
        self.email = email
        self.discipline = discipline
        self.spreadsheet_work = spreadsheet_work
        self.discipline_leads = []
    
    # # assign discipline lead class
    # def assign_class(contact):
    #     return DisciplineLead(contact)
        
    
    # def get_discipline_contact(self, filename):
    #     for contact in ExcelSheets.ExcelAdapter.df_to_list(filename):
    #         self.discipline_leads.append(self.assign_class(contact))
    #     print(self.discipline_leads)

    # # get certain names from a list input. loop and check names 
    # def find_specific_disciplines():
    #     pass


contacts_list_filename = r'C:\Users\ukaea001\Documents\PythonPrograms\PLMB\List_of_contacts.xlsx'

discipline_lead = []
for lead in ExcelAdapter(contacts_list_filename).df_to_list():
    # DisciplineLead(lead[0:4], spreadsheet_work= lead[-2])
    # print(DisciplineLead(lead))
    # discipline_lead.append(DisciplineLead(lead))

print(discipline_lead)

    


