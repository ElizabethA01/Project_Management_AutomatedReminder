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



    


