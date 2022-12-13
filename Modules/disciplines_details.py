# define class for discipline leads
class MDLTracker:
    def __init__(self, start_row: int, end_row: int) -> None:
        self.start_row = start_row
        self.end_row = end_row

class DisciplineLead:

    def __init__(self, first_name: str, last_name: str, email: str, discipline: str, spreadsheet_work: MDLTracker = None) -> None:
        self.first_name = first_name
        self.last_name = last_name
        self.email = email
        self.discipline = discipline
        self.spreadsheet_work = spreadsheet_work
        self.discipline_leads = []



    


