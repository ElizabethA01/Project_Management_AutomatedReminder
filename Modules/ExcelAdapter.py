import pandas as pd

class ExcelAdapter:
    def __init__(self, contacts_list_file: str) -> None:
        self.contacts_list_file = contacts_list_file

    def extract_data_to_df(self):
        df = pd.read_excel(self.contacts_list_file, na_values='Missing')
        return df
    
    def df_to_list(self):
        df = self.extract_data_to_df()
        df_list = df.values.tolist()
        return df_list
    


