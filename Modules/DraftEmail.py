import os
import codecs
import win32com.client as win32
from datetime import datetime, timedelta
import calendar


# import relevant files
from disciplines_details import DisciplineLead
disciplines_leads = DisciplineLead()

class TimeStamp:
    now = datetime.now()
    today_date = int(now.strftime('%d'))
    starting_monday = now - timedelta(4)
    week_range = starting_monday.strftime('%d %b %Y') + ' to ' + now.strftime('%d %b %Y')
    last_day_of_month = calendar.monthrange(now.year, now.month)[1]

    # do you need to do def __init__ with variables above?

# access signature and add to email body
class EmailSignature:
    sig_files_path = 'AppData\Roaming\Microsoft\Signatures\Elizabeth Adejumo_py_files\\'
    sig_html_path = 'AppData\Roaming\Microsoft\Signatures\Elizabeth Adejumo_py.htm'
    img_path = r'C:\Users\ukaea001\AppData\Roaming\Microsoft\Signatures\Elizabeth Adejumo_files\image001.png'

    def __init__(self, sig_files_path, sig_html_path, img_path ) -> None:
        self.sig_files_path = sig_files_path
        self.sig_html_path = sig_html_path
        self.img_path = img_path

    def get_signature(self):
        # Finds the path to outlook signature files with signature name
        signature_path = os.path.join((os.environ['USERPROFILE']), self.sig_files_path)
        # specifies the name of the HTML version  of the stored signature
        html_doc = os.path.join((os.environ['USERPROFILE']), self.sig_html_path)
        html_doc = html_doc.replace('\\\\', '\\') # removes escape backlashes from path string

        html_file = codecs.open(html_doc, 'r', 'utf-8', errors='ignore') #Opens HTML file and converts to UTF-8, ignoring errors
        signature_code = html_file.read() #Writes contents of HTML signature file to a string

        signature_code = signature_code.replace(('Elizabeth Adejumo_files/'), signature_path) #Replaces local directory with full directory path
        html_file.close()
        return signature_code

class SendEmail:
    time_stamp = TimeStamp
    email_signature = EmailSignature
    mail = win32.Dispatch('outlook.application').CreateItem(0)
    mailCC_contacts = 'carolina.morales@wsp.com' # need to modify and add to ignore file

    def __init__(self, time_stamp, mail, email_signature, mailCC_contacts) -> None:
        self.time_stamp = time_stamp
        self.mail = mail
        self.email_signature = email_signature
        self.mailCC_contacts = mailCC_contacts

    def draft_email(self, email_body, subject):
        try:
            self.mail.To = disciplines_leads.email
            self.mail.cc = self.mailCC_contacts
            signature_code = self.email_signature.get_signature
            self.mail.Subject = subject
            self.mail.HTMLBody = email_body + signature_code
            self.add_signature()
            self.mail.display()
            # self.mail.Send()
        except Exception as e:
            print("Invoice email alert failed to send: " + str(e))

    def add_signature(self):
        # Adding signature image
        inspector = self.mail.getInspector
        doc = inspector.WordEditor
        selection = doc.Content
        selection.Find.Text = "insert image"
        selection.Find.Execute()
        selection.Text = ""
        img = selection.InlineShapes.AddPicture(self.email_signature.img_path, 0, 1)


class InvoiceEmail(SendEmail):
    def __init__(self, time_stamp, mail, email_signature) -> None:
        super().__init__(time_stamp, mail, email_signature)

    def send_invoice_reminder(self):
        previous_month = calendar.month_name[self.time_stamp.now.month-1]
        subject = f'PLMB {disciplines_leads.discipline} - Invoice Reminder for {previous_month} ' + self.time_stamp.now.strftime('%Y')
        email_body = r'''
            Hi {2},<br><br>
            This is a reminder that we are still awaiting your invoice for <b>{0} {1}</b>. <br><br>
            Could you please send this invoice as soon as possible.
            Please ensure that when you send through your invoices, you include a copy of the timesheets to back up the invoice. <br><br>
            Please ignore this email if you have sent your invoice previously. <br><br>
            Thanks.<br><br>
            '''.format(previous_month, self.time_stamp.now.year, disciplines_leads.first_name)
        self.draft_email(email_body=email_body, subject=subject)
    
    

