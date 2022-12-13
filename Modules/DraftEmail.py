import os
import codecs
import win32com.client as win32
from datetime import datetime, timedelta
import calendar

class TimeStamp:
    now = datetime.now()
    today_date = int(now.strftime('%d'))
    starting_monday = now - timedelta(4)
    week_range = starting_monday.strftime('%d %b %Y') + ' to ' + now.strftime('%d %b %Y')
    last_day_of_month = calendar.monthrange(now.year, now.month)[1]

# access signature and add to email body
class EmailSignature:
    sig_files_path = 'AppData\Roaming\Microsoft\Signatures\Elizabeth Adejumo_py_files\\'
    sig_html_path = 'AppData\Roaming\Microsoft\Signatures\Elizabeth Adejumo_py.htm'

    @classmethod
    def get_signature(cls):
        # Finds the path to outlook signature files with signature name
        signature_path = os.path.join((os.environ['USERPROFILE']), cls.sig_files_path)
        # specifies the name of the HTML version  of the stored signature
        html_doc = os.path.join((os.environ['USERPROFILE']), cls.sig_html_path)
        html_doc = html_doc.replace('\\\\', '\\') # removes escape backlashes from path string
        html_file = codecs.open(html_doc, 'r', 'utf-8', errors='ignore') #Opens HTML file and converts to UTF-8, ignoring errors
        signature_code = html_file.read() #Writes contents of HTML signature file to a string
        signature_code = signature_code.replace(('Elizabeth Adejumo_files/'), signature_path) #Replaces local directory with full directory path
        html_file.close()
        return signature_code

class SendEmail:
    img_path = r'C:\Users\ukaea001\AppData\Roaming\Microsoft\Signatures\Elizabeth Adejumo_files\image001.png'

    @classmethod
    def draft_email(cls, email_body: str, subject: str, email_to: str, cc_contacts: str = None):
        try:
            mail = win32.Dispatch('outlook.application').CreateItem(0)
            mail.To, mail.cc, mail.Subject = email_to, cc_contacts, subject
            signature_code = EmailSignature.get_signature()
            mail.HTMLBody = email_body + signature_code
            # Adding signature image
            inspector = mail.getInspector
            doc = inspector.WordEditor
            selection = doc.Content
            selection.Find.Text = "insert image"
            selection.Find.Execute()
            selection.Text = ""
            img = selection.InlineShapes.AddPicture(cls.img_path, 0, 1)
            mail.display()
            # mail.Send()
            return True
            # raise exception instead! ----------------------------------------------
        except Exception as e:
            print("Invoice email alert failed to send: " + str(e) + ". You need to open the outlook application to ensure email sends.")
                
class InvoiceEmail(SendEmail):
    def __init__(self) -> None:
        super().__init__()

    @classmethod
    def send_invoice_reminder(cls, first_name: str, discipline: str, email_to: str, cc_contacts: str = None):
        previous_month = calendar.month_name[TimeStamp.now.month-1]
        subject = f'PLMB {discipline} - Invoice Reminder for {previous_month} ' + TimeStamp.now.strftime('%Y')
        email_body = r'''
            Hi {2},<br><br>
            This is a reminder that we are still awaiting your invoice for <b>{0} {1}</b>. <br><br>
            Could you please send this invoice as soon as possible.
            Please ensure that when you send through your invoices, you include a copy of the timesheets to back up the invoice. <br><br>
            Please ignore this email if you have sent your invoice previously. <br><br>
            Thanks.<br><br>
            '''.format(previous_month, TimeStamp.now.year, first_name)
        outcome = cls.draft_email(email_body = email_body, subject = subject, email_to = email_to, cc_contacts=cc_contacts)
        return outcome


    
    
