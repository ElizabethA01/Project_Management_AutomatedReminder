import os
import codecs
import win32com.client as win32
from datetime import datetime, timedelta
import calendar
from email_validator import validate_email, EmailNotValidError
    
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
    now = datetime.now()

    @classmethod
    def draft_email(cls, email_body: str, subject: str, email_to: str, cc_contacts: str = None):
        try:
            if ValidateEmail.check_email(email_to) and ValidateEmail.check_email(cc_contacts) is True:
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
            return "Email alert failed to send: " + str(e)
                
class InvoiceEmail(SendEmail):
    def __init__(self) -> None:
        super().__init__()

    @classmethod
    def send_invoice_reminder(cls, first_name: str, discipline: str, email_to: str, cc_contacts: str = None):
        previous_month = calendar.month_name[cls.now.month-1]
        subject = f'PLMB {discipline} - Invoice Reminder for {previous_month} ' + cls.now.strftime('%Y')
        email_body = r'''
            Hi {2},<br><br>
            This is a reminder that we are still awaiting your invoice for <b>{0} {1}</b>. <br><br>
            Could you please send this invoice as soon as possible.
            Please ensure that when you send through your invoices, you include a copy of the timesheets to back up the invoice. <br><br>
            Please ignore this email if you have sent your invoice previously. <br><br>
            Thanks.<br><br>
            '''.format(previous_month, cls.now.year, first_name)
        outcome = cls.draft_email(email_body = email_body, subject = subject, email_to = email_to, cc_contacts=cc_contacts)
        return outcome
    
class TimesheetEmail(SendEmail):
    def __init__(self) -> None:
        super().__init__()

    @classmethod
    def send_friday_month_end(cls, first_name: str, discipline: str, email_to: str, cc_contacts: str = None):
        subject = f'PLMB {discipline} Timesheet reminder - Month end'
        email_body = r'''
            Hi {0},<br><br>
            This is a reminder that it is month end. Could you please send your timesheet for this week by <b>COB today</b>. <br><br>
            Please ignore this email if you have sent your timesheet previously. <br><br>
            Thanks.<br><br>
            '''.format(first_name)
        outcome = cls.draft_email(email_body = email_body, subject = subject, email_to = email_to, cc_contacts=cc_contacts)
        return outcome

    @classmethod
    def send_friday_alert(cls, first_name: str, discipline: str, email_to: str, cc_contacts: str = None):
        starting_monday = cls.now - timedelta(4)
        week_range = starting_monday.strftime('%d %b %Y') + ' to ' + cls.now.strftime('%d %b %Y')
        subject = f'PLMB {discipline} Timesheet reminder - ' + week_range
        email_body = r'''
            Hi {0},<br><br>
            This is a reminder to send your timesheet for {1} by <b>COB today</b>. <br><br>
            Please ignore this email if you have sent your timesheet previously. <br><br>
            Thanks.<br><br>
            '''.format(first_name, week_range)
        outcome = cls.draft_email(email_body = email_body, subject = subject, email_to = email_to, cc_contacts=cc_contacts)
        return outcome
    
    @classmethod
    def send_midweek_month_end(cls, first_name: str, discipline: str, email_to: str, cc_contacts: str = None):
        subject = f'PLMB {discipline} Timesheet reminder - Month end tomorrow' 
        email_body = r'''
            Hi {0},<br><br>
            This is a reminder that it is month end tomorrow. Could you please send your timesheet for this week by <b>12PM tomorrow</b>. <br><br>
            Thanks.<br><br>
            '''.format(first_name)
        outcome = cls.draft_email(email_body = email_body, subject = subject, email_to = email_to, cc_contacts=cc_contacts)
        return outcome

    @classmethod
    def send_timesheet_reminder(cls, first_name: str, discipline: str, email_to: str, cc_contacts: str = None):
        today_date = int(cls.now.strftime('%d'))
        last_day_of_month = calendar.monthrange(cls.now.year, cls.now.month)[1]
        if today_date == last_day_of_month - 1:
            outcome = cls.send_midweek_month_end(first_name, discipline, email_to, cc_contacts)
            alert = 'Month end'
        elif cls.now.strftime('%A') == 'Friday' and today_date != last_day_of_month:   
            if today_date + 1 == last_day_of_month or today_date + 2 == last_day_of_month or today_date + 3 == last_day_of_month:
                outcome = cls.send_friday_month_end(first_name, discipline, email_to, cc_contacts)
                alert = 'Month end (during weekend)'
            else:
                outcome = cls.send_friday_alert(first_name, discipline, email_to, cc_contacts)
                alert = 'Friday timesheet'
        else:
            outcome = "no timesheet reminder today"
            alert = ""
        return outcome, alert
        
class ValidateEmail(): 
    def check_email(email):
        try:
            v = validate_email(email)
            email = v["email"]
            return True
        except EmailNotValidError as e:
            raise AssertionError(f'{e} - {email}')
