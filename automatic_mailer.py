"""An Application that automatically send NanoWeb Journal Emails."""
import win32com.client as win32
import sys
import os

class Automatic_mailer:
    date = ""
    default_body = ""
    default_path = 'C:\\Journals\\'
    emails = 0
    who = []
    emails = 0
    output_string = ""
    recipients = []

    def __init__(self, date):
        self.date = date
        emails_file = open("C:\\Journals\\emails.txt", "r")
        name = emails_file.readline().rstrip()
        text = emails_file.readlines()
        for line in text:
            line = line.rstrip()
            line = line.split(', ')
            i = 2
            cc = []
            while(i < len(line)):
                cc.append(line[i])
                i+=1
            try:
                self.recipients.append({'email': line[1], 'name': line[0], 'cc': cc})
            except:
                self.output_string += "There are empty lines in your emails file...Delete those."
        self.default_body = ("\n\nAttached are the invoices for your group’s usage of the"
        "NanoWeb facility for {}. "
        "Each user with an invoice has a tab located at the "
        "bottom of the spreadsheet.  Please review each tab and verify the cost center "
        "numbers are correct.\n"
        "Invoicing will be handled through the auto JE system. "
        "If you have any questions about this invoice please feel free to contact me. "
        "If a cost center does need to be changed please let me know as soon as "
        "possible, otherwise the supplied cost center(s) will be charged directly.\n\n"
        "Thank you,\n"
        "{}\n"
        "NanoTech Student Worker\n"
        ).format(date, name)


    def send_mail(self, to, cc, title, body, attachments):
        """Send an Email."""
        default_at = "@utdallas.edu"
        olMailItem = 0
        try:
            ol = win32.gencache.EnsureDispatch('Outlook.Application')
        except:
            print("Open Outlook and try the program again")
            self.output_string += "Open Outlook before running this."
            exit()
        msg = ol.CreateItem(olMailItem)
        try:
            msg.Recipients.Add(to + default_at)
        except:
            self.output_string += "ERROR: Open Outlook before running this."
        msg.Subject = title
        msg.Body = body
        for ccs in cc:
            x = msg.Recipients.Add(ccs + default_at)
            x.Type = 2
        for attachment in attachments:
            try:
                msg.Attachments.Add(attachment)
                self.who.append(to)
            except:
                self.output_string += "No attachment for {}...skipping".format(to) + "\n"
                return
        msg.Send()
        self.emails += 1


    def run(self):
        """Start the automatic mailer."""
        print("Looking for atachments in {}".format(self.default_path))
        for recipient in self.recipients:
            temp = recipient['name'].split(' ')
            last_name = temp[1]
            body = "Hello Dr. " +  last_name + "," + self.default_body
            to = recipient['email']
            title = "NanoWeb Invoices {}".format(self.date)
            attachment = self.default_path + "\\" + self.date + "\\" + self.date + " - "
            attachment += recipient['name'] + '.xls'
            try:
                cc = recipient['cc']
            except:
                cc = []
            self.send_mail(to, cc, title, body, [attachment])

        self.output_string += "Sent {} emails\n".format(self.emails)
        if self.emails > 0:
            self.output_string += "Emails Sent to:"
            for person in self.who:
                self.output_string += "\n" + person
        if self.emails == 0:
            self.output_string += "\nHuh...Didn't send any emails."
            self.output_string += "\nMake sure the Journals folder is in 'Local Disk (C)'."
            self.output_string += "\nAlso make sure you typed in the correct journal folder name and make sure Outlook is running."

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Needs a date argument")
        exit()
    elif len(sys.argv) > 2:
        print("Too many arguments: only use a date")
        exit()
    else:
        a = Automatic_mailer(sys.argv[1])
        a.run()
        print(a.output_string)
