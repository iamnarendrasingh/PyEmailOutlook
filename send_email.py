import win32com.client as win32
import sys
import openpyxl

def send_email_with_outlook(recipient, cc_recipients, subject, body, attachment_path):
    outlook = win32.Dispatch("Outlook.Application")
    mail = outlook.CreateItem(0)  # 0 represents an email item

    # Set email properties
    mail.To = recipient
    mail.CC = ";".join(cc_recipients)
    mail.Subject = subject
    mail.Body = body

    if attachment_path:
        # Add attachment to the email
        attachment = mail.Attachments.Add(attachment_path)
        attachment.DisplayName = "Email_Automation_PDF.txt"  # Set the displayed name of the attachment

    # Send the email
    mail.Send()
    print("Email sent successfully!")

def read_email_details_from_excel(filename):
    wb = openpyxl.load_workbook(filename)
    sheet = wb.active

    for row in sheet.iter_rows(min_row=2, values_only=True):
        recipient, cc_recipients, subject, body, attachment_path = row
        cc_recipients = cc_recipients.split(";") if cc_recipients else []
        send_email_with_outlook(recipient, cc_recipients, subject, body, attachment_path)

    wb.save("Sent_Emails.xlsx")
    print("Emails sent and Excel sheet updated.")

if __name__ == "__main__":
    if len(sys.argv) != 2:
        print("Usage: python send_email.py <excel_filename>")
        sys.exit(1)

    excel_filename = sys.argv[1]
    read_email_details_from_excel(excel_filename)
