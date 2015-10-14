__author__ = 'Jessica'

import smtplib
from email.mime import multipart
from email.mime import text
from email.mime import application
import getpass
import time
import argparse
import csv
import os
import sys

SENDER_LIST = {'Name': [], 'Email': [], 'Company': []}
OLD_LIST = []
LIMIT = 90


def get_details():
    """
    Get arguments including list of senders, the text of the email message, and attachments
    Send email and resume (or other attachment to Linked in employees
    :return:
    """
    parser = argparse.ArgumentParser(description='Send email attachment to recipients.')
    parser.add_argument('senderList', help='List of senders')
    parser.add_argument('attachment', help='Attachment')
    parser.add_argument('emailBody', help='Path to the text file containing email message')
    parser.add_argument('-ex', '--exclude', metavar='old_file',
                        help='Send mail to new recipients only')
    args = parser.parse_args()

    if args.exclude:
        if os.path.splitext(args.exclude)[1] == ".csv":
            with open(args.exclude, newline='') as old_csv_data:
                reader = csv.reader(old_csv_data)
                for row in reader:
                    if not row[1] == "First Name":
                        if row[1] and row[5] and row[29]:
                            OLD_LIST.append(row[5])
        else:
            raise Exception("Old sender's list should also be in .csv format")

    if os.path.splitext(args.sender_list)[1] == ".csv":
        with open(args.sender_list, newline='') as csv_data:
            reader = csv.reader(csv_data)
            for row in reader:
                if not row[1] == "First Name":
                    if OLD_LIST and row[5] in OLD_LIST:
                        print("Skipping {0} - {1}. Email was sent in the previous run.".format(row[1], row[5]))
                        continue
                    elif "Recruit" in row[31] or "Talent" in row[31] and row[1] and row[5] and row[29]:
                        SENDER_LIST['Name'].append(row[1].encode('ascii', 'ignore'))
                        SENDER_LIST['Email'].append(row[5].encode('ascii', 'ignore'))
                        SENDER_LIST['Company'].append(row[29].encode('ascii', 'ignore'))
        print()
    else:
        raise Exception("Sender's list should be in .csv format")

    if os.path.isfile(args.attach_path):
        attachment = args.attach_path
    else:
        raise FileExistsError('Attachment does not exist')

    user = input("First name: ")
    gmail_user = input("Username: ")
    gmail_pwd = getpass.getpass("Password: ")

    counter = 0
    if os.path.isfile(args.text_body):
        while True:
            while counter < LIMIT and SENDER_LIST.get('Name'):
                recipient = SENDER_LIST.get('Name').pop().decode()
                recipient_company = SENDER_LIST.get('Company').pop().decode()
                mail_to = SENDER_LIST.get('Email').pop().decode()
                with open(args.text_body, 'rb') as data:
                    content = data.read().decode().format(Name=recipient, Company=recipient_company, User=user)

                msg = multipart.MIMEMultipart()
                msg['From'] = gmail_user
                msg['To'] = str(mail_to)
                msg['Subject'] = "Full-time opportunities at {}".format(recipient_company)
                msg.attach(text.MIMEText(content))

                part = application.MIMEApplication(open(attachment, 'rb').read())
                part.add_header('Content-Disposition', 'attachment; filename="{}"'.format(os.path.basename(attachment)))
                msg.attach(part)

                try:
                    mailServer = smtplib.SMTP("Page on gmail.com", 587)
                    mailServer.ehlo()
                    mailServer.starttls()
                    mailServer.ehlo()
                    mailServer.login(gmail_user, gmail_pwd)
                    mailServer.sendmail(gmail_user, str(mail_to), msg.as_string())
                    # Should be mailServer.quit(), but that crashes...
                    mailServer.close()
                    print("Successfully sent the email to {}".format(recipient))

                except Exception as e:
                    print("Failed to send email to {}. Reason: {}".format(recipient, str(e)))
                counter += 1

            if not SENDER_LIST.get('Name'):
                break
            time.sleep(180)
            counter = 0
    else:
        raise FileExistsError('Message file does not exists')


if __name__ == "__main__":
    sys.exit(get_details())