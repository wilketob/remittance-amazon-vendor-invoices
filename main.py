#!/bin/python3

import json
import os.path
import imaplib
import email.message
import datetime
import time
import pandas as pd
import lxml
import pdfkit  # Don't forget to install wkhtmltopdf: sudo apt install wkhtmltopdf
from PyPDF2 import PdfFileMerger, PdfFileReader

f = open('settings.json',)  # Opening JSON file with setting patams

# returns JSON object as a dictionary
json_data = json.load(f)
f.close()  # Closing file


def main():

    # create an IMAP4 class with SSL
    imap = imaplib.IMAP4_SSL(json_data["server_in"])
    # authenticate
    imap.login(json_data["username"], json_data["password"])
    status, messages = imap.select("INBOX")
    # search for messages that has the word remittance in subject
    typ, msg_ids = imap.search(None, '(SUBJECT "Remittance")')
    print(f'Status: {status} Typ: {typ} Messages: {msg_ids}')
    # msg_ids is byte -> convert to string and split the string by whitespace
    for msg_id in msg_ids[0].decode().split():
        # print(type(msg_id))
        typ, msg_data = imap.fetch(msg_id, '(BODY.PEEK[TEXT] FLAGS)')
        # print(msg_data)

        msg_str = str(msg_data[0][1]).lower()  # For better readibility

        # read the html tables from the document with pandas:
        tables_html = pd.read_html(msg_str, skiprows=1)
        payment_n = tables_html[0][1][2]
        payment_d = tables_html[0][1][3]
        supplier_site = tables_html[0][1][1]
        print(f'#### Start {payment_n} ###')
        # convert format 01-Sep-2020 to 2020-09-01
        payment_d = str(datetime.datetime.strptime(
            payment_d, '%d-%b-%Y').strftime('%Y-%m-%d'))

        # check if file A-2020-09-01-remittance number exists
        if not os.path.exists(json_data["pwd"] + 'A-' + payment_d + '-' + payment_n + '.pdf'):
            pdfkit.from_string(msg_data[0][1].decode(
            ), json_data["pwd"] + 'A-' + payment_d + '-' + payment_n + '.pdf')
            print(
                f'[++] The pdf file for remittance {payment_n} has been created')
        else:
            print(
                f'[++] The pdf file for remittance {payment_n} does already exists')

        # print(payment_n)
        # print(payment_d)
        # print(type(payment_d))
        # Get the files of working dic

        # list for existing invoices
        list_success = ['A-' + payment_d + '-' + payment_n + '.pdf']

        # list for missing invoices
        list_missing = []

        file_list = os.listdir(json_data['pwd'])

        # hol die RGnr der 2. Tabelle und pack sie in die Liste and check if exists locall

        for inv_nr in tables_html[1][0]:
            i = 0
            for fn in file_list:
                if(str(inv_nr).upper() in fn):
                    list_success.append(fn)
                    list_missing.remove(inv_nr)
                    i += i
                else:
                    if(i == 0):
                        list_missing.append(inv_nr)
                        i = i+1
            # create list with needed invoices

        # check if missing invoices exists
        print(f'[--] MISSING {payment_n} {list_missing}')
        print(f'[++] SUCCESS {payment_n} {list_success}')
        if len(list_missing) > 0:
            new_message = email.message.Message()
            new_message["From"] = "hello@itsme.com"
            new_message["Subject"] = json_data['email_subject'] + \
                payment_n + supplier_site
            new_message.set_payload(','.join(map(str, list_missing)))
            imap.append('INBOX', '', imaplib.Time2Internaldate(
                time.time()), str(new_message).encode('utf-8'))

        else:
            if not os.path.exists(json_data["pwd"] + "A-AMZ-" + payment_d + "-" + payment_n + ".pdf"):
                # create invoice with pdftk and store it
                # Call the PdfFileMerger
                mergedObject = PdfFileMerger(strict=False)
                for list_item in list_success:
                    mergedObject.append(PdfFileReader(
                        json_data["pwd"] + list_item, 'rb'))

                # Write all the files into a file which is named as shown below
                mergedObject.write(
                    json_data["pwd"] + "A-AMZ-" + payment_d + "-" + payment_n + ".pdf")

                # delete email

                imap.store(msg_id, '+FLAGS', '\\Deleted')
                imap.expunge()

                # delete files
                for del_item in list_success:
                    if os.path.exists(json_data["pwd"] + del_item):
                        os.remove(json_data["pwd"] + del_item)
        print(f'#### END {payment_n} ###\n')

    try:
        imap.close()
    except:
        pass
    imap.logout()

    return 0


if __name__ == '__main__':
    main()