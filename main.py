import pandas as pd
import csv
import string
import random
import os
import shutil
import num2words
from mailmerge import MailMerge
from datetime import date
import locale
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders


def send_email_using_sendinblue(files_for_attachment, borrower_name, loan_scheme):
    mail_content = '''Hello,
    In this mail we are sending filled documents for loans.
    This is a system generated mail so donot-reply to this mail.
    Thank You
    Documation.in
    '''

    smtp_server_port = 'smtp-relay.sendinblue.com:587'
    smtp_user = 'documentcreator@documation.in'
    smtp_pass = 'awOs1V50QnYd29EU'

    receiver_address = "iob2316@iob.in"
    message = MIMEMultipart()
    message['From'] = smtp_user
    message['To'] = receiver_address
    message['Subject'] = '{} - {}'.format(borrower_name, loan_scheme)
    message.attach(MIMEText(mail_content, 'plain'))
    for filename in files_for_attachment:
        attachment = open(filename, 'rb')
        part = MIMEBase("application", "octet-stream")
        part.set_payload(attachment.read())
        encoders.encode_base64(part)
        filename = filename.split("\\")[-1]
        part.add_header("Content-Disposition", f"attachment; filename={filename}")
        message.attach(part)
    message = message.as_string()
    server = smtplib.SMTP(smtp_server_port)
    server.ehlo()
    server.starttls()
    server.login(smtp_user, smtp_pass)
    server.sendmail(smtp_user, receiver_address, message)
    server.quit()
    print('Email sent successfully')


basedir = os.path.abspath(os.path.dirname(__file__))
locale.setlocale(locale.LC_MONETARY, 'en_IN')
files_for_attachment = []
global scheme_code


def get_initial_for_folder_making(loan_scheme):
    folder_initial_names = {1: 'kccdc_coborrower_',
                            2: 'kccdc_',
                            3: 'kccdc_guar_',
                            4: 'pmsva_',
                            5: 'ccwms_',
                            6: 'pension_loan_'
                            }
    return folder_initial_names[loan_scheme]


def get_folder_for_scheme(loan_scheme):
    scheme_document_folder = {
                              2: os.path.join(basedir, 'kcc_document_with_coborrower'),
                              1: os.path.join(basedir, 'kcc_document_single_borrower'),
                              3: os.path.join(basedir, 'kcc_document_with_guarantor'),
                              4: os.path.join(basedir, 'pmsva_single_borrower'),
                              5: os.path.join(basedir, 'ccwms'),
                              6: os.path.join(basedir, 'pension_loan')
                              }
    return scheme_document_folder[loan_scheme]


def copy_proper_documents_to_be_filled_for_scheme_in_root_folder(document_directory):
    loan_documents = [f for f in os.listdir(document_directory) if f.endswith('.docx')]
    for document in loan_documents:
        shutil.copy(os.path.join(document_directory, document), os.path.join(basedir, document))
    return True


def documents_to_be_filled_for_scheme(scheme='default'):
    loan_documents_to_be_filled = [f for f in os.listdir(basedir) if f.endswith('.docx')]
    return loan_documents_to_be_filled


def id_generator(size=6, chars=string.ascii_uppercase + string.digits):
    return ''.join(random.choice(chars) for _ in range(size))


def get_indian_date(date_to_be_formatted):
    year, month, day = date_to_be_formatted.split('-')
    return '{:%d-%b-%Y}'.format(date(month=int(month), day=int(day), year=int(year)))


def generate_filled_document(template_document, destination_document, document_merge_fields):
    with MailMerge(template_document) as loan_document:
        loan_document.merge_templates([document_merge_fields], separator='page_break')
        loan_document.write(destination_document)
        return destination_document


def create_destination_directory(borrower_name, sb_column):
    destination_directory = os.path.join(basedir, get_initial_for_folder_making(scheme_code)+borrower_name + '_' + sb_column)
    if not os.path.exists(destination_directory):
        os.makedirs(destination_directory)
        print("Destination Directory Created Successfully:{}".format(destination_directory))
    else:
        print("Destination Directory already present")
    return destination_directory


def create_csv_for_document_generation_from_excel_sheet(excel_sheet_name=None):
    scheme_folder = get_folder_for_scheme(loan_scheme=scheme_code)
    if excel_sheet_name is None:
        excel_sheet_name = os.path.join(scheme_folder, "auto_fill.xlsm")
    new_data = pd.read_excel(excel_sheet_name, sheet_name='MERGEFIELDS')
    new_csv_name = get_initial_for_folder_making(scheme_code)+id_generator() + '.csv'
    new_data.to_csv(new_csv_name)
    return new_csv_name


def required_documents_merge_fields_to_fill_the_documents():
    merge_fields_of_templates = set()
    required_documents = documents_to_be_filled_for_scheme()
    for req_doc in required_documents:
        with MailMerge(req_doc) as document:
            fields = document.get_merge_fields()
            for field in fields:
                merge_fields_of_templates.add(field)
    merge_fields_of_templates = dict.fromkeys(merge_fields_of_templates, '')
    return merge_fields_of_templates


def delete_copied_files():
    loan_documents_copied = [f for f in os.listdir(basedir) if f.endswith('.docx')]
    for document in loan_documents_copied:
        os.remove(document)


def data_for_filling_documents():
    pass


def print_menu():
    print("Select Documents to be generated from following  Menu")
    print("1. KCC with Single Borrower.")
    print("2. KCC with Co-borrower.")
    print("3. KCC with Guarantor.")
    print("4. PMSVA single borrower")
    print("5. Cash credit Weavers Mudra")
    print("6. Pension Loan")


def generate_ccwms_documents():
    csv_name = create_csv_for_document_generation_from_excel_sheet()
    with open(csv_name, 'r') as file:
        reader = csv.DictReader(file)
        for row in reader:
            merge_fields = dict(row)
            if 'borrower_name' in merge_fields.keys() and 'sb_column' in merge_fields.keys():
                loan_document_directory = create_destination_directory(merge_fields['borrower_name'], merge_fields['sb_column'])



print("____________________________")
print("|Document Generation Program|")
print("____________________________")
print_menu()
scheme_code = int(input('Enter Scheme Serial Number'))
csv_name = create_csv_for_document_generation_from_excel_sheet()
with open(csv_name, 'r') as file:
    reader = csv.DictReader(file)
    try:
        for row in reader:
            merge_fields = dict(row)
            print(merge_fields)
            merge_fields['loan_date'] = get_indian_date(merge_fields['loan_date'])
            merge_fields['dob'] = get_indian_date(merge_fields['dob'])
            loan_amount_in_words = 'Rupees '+''.join(num2words.num2words(merge_fields['loan_amount'], lang='en_IN').split(',')) +' only'
            merge_fields['amount_in_words'] = loan_amount_in_words.title()
            #merge_fields['emi_start_date'] = get_indian_date(merge_fields['emi_start_date'])
            #merge_fields['co_dob'] = get_indian_date(merge_fields['co_dob'])
            #merge_fields['g_dob'] = get_indian_date(merge_fields['g_dob'])
            merge_fields['loan_amount'] = locale.currency(int(merge_fields['loan_amount']), grouping=True).replace('?', '')
            #merge_fields['emi'] = locale.currency(int(merge_fields['emi']), grouping=True).replace('?', '')
            merge_fields['scale_of_finance'] = locale.currency(int(merge_fields['scale_of_finance']), grouping=True).replace('?', '')
            directory = create_destination_directory(merge_fields['borrower_name'], merge_fields['sb_column'])
            document_folder = get_folder_for_scheme(loan_scheme=scheme_code)
            print(document_folder)
            copy_proper_documents_to_be_filled_for_scheme_in_root_folder(document_folder)
            for doc in documents_to_be_filled_for_scheme():
                doc_path = os.path.join(directory, doc)
                print(doc_path)
                generate_filled_document(doc, doc_path, merge_fields)
                files_for_attachment.append(doc_path)
    except Exception as e:
        print(e)
        print("There is error in generating the documents. Try to rectify the error.")
    else:
        send_email_using_sendinblue(files_for_attachment, merge_fields['borrower_name'], ' '.join(get_folder_for_scheme(loan_scheme=scheme_code).split('\\')[-1].split('_')))
    finally:
        delete_copied_files()

