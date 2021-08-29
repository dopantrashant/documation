import pandas as pd
import csv
import string
import random
import os
from mailmerge import MailMerge
from datetime import date
import locale

basedir = os.path.abspath(os.path.dirname(__file__))
loan_documents_to_be_filled = [f for f in os.listdir(basedir) if f.endswith('.docx')]
locale.setlocale(locale.LC_MONETARY, 'en_IN')


def id_generator(size=6, chars=string.ascii_uppercase+string.digits):
    return ''.join(random.choice(chars) for _ in range(size))


def get_indian_date(date_to_be_formatted):
    year, month, day = date_to_be_formatted.split('-')
    return '{:%d-%b-%Y}'.format(date(month=int(month), day=int(day), year=int(year)))


excel_file = "C:\\Users\\User\\PycharmProjects\\Documation\\kcc_document_with_coborrower\\kcc_with_coborrower.xlsm"
new_data = pd.read_excel(excel_file, sheet_name='MERGEFIELDS')
csv_name = id_generator()+'.csv'
new_data.to_csv(csv_name)
print("Document Generation Program")
with open(csv_name, 'r') as file:
    reader = csv.DictReader(file)
    for row in reader:
        merge_fields = dict(row)
        year, month, day = merge_fields['loan_date'].split('-')
        print(merge_fields)
        merge_fields['loan_date'] = get_indian_date(merge_fields['loan_date'])
        merge_fields['dob'] = get_indian_date(merge_fields['dob'])
        merge_fields['co_dob'] = get_indian_date(merge_fields['co_dob'])
        merge_fields['loan_amount'] = locale.currency(int(merge_fields['loan_amount']), grouping=True).replace('?','')
        directory = os.path.join(basedir, merge_fields['borrower_name'] + '-' + merge_fields['sb_column'])
        if not os.path.exists(directory):
            os.mkdir(directory)
        print(directory)
        for doc in loan_documents_to_be_filled:
            doc_path = None
            with MailMerge(doc) as document:
                document.merge_templates([merge_fields], separator='page_break')
                doc_path = os.path.join(directory, doc)
                print(doc_path)
                document.write(doc_path)


