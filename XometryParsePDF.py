#!usr/env/bin python3
# XometryParsePDF.py - script that parses PDF files from Xometry and returns certain values as a table.

import PyPDF2
import openpyxl
import os
import logging
import re
import shutil

logging.basicConfig(level=logging.DEBUG, format='%(asctime)s.%(msecs)03d: %(message)s', datefmt='%H:%M:%S')
logging.basicConfig(level=logging.INFO, format='%(asctime)s.%(msecs)03d: %(message)s', datefmt='%H:%M:%S')
# logging.disable(logging.DEBUG) # un/comment to un/block debug log messages
# logging.disable(logging.INFO) # un/comment to un/block info log messages


def read_document(abs_folder_path):
    """ Goes through and reads all files ending in '.pdf' in the given directory with PyPDF.
        Depending on document content, calls the appropriate function to process. """
    os.chdir(abs_folder_path)
    for file in os.listdir('.'):
        logging.debug(f'Checking file: {file}')
        if file.endswith('.pdf'):
            logging.debug(f'{file} is a PDF, opening contents to redirect.')
            pdf_file_obj = open(file, 'rb')
            pdf_reader = PyPDF2.PdfFileReader(pdf_file_obj)
            page_obj = pdf_reader.getPage(0)
            sort_page = page_obj.extractText()
            pdf_file_obj.close()
            if sort_page[0:8] == 'PURCHASE':
                logging.info(f'{file} is an Xometry Purchase Order')
            if sort_page[1:9] == 'Purchase':
                logging.info(f'{file} is an Xometry Traveler')
                traveler_process(file)


def traveler_process(filename):
    """ Opens traveler with PyPDF and sorts information into variables.
        Passes appropriate variables into rename_drawings and rename_traveler. """
    logging.info(f'Processing the following Xometry Traveler: {filename}')
    pattern = re.compile(r'''
        .*DateContact
        ([a-zA-Z]\w{7})                                     # PO Number
        (\d\d\/\d\d\/\d\d\d\d)                              # Due Date
        (.*@.*.com)                                         # Contact
        .*
        (0\w{6})                                            # Part ID
        (.*)                                                # Part name
        (\.sldprt|\.SLDPRT|\.step|\.STEP|\.stp|\.STP)       # Part Name extension
        ''', re.VERBOSE | re.DOTALL)
    pdf_file_obj = open(filename, 'rb')
    pdf_reader = PyPDF2.PdfFileReader(pdf_file_obj)
    numPages = pdf_reader.numPages
    parse_string = ''
    for i in range(0, numPages):
        page_obj = pdf_reader.getPage(i)
        parse_string += page_obj.extractText()

    pdf_file_obj.close()

    logging.debug(f'parse_string is: {parse_string}')
    matches = pattern.match(parse_string)
    logging.debug(f'PO Number is: {matches.group(1)}')
    logging.debug(f'Due Date is: {matches.group(2)}')
    logging.debug(f'Contact is: {matches.group(3)}')
    logging.debug(f'Job Number is: {matches.group(4)}')
    job_number = matches.group(4)
    logging.debug(f'Part File is: {matches.group(5)}')

    rename_drawings(job_number)
    rename_traveler(filename, job_number)

def rename_drawings(job_number):
    """ Renames files to remove long string using the given Job Number from the Customer Traveler """
    drawing_pattern = re.compile(rf'({job_number}_r_drawing_d_).*(r_0).*(.pdf|.jpg|.jpeg|.PDF|.JPG|.JPEG)')
    for orig_filename in os.listdir('.'):
        matches = drawing_pattern.search(orig_filename)
        if matches == None:
            continue
        prefix = matches.group(1)
        suffix = matches.group(2)
        extension = matches.group(3)

        counter = 1
        new_file_name = prefix + suffix + ' (' + str(counter) + ')' + extension
        if new_file_name == orig_filename:
            continue
        while os.path.isfile(new_file_name):
            counter += 1
            new_file_name = prefix + suffix + ' (' + str(counter) + ')' + extension
        logging.info(f'Renaming "{orig_filename}" TO "{new_file_name}"')
        shutil.move(orig_filename, new_file_name)


def rename_traveler(original_traveler, job_number):
    traveler_pattern = re.compile(r'(CT )?(.*)(\..*)')

    for orig_filename in os.listdir('.'):
        matches = traveler_pattern.search(orig_filename)
        if matches == None:
            continue
        if matches.group(1) != None:
            logging.info(f'{orig_filename} has already been renamed.')
            break
        prefix = matches.group(2)
        extension = matches.group(3)
        if original_traveler == prefix + extension:
            new_file_name = 'CT ' + job_number + extension
            logging.info(f'Renaming "{original_traveler}" TO "{new_file_name}"')
            shutil.move(original_traveler, new_file_name)


if __name__ == '__main__':
    folder_path = input('Please paste the absolute folder path with the files you wish to process: \n')
    logging.debug(f'Getting data from the following directory: {folder_path}')
    read_document(folder_path)
