#!usr/bin/env python3
# XometryParsePDF.py - script that parses PDF files from Xometry and returns certain values as a table.

import PyPDF2
import openpyxl
import os
import logging
import re
import shutil

logging.basicConfig(level=logging.DEBUG, format='%(asctime)s.%(msecs)03d: %(message)s', datefmt='%H:%M:%S')
logging.basicConfig(level=logging.INFO, format='%(asctime)s.%(msecs)03d: %(message)s', datefmt='%H:%M:%S')
# logging.disable(logging.DEBUG)  # un/comment to un/block debug log messages
# logging.disable(logging.INFO)  # un/comment to un/block info log messages


def read_document(abs_folder_path):
    """ Goes through and reads all files ending in '.pdf' in the given directory with PyPDF.
        Depending on document content, calls the appropriate function to process. """
    # Change working directory to provided path argument.
    os.chdir(abs_folder_path)

    # Pattern check for purchase orders.
    po_regex = re.compile(r'.*(PURCHASE ORDER).*7951.*', re.DOTALL)

    # Pattern check for travelers.
    traveler_regex = re.compile(r'.*(Purchase Order)Due.*', re.DOTALL)

    # Searches all files in the provided directory. Does not include folders.
    for file in os.listdir('.'):
        logging.debug(f'Checking file: {file}')

        # Only interested in files that are PDF filetypes.
        if file.endswith('.pdf'):

            # Opens file with PyPDF2 and extracts information from the first page to determine document type.
            logging.info(f'{file} is a PDF, opening contents to check document type.')
            pdf_file_obj = open(file, 'rb')
            pdf_reader = PyPDF2.PdfFileReader(pdf_file_obj)
            page_obj = pdf_reader.getPage(0)
            sort_page = page_obj.extractText()

            # Closes file after extracting information needed.
            pdf_file_obj.close()
            logging.debug(f'read_document page 1 parse: \n{sort_page}')

            po_match = po_regex.match(sort_page)
            traveler_match = traveler_regex.match(sort_page)

            # If initial characters match 'PURCHASE', file is a PO.
            if po_match is not None:
                logging.info(f'{file} is an Xometry Purchase Order')
                purchase_order_process(file)

            # If initial characters match 'Purchase', file is a Customer Traveler.
            if traveler_match is not None:
                logging.info(f'{file} is an Xometry Traveler')
                traveler_process(file)


def traveler_process(filename):
    """ Opens traveler with PyPDF and sorts information into variables.
        Passes appropriate variables into rename_drawings and rename_traveler. """
    logging.info(f'Processing the following Xometry Traveler: {filename}')

    # Regex pattern, probably the part that will need editing the most to fit all traveler patterns.
    pattern = re.compile(r'''
        .*?DateContact
        ([a-zA-Z]\w{7})                                     # PO Number
        (\d\d/\d\d/\d\d\d\d)                                # Due Date
        (.*?@.*?\.com)                                      # Contact
        .*?
        Quantity
        (0\w{6})                                            # Part ID
        (.*?)                                               # Part name
        (\.sldprt|\.SLDPRT|\.step|\.STEP|\.stp|\.STP)       # Part Name extension
        ''', re.VERBOSE | re.DOTALL)

    # Open traveler with PyPDF2, this time reading everything since we know it's the document we are looking for.
    parse_string = open_parse_pdf(filename)

    matches = pattern.match(parse_string)
    logging.debug(f'PO Number is: {matches.group(1)}')
    logging.debug(f'Due Date is: {matches.group(2)}')
    logging.debug(f'Contact is: {matches.group(3)}')
    logging.debug(f'Job Number is: {matches.group(4)}')
    job_number = matches.group(4)
    logging.debug(f'Part File is: {matches.group(5) + matches.group(6)}')

    # Rename drawings after we have the provided Job Number.
    logging.info(f'Sending {job_number} to "rename_drawings"')
    rename_drawings(job_number)

    # Rename travelers after we have the provided Job Number and Traveler filename.
    logging.info(f'Sending {filename} and {job_number} to "rename_traveler"')
    rename_traveler(filename, job_number)


def purchase_order_process(filename):
    """ Opens purchase order with PyPDF and sorts information into variables.
        Passes appropriate variables into rename_drawings and rename_traveler. """
    logging.info(f'Processing the following Xometry PO: {filename}')

    # Regex pattern to grab Part ID (also used as Job Number) from purchase orders.
    part_id_pattern = re.compile(r'(Qty\.)(\n)(.*?)(\w{7})(\n)', re.DOTALL)

    # Open traveler with PyPDF2, this time reading everything since we know it's the document we are looking for.
    parse_string = open_parse_pdf(filename)

    # Create match group for job_number, then call rename_drawings
    job_number_match = part_id_pattern.search(parse_string)
    job_number = job_number_match.group(4)
    rename_drawings(job_number)


def rename_drawings(job_number):
    """ Renames files to remove long string using the given Job Number from the Customer Traveler """
    logging.info(f'Renaming drawings for the following Job Number: {job_number}')

    # Pattern that looks for files with the provided Job_Number, drawing, and drawing filetype.
    drawing_pattern = re.compile(rf'({job_number}_r_drawing_d_)(.*)(r_\w).*(\.pdf|\.jpg|\.jpeg|\.PDF|\.JPG|\.JPEG)')

    # Searches through all files in the current working directory (directory provided).
    for orig_filename in os.listdir('.'):

        # If pattern matches filename, process the filename into regex match groups.
        matches = drawing_pattern.search(orig_filename)
        if matches is None:
            continue
        prefix = matches.group(1)
        long_alphanum = matches.group(2)
        suffix = matches.group(3)
        extension = matches.group(4)

        # If orig_filename does not have the long alphanumeric, it means file has already been renamed, skips.
        if long_alphanum == '':
            logging.info(f'{orig_filename} already renamed. Skipping.')
            continue

        # Loop through valid files, creating a file name with customer provided format.
        counter = 1
        new_file_name = prefix + suffix + ' (' + str(counter) + ')' + extension

        # If file name already exists, appends an incrementing number just before file extension.
        while os.path.isfile(new_file_name):
            counter += 1
            new_file_name = prefix + suffix + ' (' + str(counter) + ')' + extension

        # After creating the filenames, move the files.
        logging.info(f'Renaming "{orig_filename}" TO "{new_file_name}"')
        shutil.move(orig_filename, new_file_name)


def rename_unlinked_drawings(abs_folder_path):
    """ Renames files to remove long string using the given Job Number from the Customer Traveler """

    logging.info(f'Renaming unlinked drawings in the following folder: {abs_folder_path}')
    os.chdir(abs_folder_path)

    # Pattern that looks for files with the provided Job_Number, drawing, and drawing filetype.
    drawing_pattern = re.compile(rf'(.*)(_r_drawing_d_)(.*)(r_\w).*(\.pdf|\.jpg|\.jpeg|\.PDF|\.JPG|\.JPEG)')

    # Searches through all files in the current working directory (directory provided).
    for orig_filename in os.listdir('.'):

        # If pattern matches filename, process the filename into regex match groups.
        matches = drawing_pattern.search(orig_filename)
        if matches is None:
            continue
        part_id = matches.group(1)
        prefix = matches.group(2)
        long_alphanum = matches.group(3)
        suffix = matches.group(4)
        extension = matches.group(5)

        # If orig_filename does not have the long alphanumeric, it means file has already been renamed, skips.
        if long_alphanum == '':
            logging.info(f'{orig_filename} already renamed. Skipping.')
            continue

        # Loop through valid files, creating a file name with customer provided format.
        counter = 1
        new_file_name = part_id + prefix + suffix + ' (' + str(counter) + ')' + extension

        # If file name already exists, appends an incrementing number just before file extension.
        while os.path.isfile(new_file_name):
            counter += 1
            new_file_name = part_id + prefix + suffix + ' (' + str(counter) + ')' + extension

        # After creating the filenames, move the files.
        logging.info(f'Renaming "{orig_filename}" TO "{new_file_name}"')
        shutil.move(orig_filename, new_file_name)


def rename_traveler(original_traveler, job_number):
    """ Renames Customer Traveler files to match the following format: CT (Job Number).pdf """

    # Pattern that looks for optional 'CT ' followed by any "filename.pdf or .PDF" format.
    traveler_pattern = re.compile(r'(CT )?(.*)(\.pdf|\.PDF)')

    # Loops over all files in current working directory.
    for orig_filename in os.listdir('.'):
        matches = traveler_pattern.search(orig_filename)

        # If pattern for file does not match, skips.
        if matches is None:
            continue

        # If match group 1 (for 'CT ') is not empty, file has already been processed, leave function.
        if matches.group(1) is not None:
            logging.info(f'{orig_filename} has already been renamed. Skipping.')
            break

        prefix = matches.group(2)
        extension = matches.group(3)

        # If file name being passed in matches the name search exactly, we are working with the right file.
        if original_traveler == prefix + extension:

            # Create the new file name.
            new_file_name = 'CT ' + job_number + extension

            # Rename the file.
            logging.info(f'Renaming "{original_traveler}" TO "{new_file_name}"')
            shutil.move(original_traveler, new_file_name)


def open_parse_pdf(filename):
    pdf_file_obj = open(filename, 'rb')
    pdf_reader = PyPDF2.PdfFileReader(pdf_file_obj)
    num_pages = pdf_reader.numPages
    parse_string = ''
    for i in range(0, num_pages):
        page_obj = pdf_reader.getPage(i)
        parse_string += page_obj.extractText()

    # Close the document after getting all the information.
    pdf_file_obj.close()
    logging.debug(f'parse_string is: \n{parse_string}')
    return parse_string


def main():
    print('Press CTRL + C at anytime to exit. Closing the command line window will also work.')
    while True:
        folder_path = input('Please paste the absolute folder path with the files you wish to process: \n')
        logging.info(f'Getting data from the following directory: {folder_path}')
        read_document(folder_path)
        rename_unlinked_drawings(folder_path)
        print('Folder processed, please check files to make sure everything went accordingly.')


if __name__ == '__main__':
    main()
