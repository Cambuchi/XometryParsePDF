#!usr/bin/env python3
# XometryParsePDF.py - script that parses PDF files from Xometry and returns certain values as a table.
# command to package as an executable file:
# pyinstaller --onefile --add-data="templates/TravelerTemplate.xlsx;templates" XometryParsePDF.py

import sys
import PyPDF2
import openpyxl
import os
import logging
import re
import shutil
import numpy as np
from datetime import datetime, timedelta
from PIL import Image
from openpyxl.drawing.spreadsheet_drawing import AbsoluteAnchor
from openpyxl.drawing.xdr import XDRPoint2D, XDRPositiveSize2D
from openpyxl.utils.units import pixels_to_EMU
from win32com import client

logging.basicConfig(level=logging.DEBUG, format='%(asctime)s.%(msecs)03d: %(message)s', datefmt='%H:%M:%S')
logging.basicConfig(level=logging.INFO, format='%(asctime)s.%(msecs)03d: %(message)s', datefmt='%H:%M:%S')
# logging.disable(logging.DEBUG)  # uncomment to view debug log messages
# logging.disable(logging.INFO)  # uncomment to view info log messages


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
        ([a-zA-Z]\w{7})\n?                                                  # 1 PO Number
        (\d\d/\d\d/\d\d\d\d)                                                # 2 Due Date
        (.*?@.*?\.com)                                                      # 3 Contact
        .*?
        Quantity
        (0\w{6})                                                            # 4 Part ID
        (.*?)                                                               # 5 Part name
        (\.\n?sldprt|\.\n?SLDPRT|\.\n?step|\.\n?STEP|\.\n?stp|\.\n?STP|\.\n?x_t|\.\n?s\n?tp)        # 6 Extension
        .*?
        (\d+)                                                               # 7 Quantity
        .*?
        tions
        (.*?[a-z])                                                          # 8 Finish
        (\n?[A-Z\d].*?)                                                     # 9 Material
        (Cert.*?)                                                           # 10 Certifications
        Inspection.*?[a-z]
        ([A-Z].*?)                                                          # 11 Inspection Requirements
        (Features:.*)                                                       # 12 Notes
        ''', re.VERBOSE | re.DOTALL)

    # Open traveler with PyPDF2, this time reading everything since we know it's the document we are looking for.
    parse_string = open_parse_pdf(filename)

    # Match traveler information into groups with regex
    matches = pattern.match(parse_string)

    # Create dictionary and sort matches into keys for passing into traveler creation.
    traveler_dictionary = {}

    logging.debug(f'PO Number is: {matches.group(1)}')
    traveler_dictionary['po_number'] = matches.group(1)

    logging.debug(f'Due Date is: {matches.group(2)}')
    traveler_dictionary['due_date'] = matches.group(2)

    logging.debug(f'Contact is: {matches.group(3)}')
    traveler_dictionary['contact'] = matches.group(3)

    # Passes job number into rename_drawings and rename_traveler
    logging.debug(f'Job Number is: {matches.group(4)}')
    traveler_dictionary['job_number'] = matches.group(4)
    job_number = matches.group(4)

    # Sends filepath and job_number to grab image from pdf
    image_grab(filename, job_number)

    logging.debug(f'Part File is: {matches.group(5) + matches.group(6)}')
    traveler_dictionary['part_file'] = remove_newlines(matches.group(5) + matches.group(6))

    logging.debug(f'Quantity is: {matches.group(7)}')
    traveler_dictionary['quantity'] = matches.group(7)

    logging.debug(f'Finish is: {matches.group(8)}')
    traveler_dictionary['finish'] = regex_film_check(matches.group(8))

    logging.debug(f'Material is: {matches.group(9)}')
    traveler_dictionary['material'] = matches.group(9)

    logging.debug(f'Certifications required are: {matches.group(10)}')
    traveler_dictionary['certifications'] = regex_certificate_check(low_up(remove_newlines(matches.group(10))))

    logging.debug(f'Inspection requirements are: {matches.group(11)}')
    traveler_dictionary['inspection'] = low_up(remove_newlines(matches.group(11)))

    logging.debug(f'Notes are: {matches.group(12)}')
    traveler_dictionary['notes'] = matches.group(12)

    # Rename drawings after we have the provided Job Number.
    logging.info(f'Sending {job_number} to "rename_drawings"')
    rename_drawings(job_number)

    # Rename travelers after we have the provided Job Number and Traveler filename.
    logging.info(f'Sending {filename} and {job_number} to "rename_traveler"')
    rename_traveler(filename, job_number)

    logging.info(f'traveler_dictionary contains the following: \n{traveler_dictionary}')
    logging.info(f'Sending traveler information to create_excel.')
    create_excel(traveler_dictionary, os.getcwd())


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
    """ Renames files to remove long strings from drawing titles unrelated to Travelers/POs. """

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


def create_excel(traveler_dictionary, folder_path):
    """
    Creates an excel template that mimics the customer traveler so that we can pass in information from the PDF.
    :traveler_dictionary: dictionary with information from traveler sorted by category.
    :folder_path: path to the directory directly containing the traveler
    :return: None
    """
    # Gets the directory of the executable script, which contains the template we will open.
    script_path = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
    os.chdir(script_path)
    wb = openpyxl.load_workbook(resource_path('templates/TravelerTemplate.xlsx'))

    # Transfer the traveler dictionary values into the excel template.
    sheet = wb['Sheet1']
    sheet['A1'] = f'CT {traveler_dictionary["job_number"]}'
    sheet['A6'] = traveler_dictionary['po_number']
    sheet['B6'] = traveler_dictionary['due_date']
    sheet['C6'] = 'CUSTOMER'
    sheet['A8'] = traveler_dictionary['job_number']
    sheet['B8'] = traveler_dictionary['part_file'].strip('\n')
    sheet['C8'] = traveler_dictionary['quantity']
    sheet['A10'] = traveler_dictionary['finish'].strip('\n')
    sheet['B10'] = traveler_dictionary['material'].strip('\n')
    sheet['C10'] = traveler_dictionary['certifications'].strip('\n')
    sheet['A12'] = traveler_dictionary['inspection'].strip('\n')

    # Replace all of Xometry mentions with 'CUSTOMER'
    no_customer = re.sub(r'xometry|Xometry|XOMETRY', 'CUSTOMER', traveler_dictionary['notes'].strip('\n'))
    no_trademark = replace_trademark(no_customer)
    no_fluid = replace_fluid(no_trademark)
    final_notes = process_notes(no_fluid)
    sheet['A14'] = final_notes

    # Date modifications according to customer instructions
    date = traveler_dictionary['due_date']
    date_time_obj = datetime.strptime(date, "%m/%d/%Y")

    # 1) If “Finish” block says “Standard”, change due date to two business days before current due date.
    if traveler_dictionary['finish'] == 'Standard':
        sheet['B6'] = f'{date_time_obj.date() - timedelta(days=2):%m/%d/%Y}'

    # 2) If anything in the pdf mentions “mask” or “masking”, “heat treat”, “heat treating”, “harden”,
    # or “through harden” change due date to 7 business days before current due date
    post_process_pattern = re.compile(r'(mask|masking|heat treat|heat treating|harden|through harden)')
    matches = post_process_pattern.search(traveler_dictionary['finish'])
    if matches is not None:
        sheet['B6'] = f'{date_time_obj.date() - timedelta(days=7):%m/%d/%Y}'
    matches = post_process_pattern.search(traveler_dictionary['material'])
    if matches is not None:
        sheet['B6'] = f'{date_time_obj.date() - timedelta(days=7):%m/%d/%Y}'
    matches = post_process_pattern.search(traveler_dictionary['notes'])
    if matches is not None:
        sheet['B6'] = f'{date_time_obj.date() - timedelta(days=7):%m/%d/%Y}'

    # 3) If “Finish” block says “Custom” but there’s no mention of masking in rest of pdf,
    # change due date to 5 business days before current due date
    finish_pattern = re.compile(r'custom|CUSTOM|Custom')
    mask_pattern = re.compile(r'mask|MASK|masking|MASKING')
    finish_matches = finish_pattern.search(traveler_dictionary['finish'])
    if finish_matches is not None:
        sheet['B6'] = f'{date_time_obj.date() - timedelta(days=5):%m/%d/%Y}'
        mask_matches = mask_pattern.search(traveler_dictionary['finish'])
        if mask_matches is not None:
            sheet['B6'] = f'{date_time_obj.date() - timedelta(days=7):%m/%d/%Y}'
        mask_matches = mask_pattern.search(traveler_dictionary['material'])
        if mask_matches is not None:
            sheet['B6'] = f'{date_time_obj.date() - timedelta(days=7):%m/%d/%Y}'
        mask_matches = mask_pattern.search(traveler_dictionary['notes'])
        if mask_matches is not None:
            sheet['B6'] = f'{date_time_obj.date() - timedelta(days=7):%m/%d/%Y}'

    # 4) If the pdf is absent of any of the phrases in a,b, or c and “Finish” block mentions any other kind of finish,
    # change due date to three business days before current due date.
    elif traveler_dictionary['finish'] is not None:
        sheet['B6'] = f'{date_time_obj.date() - timedelta(days=3):%m/%d/%Y}'

    # 5) If in following the rules, the resulting date is less than today’s date,
    # replace the date with the text “ASAP”
    compare_today = sheet['B6'].value
    compare_time_obj = datetime.strptime(compare_today, "%m/%d/%Y")
    if compare_time_obj.date() < datetime.today().date():
        sheet['B6'] = 'ASAP'

    # Change to folder path to grab image and save excel file.
    os.chdir(folder_path)

    # Opens image and converts to RGBA format
    im = Image.open(f'{traveler_dictionary["job_number"]}.png')
    im = im.convert('RGBA')

    # Uses numpy to convert img background to white
    data = np.array(im)  # "data" is a height x width x 4 numpy array
    red, green, blue, alpha = data.T  # Temporarily unpack the bands for readability

    # Replace black with white... (leaves alpha values alone...)
    black_areas = (red == 0) & (blue == 0) & (green == 0)
    data[..., :-1][black_areas.T] = (255, 255, 255)  # Transpose back needed

    # Closes original image, than saves new image with white background
    im2 = Image.fromarray(data)
    im.close()
    im2.save(f'{traveler_dictionary["job_number"]}.png')

    # Opens image and resizes it to fit the template.
    im = Image.open(f'{traveler_dictionary["job_number"]}.png')
    resized_im = im.resize((round(im.size[0]*0.75), round(im.size[1]*0.75)))
    resized_im.save(f'{traveler_dictionary["job_number"]}.png')

    # Anchors the image in the template excel according to absolute values to horizontally align center.
    img = openpyxl.drawing.image.Image(f'{traveler_dictionary["job_number"]}.png')
    p2e = pixels_to_EMU
    h, w = img.height, img.width
    position = XDRPoint2D(p2e(210), p2e(80))
    size = XDRPositiveSize2D(p2e(h), p2e(w))
    img.anchor = AbsoluteAnchor(pos=position, ext=size)
    sheet.add_image(img)

    # Saves the template file as the new traveler name. Customer can now open and print as PDF.
    wb.save(f'CT {traveler_dictionary["job_number"]}.xlsx')

    # Deletes the image file since we no longer need it.
    os.remove(f'{traveler_dictionary["job_number"]}.png')


def image_grab(pdf, job_number):
    """
    Grabs the preview image from the PDF for later processing into the template excel traveler.
    :pdf: the filename of the traveler to grab image from
    :job_number: the job number used to name the image file
    :return: None
    """

    # Open the PDF, only parse the first page since that has the preview image.
    pdf_obj = open(pdf, 'rb')
    input1 = PyPDF2.PdfFileReader(pdf_obj)
    page0 = input1.getPage(0)

    # Parse the PDF for iamge filetypes and saves them to the same directory.
    if '/XObject' in page0['/Resources']:
        xobject = page0['/Resources']['/XObject'].getObject()

        for obj in xobject:
            if xobject[obj]['/Subtype'] == '/Image':
                size = (xobject[obj]['/Width'], xobject[obj]['/Height'])
                data = xobject[obj].getData()
                if xobject[obj]['/ColorSpace'] == '/DeviceRGB':
                    mode = "RGB"
                else:
                    mode = "P"

                # Filter for image sizes to avoid processing the customer logo image as well.
                if '/Filter' in xobject[obj]:
                    if xobject[obj]['/Filter'] == '/FlateDecode':
                        img = Image.frombytes(mode, size, data)
                        if img.height > 100:
                            img.save(str(job_number) + ".png")
                            # img.save(str(job_number) + ' ' + obj[1:] + ".png")
                    elif xobject[obj]['/Filter'] == '/DCTDecode':
                        img = open(obj[1:] + ".jpg", "wb")
                        img.write(data)
                        img.close()
                    elif xobject[obj]['/Filter'] == '/JPXDecode':
                        img = open(obj[1:] + ".jp2", "wb")
                        img.write(data)
                        img.close()
                    elif xobject[obj]['/Filter'] == '/CCITTFaxDecode':
                        img = open(obj[1:] + ".tiff", "wb")
                        img.write(data)
                        img.close()
                else:
                    img = Image.frombytes(mode, size, data)
                    if img.height > 100:
                        img.save(str(job_number) + ".png")

    else:
        print("No image found.")
    pdf_obj.close()


def excel_to_pdf(folder_path):
    """ Goes through all files in a folder that are excel files, appends '_m' and saves as PDF format. """
    os.chdir(folder_path)
    for file in os.listdir('.'):
        if file.endswith('.xlsx'):
            logging.debug(f'Converting {file} from excel to PDF.')

            output_file = file.split('.')[0] + '_m.pdf'
            logging.debug(f'File name will be {output_file}')

            app = client.DispatchEx("Excel.Application")
            app.Interactive = False
            app.Visible = False
            workbook = app.Workbooks.Open(folder_path + '\\' + file)
            try:
                workbook.ActiveSheet.ExportAsFixedFormat(0, folder_path + '\\' + output_file)
            except Exception as e:
                print("Failed to convert in PDF format.")
                print(str(e))
            finally:
                workbook.Close(False)
                app.Quit()
                workbook = None


def process_notes(string):
    """ Processes incoming string for 'Notes' section with a bunch of specific Regex because reading PDFs is messy """
    logging.debug(f'Notes to process: \n{string}')
    new_string = ''

    thread_pattern = re.compile(r'(Threads/Tapped Holes: \n)(\d+)', re.DOTALL)
    tolerance_pattern = re.compile(r'Tolerances: \n(.*?)\n', re.DOTALL)
    roughness_pattern = re.compile(r'.*Surface Roughness: \n.*?ish (\d.*?)(?:Ra, )(speci\nc area|entire part).*?(?:Part|Notes).*', re.DOTALL)
    marking_pattern = re.compile(r'Part Markings: \n(.*?)Notes:', re.DOTALL)
    main_notes_pattern = re.compile(r'Notes:(.*)', re.DOTALL)

    thread_matches = thread_pattern.search(string)
    logging.debug(f'thread_matches: {thread_matches}')
    tolerance_matches = tolerance_pattern.search(string)
    logging.debug(f'tolerance_matches: {tolerance_matches}')
    roughness_matches = roughness_pattern.search(string)
    logging.debug(f'roughness_matches: {roughness_matches}')
    marking_matches = marking_pattern.search(string)
    logging.debug(f'marking_matches: {marking_matches}')
    main_matches = main_notes_pattern.search(string)
    logging.debug(f'main_matches: {main_matches}')

    if thread_matches is not None:
        new_string += 'Threads/Tapped Holes: ' + remove_newlines(thread_matches.group(2)) + '\n'
        logging.debug(f'current new string after thread_matches: \n{new_string}')
    if tolerance_matches is not None:
        new_string += 'Tolerances: ' + remove_newlines(tolerance_matches.group(1)) + '\n'
        logging.debug(f'current new string after tolerance_matches: \n{new_string}')
    if roughness_matches is not None:
        new_string += 'Surface Roughness: ' + remove_newlines(roughness_matches.group(1)) + ' Ra, ' + regex_specic_check(remove_newlines(roughness_matches.group(2))) + '\n'
        logging.debug(f'current new string after roughness_matches: \n{new_string}')
    if marking_matches is not None:
        new_string += 'Part Markings: ' + remove_newlines(marking_matches.group(1)) + '\n'
        logging.debug(f'current new string after marking_matches: \n{new_string}')
    if main_matches is not None:
        new_string += 'Notes: ' + main_matches.group(1)
        logging.debug(f'current new string after main_matches: \n{new_string}')

    return new_string


def remove_newlines(string):
    """ Remove new lines from a string. """
    new_string = string.replace('\n', '')
    return new_string


def low_up(string):
    """ Insert space between lowercase letters against uppercase letters. """
    new_string = re.sub(r'([a-z](?=[A-Z])|[A-Z](?=[A-Z][a-z]))', r'\1 ', string)
    return new_string


def remove_l_stroke(string):
    """ Remove L stroke character that gets parsed by the PyPDF2 for bullets. """
    new_string = re.sub(r'Ł', '', string)
    return new_string


def replace_trademark(string):
    """ Replace trademark character for apostrophe. """
    new_string = re.sub(r'™', '\'', string)
    return new_string


def replace_fluid(string):
    """ Replace fluid character for double apostrophe. """
    new_string = re.sub(r'ﬂ', '\"', string)
    return new_string


def regex_certificate_check(string):
    """ Check for 'certicate/certication' since PyPDF2 has a hard time parsing the kerning in certificate. """
    new_string = re.sub(r'(certic|Certic)', 'Certific', string)
    return new_string


def regex_specic_check(string):
    """ Check for 'specic' since PyPDF2 has a hard time parsing the kerning in specific. """
    new_string = re.sub(r'specic', 'specific', string)
    return new_string


def regex_film_check(string):
    """ Check for 'Chem-lm' since PyPDF2 has a hard time parsing the kerning in Chem-film. """
    new_string = re.sub(r'Chem-lm', 'Chem-Film', string)
    return new_string


def open_parse_pdf(filename):
    """ Opens a PDF file and extracts all of the text data from every page """

    # Opens the file in cwd, makes a reader object, and appends string to parse_string every time page is looped.
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

    return remove_l_stroke(parse_string)


def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)


def main():
    print('Press CTRL+C or close the window to exit.')
    try:
        while True:
            folder_path = input('Please paste the absolute folder path with the files you wish to process, '
                                'or press CTRL+C to exit: \n')
            logging.info(f'Getting data from the following directory: {folder_path}')
            read_document(folder_path)
            rename_unlinked_drawings(folder_path)
            excel_to_pdf(folder_path)
            print('Folder processed, please check files to make sure everything went accordingly.')
    except KeyboardInterrupt:
        sys.exit()


if __name__ == '__main__':
    main()
