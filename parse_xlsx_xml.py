# -*- coding: utf-8 -*-
"""
    Xlsx xml-parser for Reporting Services.
    Converts text to formulae, eg. '=SUM(A1:A10)'
    ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    Notice: Only Reporting Services 2012 (or higher) is supporting export reports to
            xlsx-format.
"""

import os
import shutil
import time
from zipfile import ZipFile, ZIP_DEFLATED
from lxml import etree

from xlsx_rc_convertor import convert_rc_formula


class RecursiveFileIterator:
    def __init__(self, *root_dirs):
        self.dir_queue = list(root_dirs)
        self.include_dirs = None
        self.file_queue = []

    def __getitem__(self, index):
        while not len(self.file_queue):
            self.next_dir()
        result = self.file_queue[0]
        del self.file_queue[0]
        return result

    def next_dir(self):
        dir = self.dir_queue[0]   # fails with IndexError, which is fine
        # for iterator interface
        del self.dir_queue[0]
        list = os.listdir(dir)
        join = os.path.join
        isdir = os.path.isdir

        for basename in list:
            full_path = join(dir, basename)
            if isdir(full_path):
                self.dir_queue.append(full_path)
                if self.include_dirs:
                    self.file_queue.append(full_path)
            else:
                self.file_queue.append(full_path)


class ParseXlsx:
    """ Parse xlsx file and replace formulas strings to formulas format """

    def __init__(self, file_name, task_id=0, show_log=False, run=False):
        """ Init start parameters """
        self.file_name = file_name
        self.task_id = task_id
        self.main_temp_dir = 'temp'
        self.show_log = show_log
        self.shared_strings = []
        if run:
            self.main()

    def main(self):
        """ Read xlsx file, extract files from it and parse each sheet """
        if not os.path.exists(self.file_name):
            print('Source file not found. Exit.')
        else:
            if not os.path.isdir(self.main_temp_dir):
                self.print_log('Creating temp directory')
                os.mkdir(os.path.join(os.getcwd(), self.main_temp_dir))
            os.chdir(self.main_temp_dir)
            # Create temp dir
            temp_dir = str(self.task_id) + str(time.time())
            os.mkdir(os.path.join(os.getcwd(), temp_dir))
            os.chdir(temp_dir)
            # Extract xlsx and process it
            zip_file_name = os.path.join("../" * 2, self.file_name)
            with ZipFile(zip_file_name, 'a', ZIP_DEFLATED) as report_zip:
                report_zip.extractall(os.getcwd())
                # Extract all strings from sharedStrings.xml
                shared_string_xml_object = etree.parse('xl/sharedStrings.xml')
                for t_tags in shared_string_xml_object.getroot().getchildren():
                    for t_tag in t_tags.getchildren():
                        self.shared_strings.append(t_tag.text)
                    # Process each sheet
                for sheet_file_name in report_zip.namelist():
                    if 'xl/worksheets/sheet' in sheet_file_name:
                        self.parse_sheet(sheet_file_name)

            self.print_log('Deleting source file')
            os.remove(zip_file_name)
            with ZipFile(zip_file_name, "w") as cur_file:
                for name in RecursiveFileIterator('.'):
                    self.print_log('Writing to new Excel file. File -> {0}'.format(name))
                    if os.path.isfile(name):
                        cur_file.write(name, name, ZIP_DEFLATED)

            os.chdir('..')
            self.print_log('Removing temp files')
            shutil.rmtree(os.path.join(os.getcwd(), temp_dir))
            self.print_log('Done')

    def parse_sheet(self, sheet_file_name):
        """ Parse sheet and  replace formulas strings to formulas format """
        sheet_xml_object = etree.parse(sheet_file_name)
        c_tags = sheet_xml_object.getroot().xpath("//*[local-name()='sheetData']/*[local-name()='row']/*[local-name()='c'][@t='s']")

        for c_tag in c_tags:
            v_tag = c_tag.xpath("*[local-name()='v']")

            if self.shared_strings[int(v_tag[0].text) + 1]:
                cur_shared_string = self.shared_strings[int(v_tag[0].text) + 1]
                if cur_shared_string[0] == '=':
                    self.print_log(
                        'Find formula -> {0} in row {1}'.format(cur_shared_string, c_tag.get('r')))
                    right_formula = convert_rc_formula(cur_shared_string, c_tag.get('r'))
                    c_tag.remove(v_tag[0])
                    c_tag.append(etree.Element("f"))
                    f_tag = c_tag.xpath("*[local-name()='f']")
                    f_tag[0].text = right_formula
                    del c_tag.attrib["t"]

        file_handler = open(sheet_file_name, "w")
        file_handler.writelines(etree.tostring(sheet_xml_object, pretty_print=True))
        file_handler.close()

    def print_log(self, message):
        if self.show_log:
            print(message)


if __name__ == '__main__':
    ParseXlsx('formula_test.xlsx', show_log=True, run=True)
