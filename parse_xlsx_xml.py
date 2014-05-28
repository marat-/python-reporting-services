# -*- coding: utf-8 -*-
"""
    Xlsx xml-parser for Reporting Services.
    Converts text to formulae, eg. '=SUM(A1:A10)'
    ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    Notice: Only Reporting Services 2012 (or higher) is supporting export reports to
            xlsx-format.
"""
from __future__ import unicode_literals

import sys
import os
import shutil
import time
import re
from zipfile import ZipFile, ZIP_DEFLATED
from lxml import etree
from copy import deepcopy

from xlsx_rc_convertor import convert_rc_formula, get_cell_format, col2str


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
        dir = self.dir_queue[0]  # fails with IndexError, which is fine
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

    def __init__(self, file_name, task_id=0, show_log=False, run=False, **kwargs):
        """ Init start parameters """
        self.file_name = file_name
        self.task_id = task_id
        self.main_temp_dir = 'temp'
        self.show_log = show_log
        self.shared_strings = []
        self.style_list = None

        # Print view params
        self.print_view = kwargs.get('print_view')
        self.landscape = 'landscape' if kwargs.get('landscape') else 'portrait'
        self.fit_to_width = str(int(kwargs.get('fit_to_width', 0)))
        self.fit_to_height = str(int(kwargs.get('fit_to_height', 0)))
        self.fix_area = kwargs.get('fix_area', [])

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
                # Check if file generated with sharedString or with inlineStr
                if os.path.isfile('xl/sharedStrings.xml'):
                    self.print_log('Found sharedStrings')
                    # Extract all strings from sharedStrings.xml
                    shared_string_xml_object = etree.parse('xl/sharedStrings.xml')
                    si_tags = shared_string_xml_object.getroot().xpath("//*[local-name()='sst']/*[local-name()='si']")
                    for si_tag in si_tags:
                        t_tag = si_tag.xpath("*[local-name()='t']")
                        if not t_tag:
                            self.shared_strings.append(None)
                        else:
                            self.shared_strings.append(t_tag[0].text)
                else:
                    self.print_log('sharedStrings not found')
                # Process each sheet
                for sheet_file_name in report_zip.namelist():
                    if 'xl/worksheets/sheet' in sheet_file_name:
                        self.parse_sheet(sheet_file_name)

            self.print_log('Deleting source file')
            os.stat(zip_file_name)
            os.remove(zip_file_name)
            with ZipFile(zip_file_name, "w") as cur_file:
                for name in RecursiveFileIterator('.'):
                    self.print_log('Writing to new Excel file. File -> {0}'.format(name))
                    if os.path.isfile(name):
                        cur_file.write(name, name, ZIP_DEFLATED)

            os.chdir('..')
            self.print_log('Removing temp files')
            shutil.rmtree(os.path.join(os.getcwd(), temp_dir))
            # Return to script's work directory
            os.chdir(sys.path[0])
            self.print_log('Done')

    def parse_sheet(self, sheet_file_name):
        """ Parse sheet and  replace formulas strings to formulas format """
        styles_file = 'xl/styles.xml'
        self.style_list = etree.parse(styles_file)
        sheet_xml_object = etree.parse(sheet_file_name)
        # Removing NaN values
        v_nan_tags = sheet_xml_object.getroot().xpath(
            "//*[local-name()='c']/*[local-name()='v' and text()='NaN']"
        )
        for v_nan_tag in v_nan_tags:
            c_nan_tag = v_nan_tag.xpath("ancestor::*[local-name()='c']")
            self.print_log("Found NaN value in cell {0}".format(c_nan_tag[0].get("r")))
            v_nan_tag.text = "0"

        # If not found sharedStrings, then looking for inlineStr c tags
        if not len(self.shared_strings):
            c_tags = sheet_xml_object.getroot().xpath(
                "//*[local-name()='sheetData']/*[local-name()='row']/*[local-name()='c'][@t='inlineStr']"
            )
            for c_tag in c_tags:
                is_tag = c_tag.xpath("*[local-name()='is']")
                t_tag = c_tag.xpath("*[local-name()='is']/*[local-name()='t']")
                if len(t_tag):
                    cur_inline_string = t_tag[0].text
                    if cur_inline_string and cur_inline_string[0] == '=':
                        self.print_log(
                            'Found formula -> {0} in row {1}'.format(cur_inline_string, c_tag.get('r'))
                        )
                        right_formula = convert_rc_formula(cur_inline_string[1:], c_tag.get('r'))
                        if right_formula:
                            c_tag.remove(is_tag[0])
                            # Generate formula
                            self.gen_formula_tag(c_tag, right_formula)
                        # Set format to formula's cell
                        if '@' in cur_inline_string[1:]:
                            c_tag.attrib['s'] = self.set_format(c_tag.get('s'), get_cell_format(cur_inline_string[1:]))
        else:
            c_tags = sheet_xml_object.getroot().xpath(
                "//*[local-name()='sheetData']/*[local-name()='row']/*[local-name()='c'][@t='s']"
            )
            for c_tag in c_tags:
                v_tag = c_tag.xpath("*[local-name()='v']")
                if self.shared_strings[int(v_tag[0].text)]:
                    cur_shared_string = self.shared_strings[int(v_tag[0].text)]
                    if cur_shared_string[0] == '=':
                        self.print_log(
                            'Found formula -> {0} in row {1}'.format(cur_shared_string, c_tag.get('r'))
                        )
                        right_formula = convert_rc_formula(cur_shared_string[1:], c_tag.get('r'))
                        if right_formula:
                            c_tag.remove(v_tag[0])
                            # Generate formula
                            self.gen_formula_tag(c_tag, right_formula)
                        # Set format to formula's cell
                        if '@' in cur_shared_string[1:]:
                            c_tag.attrib['s'] = self.set_format(c_tag.get('s'), get_cell_format(cur_shared_string[1:]))

        # Save changes in styles.xml
        self.save_xml_to_file(self.style_list, styles_file)

        # Set sheet styles
        sh_num = int(re.compile(r'\d+').findall(sheet_file_name)[-1])
        if self.print_view:
            sheet_xml_object = self.set_print_view(sheet_xml_object)
        if sh_num <= len(self.fix_area):
            sheet_xml_object = self.set_fixed_area(sheet_xml_object, int(self.fix_area[sh_num-1][0]), int(self.fix_area[sh_num-1][1]))

        # Save changes in sheetN.xml
        self.save_xml_to_file(sheet_xml_object, sheet_file_name)

    @staticmethod
    def gen_formula_tag(c_tag, right_formula):
        """ Generate new formula tag """
        c_tag.append(etree.Element("f"))
        f_tag = c_tag.xpath("*[local-name()='f']")
        f_tag[0].text = right_formula
        del c_tag.attrib["t"]

    def print_log(self, message):
        """ Show log messages during work """
        if self.show_log:
            print(message)

    def set_format(self, style_id, new_format):
        """ Set formula's cell format """
        new_format = new_format.replace("'", '"')

        # Find current common format
        cell_xfs = self.style_list.getroot().xpath(
            "//*[local-name()='cellXfs']"
        )[0]
        current_xf = deepcopy(cell_xfs.xpath("*[local-name()='xf']")[int(style_id)])
        # Append copied common format
        cell_xfs.append(current_xf)
        # Save last item's id as new style_id
        style_id = cell_xfs.attrib['count']
        # Increase cellXfs' count
        cell_xfs.attrib['count'] = str(int(cell_xfs.get('count')) + 1)

        # Get new common format
        current_xf = cell_xfs.xpath("*[local-name()='xf']")[-1]

        # Edit numFmts block
        num_fmts = self.style_list.getroot().xpath(
            "//*[local-name()='numFmts'][@count]"
        )[0]

        # Check on existing current style
        exists_fmt = num_fmts.xpath(
            """*[local-name()='numFmt'][@formatCode='[$-010419]{0}']""".format(new_format)
        )

        if not exists_fmt:
            # Add new numFmt
            num_fmts.append(etree.Element('numFmt'))
            new_item = num_fmts.xpath("*[local-name()='numFmt']")[-1]
            new_item.attrib['numFmtId'] = str(style_id)
            new_item.attrib['formatCode'] = """[$-010419]{0}""".format(new_format)

            # Increase numFmts count
            num_fmts.attrib['count'] = str(int(num_fmts.get('count')) + 1)

        # Set format number's id to new common style
        current_xf.attrib["numFmtId"] = str(
            exists_fmt[0].get('numFmtId') if exists_fmt else style_id
        )

        return style_id

    def set_print_view(self, sheet_object):
        """ Set pageSetup-tag """
        # Set fixToPage property to True
        sheet_pr = sheet_object.getroot().xpath("//*[local-name()='sheetPr']")
        if not len(sheet_pr):
            sheet_object.getroot().xpath("//")[0].insert(0, etree.Element('sheetPr'))
            sheet_pr = sheet_object.getroot().xpath("//*[local-name()='sheetPr']")[0]
        else:
            sheet_pr = sheet_pr[0]
        sheet_pr.append(etree.Element('pageSetUpPr', {'fitToPage': '1'}))

        # Set orientation to landscape and fit to width and height to True
        page_setup = sheet_object.getroot().xpath("//*[local-name()='pageSetup']")[0]
        page_setup.attrib['orientation'] = self.landscape
        page_setup.attrib['fitToWidth'] = self.fit_to_width
        page_setup.attrib['fitToHeight'] = self.fit_to_height

        return sheet_object

    @staticmethod
    def set_fixed_area(sheet_object, col=0, row=0):
        """ Set fixed area to sheet """
        # Get sheetViews tag
        sheet_views = sheet_object.getroot().xpath("//*[local-name()='sheetViews']")
        if not len(sheet_views):
            sheet_object.getroot().xpath("//")[0].insert(0, etree.Element('sheetViews'))
            sheet_views = sheet_object.getroot().xpath("//*[local-name()='sheetViews']")[0]
        else:
            sheet_views = sheet_views[0]

        # Get sheetView tag
        cur_sheet_view = sheet_views.xpath("*[local-name()='sheetView']")
        if not len(cur_sheet_view):
            sheet_views.insert(0, etree.Element('sheetView'))
            cur_sheet_view = sheet_views.xpath("*[local-name()='sheetView']")[0]
        else:
            cur_sheet_view = cur_sheet_view[0]

        # Add new pane to fix current area
        cur_sheet_view.append(etree.Element('pane', {
            'xSplit': str(col),
            'ySplit': str(row),
            'topLeftCell': "{col}{row}".format(**dict(
                col=col2str(col+1, run=1),
                row=row+1,
            )),
            'activePane': "bottomRight",
            'state': "frozen",
        }))

        return sheet_object

    @staticmethod
    def save_xml_to_file(xml_object, file_name):
        """ Save edited XML-object to source-file """
        file_handler = open(file_name, "w")
        file_handler.writelines(etree.tostring(xml_object, pretty_print=True))
        file_handler.close()


if __name__ == '__main__':
    file_name = 'KeyIndicatorsTT.xlsx'
    ParseXlsx(file_name, show_log=True, run=True)
    os.stat(file_name)