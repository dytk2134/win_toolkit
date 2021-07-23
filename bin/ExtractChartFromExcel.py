#!/usr/bin/env python
# -*- coding: utf-8 -*-
# Contributed by Li-Mei Chiang <dytk2134 [at] gmail [dot] com> (2020)


import os
from PIL import ImageGrab
import win32com.client as win32
import win32clipboard
import logging
import sys
import argparse
from textwrap import dedent

__version__ = 'v1.0.0'

# logger
logger = logging.getLogger(__name__)
logger.setLevel(logging.INFO)
if not logger.handlers:
    lh = logging.StreamHandler()
    lh.setFormatter(logging.Formatter('%(levelname)-8s %(message)s'))
    logger.addHandler(lh)


def ExtractPictures(input_files, output_dir):
    # check if output directory exist
    if not os.path.exists(output_dir):
        try:
            os.makedirs(output_dir)
        except IOError as err:
            logger.error(err)

    chart_num = 1
    for input_file in input_files:
        # check if input file exist
        if not os.path.exists(input_file):
            logger.error('%s: No such file or directory.' % (input_file))
            sys.exit(1)
        # export chart
        excel = win32.gencache.EnsureDispatch('Excel.Application')
        workbook = excel.Workbooks.Open(os.path.abspath(input_file))
        for sheet in workbook.Worksheets:
            for shape in sheet.Shapes:
                if shape.Name.startswith('Chart'):
                    image_name = os.path.join(output_dir, 'image%s.png' % (str(chart_num).zfill(3)))
                    if os.path.exists(image_name):
                        os.remove(image_name)
                    shape.Copy()
                    image = ImageGrab.grabclipboard()
                    # if not image is None:
                    image.save(image_name, 'PNG', dpi=(600, 600))
                    # logger.info('Export %s to %s' % (shape.Name, image_name))
                    chart_num += 1
        win32clipboard.OpenClipboard()
        win32clipboard.EmptyClipboard()
        win32clipboard.CloseClipboard()
        workbook.Close(SaveChanges=False, Filename=os.path.abspath(input_file))
        excel.Quit()
    # excel = win32.Dispatch('Excel.Application')
    # workbook = excel.Workbooks.Open(os.path.abspath(input_file))
    # for sheet in workbook.Worksheets:
    #     chart 


def main():
    parser = argparse.ArgumentParser(formatter_class=argparse.RawDescriptionHelpFormatter, description=dedent("""\
    This script is for extracting pictures from Excel or Word.
    Quick start:
    %(prog)s -i example.xlsx -d ./images

    """))
    # argument
    parser.add_argument('-i', '--input_files', nargs='+', help='Location of input file')
    parser.add_argument('-d', '--directory', type=str, help='Location to save extracted images.')
    parser.add_argument('-v', '--version', action='version', version='%(prog)s ' + __version__)

    args = parser.parse_args()

    ExtractPictures(args.input_files, args.directory)


if __name__ == '__main__':
    main()