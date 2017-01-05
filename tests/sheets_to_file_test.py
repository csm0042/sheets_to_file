#!/usr/bin/python3
""" sheets_to_file_test.py:   
""" 

# Import Required Libraries (Standard, Third Party, Local) ****************************************
import logging
import file_logger
import unittest
import os
import sys
sys.path.insert(0, os.path.abspath('..'))
import sheets_to_file


# Authorship Info *********************************************************************************
__author__ = "Christopher Maue"
__copyright__ = "Copyright 2016, The RPi-Home Project"
__credits__ = ["Christopher Maue"]
__license__ = "GPL"
__version__ = "1.0.0"
__maintainer__ = "Christopher Maue"
__email__ = "csmaue@gmail.com"
__status__ = "Development"


# Define test class *******************************************************************************
class TestGoogSheetToFile(unittest.TestCase):
    def setUp(self):
        self.debug_file, self.info_file = file_logger.setup_log_files(__file__)
        self.logger = file_logger.setup_log_handlers(__file__, self.debug_file, self.info_file)
        self.sheet_object = sheets_to_file.GoogleSheetsToFile(logger=self.logger)


    def test_read_data(self):    
        self.results = self.sheet_object.read_data(
            sheet_id='1LJpDC0wMv3eXQtJvHNav_Yty4PQcylthOxXig3_Bwu8',
            sheet_range="fylt1!A3:L")
        self.assertEqual(len(self.results), 5)


    def test_write_file(self):    
        self.results = self.sheet_object.read_data(
            sheet_id='1LJpDC0wMv3eXQtJvHNav_Yty4PQcylthOxXig3_Bwu8',
            sheet_range="fylt1!A3:L")
        self.output_file = "c:\python_logs\schedule.txt"
        self.sheet_object.write_to_file(lines=self.results, output_file=self.output_file)
        self.assertEqual(len(self.results), 5)



if __name__ == "__main__":
    logging.basicConfig(stream=sys.stdout)
    logger = logging.getLogger(__name__)
    logger.level = logging.DEBUG
    logger.debug("\n\nStarting log\n")
    unittest.main() 