# -*- coding: utf-8 -*-
"""
Created on Fri Apr 13 19:04:08 2018

@author: Alex
"""
from openpyxl import load_workbook
from price_dictionary import PriceDictionary
from price_structure import PriceStructure
import price_goodwin as prgdw
import price_write as prwr
import datetime
import re
import argparse
import globals_param


class Transform:

    def __init__(self):
        self.file_ext = ".xlsx"
        self.path_orig = '../orig/'
        self.path_trans = '../trans/'
        self.file_name = ''
        self.hotels = self.get_hotels()

    def get_hotels(self):
        h = {'OLYMPIC': 'OLYMPIC PALACE/4/KARLOVY VARY',
             'RICHMOND': 'RICHMOND/4/KARLOVY VARY'}
        return h

    def get_header(self, cell_name):

#        f_word = self.file_name.split()[0]
        f_word = re.split('[_\W]+', self.file_name)[0]
        f_word = f_word.upper()
        name = self.hotels.get(f_word)
        if name is not None:
            return name
        else:
            return cell_name

    def start(self):
        print("The System Caliber. Price transformations for Goodwin")
        parser = argparse.ArgumentParser(description='The System Caliber. Price transformations for Goodwin')
        parser.add_argument("file", type=str,
                            help="file name of original price in Excel format")
        parser.add_argument('-src', action='store', default='./')
        parser.add_argument('-dest', action='store', default='./')
        parser.add_argument('-log', action='store_true')
        args = parser.parse_args()
        self.path_orig = args.src
        self.path_trans = args.dest
        globals_param.trace_log = args.log
        if self.check_lic():
            self.start_up(args.file)
        else:
            print("lic not found!")
        
    def check_lic(self):
        date = datetime.datetime(year=2018,month=9,day=11)
        date1 = datetime.datetime.now()
        delta = date1-date
        return delta.days < 365

    def start_up(self, file_name):

        self.file_name = file_name

        wb = load_workbook(self.get_file_orig())
        sheet = wb.active

        # заполнить словарь полей
        pr_dic = PriceDictionary()
        dic = pr_dic.collect_dic(sheet)

        # заполнить структуру прайса
        price_structure = PriceStructure(dic).fill_frame(sheet)

        self.price_goodwin = prgdw.Gen_price(price_structure, dic, pr_dic.index, pr_dic.meal_dic)
        # раскрыть части прайса в dataframe
        self.price_goodwin.build_df(sheet)
        # сгенерировать таблицу для goodwin
        df = self.price_goodwin.build_periods()

        header = self.get_header(pr_dic.hotel_name)

        print(header)

        # записать в ексель
        wr = prwr.Write_price(self.get_file_trans(), df)

        wr.open_wb(header)
        wr.write(self.price_goodwin.dops)
        wr.close()

    def get_file_orig(self):
        return self.path_orig + self.file_name + self.file_ext

    def get_file_trans(self):
        index = ' G{}'.format(datetime.datetime.now().strftime('%d-%m'))
        return self.path_trans + self.file_name + index + self.file_ext


trans = Transform()
# trans.start('Olympic Palace 2018')
# trans.start('Richmond_SPO_update 28.03.18')
trans.start()

# runfile('C:/Projects/Caliber/Caliber/Caliber.py', wdir='C:/Projects/Caliber/Caliber', args='"Richmond_SPO_update 28.03.18" -dest "../trans/" -src "../orig/"')
