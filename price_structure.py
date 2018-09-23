#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Thu Apr 26 23:19:47 2018

@author: coraleks
"""

import datetime
from reporter import reporter

# заполнение предварительной структуры прайса 
#   период-(тип номера, категория, координаты встроенной таблицы прайса)
# основной метод - fill_frame
# результат - price

class PriceStructure:
    def __init__(self, dic):
        self.dic = dic
        self.sheet = None
        self.start_cell = ''
        self.price = []
        
    def fill_frame(self, cz_sheet_price):
        self.sheet = cz_sheet_price
        # global cell_start
        self.start_cell = self.sheet[self.dic['date0']['coord']].offset(1,0)
        
        while self.start_cell.value != None:
            self.price.append(self._fill_period())
            
        reporter.print('PriceStructure->fill_frame')
        reporter.print('Price Structure ')
        reporter.pprint(self.price)

        return self.price
    
    def _fill_period(self):
        # сформировать данные периода 
        # передается поле перед датой 
        def check_date(date):
            if isinstance(date, datetime.date):
                return date.strftime('%d.%m.%Y')
            return date
        
        cell_date = self.start_cell
        price_period={'period':[]}
        # поиск дат
        for i in range(0,50): #
            if cell_date.offset(i,0).value == None:
                break
            t1 = check_date(cell_date.offset(i,0).value)
            t2 = check_date(cell_date.offset(i,1).value)
            price_period['period'].append((t1, t2))
            
        # начало прайса
        rp = cell_date.offset(i,0).row
        cp = self.dic['troom']['pos']
        self.start_cell = self.sheet.cell(row = rp, column = cp)
        
        price_period['rooms'] = self._fill_category_room()   
        return price_period
    
    def _fill_category_room(self):
        
        catrooms=[]
        rp =self.start_cell.row
        date_pos = self.dic['date0']['pos'] 
        troom_pos = self.dic['troom']['pos']
        catroom_pos = self.dic['catroom']['pos']
        place_pos = self.dic['place']['pos'] # как  начало прайса
        save_troom = ''
        save_catroom = ''
        save_row = 0
        
        # сохранение текущей позиции
        def save_par():
            nonlocal save_troom, save_catroom, save_row
            save_troom = str(self.sheet.cell(row = rp, column = troom_pos).value).strip()
            save_catroom = self.sheet.cell(row = rp, column = catroom_pos).value
            save_row = rp
            return
        
        # добавление прайса по категории
        def add_price():
            nonlocal catrooms, place_pos, rp, date_pos
            cell_start = self.sheet.cell(row =  save_row, column = place_pos).coordinate
            cell_end = self.sheet.cell(row = rp-1, column = date_pos-1).coordinate
            catrooms.append({'troom':save_troom.strip(), 'catroom':save_catroom.strip(), 'price': (cell_start,cell_end)})
            return
    
        # тестирование на все пробелы
        def is_end(cell_troom):
            rp = cell_troom.row
            is_end = True
            for field in self.dic:
                col = self.dic[field]['pos']
                if col == 0:
                    continue
                if self.sheet.cell(row = rp, column = col).value != None:
                    is_end = False
                    break
            return is_end
        
        save_par() # сохранили текущую позицию
        rp+=1
        # при появлении даты считаем что период закончился
        # ессли все пусто - конец
        while (self.sheet.cell( row = rp, column =  date_pos).value == None and 
               not is_end(self.sheet.cell( row = rp, column =  troom_pos))):
            if self.sheet.cell(row = rp, column = catroom_pos).value != None: # смена категории
                add_price()
                save_par()
            rp+=1
        add_price()
        
        self.start_cell = self.sheet.cell( row = rp, column =  date_pos)
        return catrooms
        
        
        