# -*- coding: utf-8 -*-
"""
Created on Thu Apr 26 19:15:11 2018

@author: Alex
"""

from openpyxl.utils import column_index_from_string
import Utils
from reporter import reporter

# поиск полей и определение их позиций 
# (горизонтальная структура)
# основной метод collect_dic
class PriceDictionary:
        def __init__(self):
            # область начального поиска заголовка
            self.searh_area = ('a1', 'r20')
            # словарь полей
            self.dic = {
                   'troom':{'pat':'ROOM TYPE', 'pos':0},
                   'catroom':{'pat':'ROOM CATEGORY', 'pos':0},
                   'place':{'pat':'PLACES','pos':0},
                   'extra':{'pat':'EXTRA','pos':0},
                   'taccom':{'pat':'TYPE','pos':0},
                   'yaccom':{'pat':'NAME','pos':0},
                   'age0':{'pat':'FROM','pos':0},
                   'age1':{'pat':'TO','pos':0},
                   # мнимое поле после раскрытия кодов pos=0
                   'meal':{'pat':'MEAL', 'pos':0},
                   # мнимое поле после раскрытия кодов pos=0
                   'period':{'pat':'PERIOD','pos':0,'pos1':0},
                   'date0':{'pat':'','pos':0},
                   'date1':{'pat':'','pos':0},
                   }       
            # словарь соответсвия позиции и поля
            self.index = {}
            # текущая позиция
            self.start_cell = ""
            self.sheet = None
            self.meal_dic = []
            self.hotel_name = ''
        
        def collect_dic(self, cz_sheet_price):
            self.sheet = cz_sheet_price
            if self._find_header_price():
                # собрать позиции полей в заголовке
                self.hotel_name = self._find_hotel_name()
                # для дат и питания отметить только общий заголовок
                self._collect_header()
                #  собрать даты по общему заголовку
                self._collect_dates()
                # собрать коды питания по общему заголовку 
                self._collect_meals()
                # построить обратный индекс - от позиции к полю
                self._index_dic()
                reporter.print('PriceDictionary->collect_dic')
                reporter.print('collect dictionary')
                reporter.pprint(self.dic)
                reporter.print('PriceDictionary->collect_dic')
                reporter.print('collect index')
                reporter.pprint(self.index)
            return self.dic
        
        def _find_header_price(self):
            (p1 , p2) = self.searh_area
            find, self.start_cell = Utils.pos_begin(self.sheet[p1:p2], self.dic['troom']['pat'])
            return find

        def _find_hotel_name(self):
            (p1 , p2) = self.searh_area
            (_, cell_name) = Utils.pos_begin(self.sheet[p1:p2])
            return cell_name.value
        
        def _collect_header(self):
            c = self.start_cell
            gen_cells = self.sheet.iter_rows(min_row=c.row, min_col=column_index_from_string(c.column), 
                                        max_row=c.row+2, max_col=20)
            cells=[]
            # сборка строк из генератора
            for r in gen_cells:
                cells.append(r) 
            for field in self.dic:
                if 'pat' in self.dic[field]:
                    found, cell = Utils.pos_begin(cells, self.dic[field]['pat'])
                    if found:
                        self.dic[field]['pos'] = column_index_from_string(cell.column)
                        self.dic[field]['coord'] = cell.coordinate
            reporter.print('PriceDictionary->_collect_header')
            reporter.print('collect header 1 stage')
            reporter.pprint(self.dic)
            return

        def _collect_dates(self):
            # вычисление дат периода так как поля дат имеют повторямые названия from to
            coord = self.dic['period']['coord']
            cell = self.sheet[coord].offset(1,0)
            self.dic['date0']['pos'] = column_index_from_string(cell.column)
            self.dic['date0']['coord'] = cell.coordinate
            cell = self.sheet[coord].offset(1,1)
            self.dic['date1']['pos'] = column_index_from_string(cell.column)
            self.dic['date1']['coord'] = cell.coordinate 
            # period мнимое поле, больше нам не нужно
            self.dic['period']['pos'] = 0
            return
       
        def _index_dic(self):
            for field in self.dic:
                n = self.dic[field]['pos']
                if n > 0: 
                    self.index[n]=field
            return      
        
        def _collect_meals(self):
            coord = self.dic['meal']['coord']
            cell = self.sheet[coord].offset(1,0)
            coord = self.dic['period']['coord']
            cell_end = self.sheet[coord].offset(1,-1)
            start_col=column_index_from_string(cell.column)
            end_col=column_index_from_string(cell_end.column)
            row=cell.row
            
            cols = self.sheet.iter_cols(min_col=start_col, max_col=end_col, 
                                         min_row=row, max_row=row)
            meal_list=[]
            for col in cols:
                for cell in col:
                    value = cell.value
                    pos = column_index_from_string(cell.column)
                    coord = cell.coordinate
                    meal_list.append({'pat': value, 'pos': pos, 'coord': coord})
                    self.dic[value] = {'pat': value, 'pos': pos, 'coord': coord}

#            self.dic['meal_list']['list'] = meal_list
            self.meal_dic = meal_list       
            # meal мнимое поле, больше нам не нужно
            self.dic['meal']['pos'] = 0
            return
            
                
        
