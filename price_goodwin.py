# -*- coding: utf-8 -*-
"""
Created on Sat Apr 28 14:06:09 2018

@author: Alex
"""

import pandas as pd
from openpyxl.utils import column_index_from_string
from openpyxl.utils.dataframe import dataframe_to_rows
import itertools
from math import nan
import collections
from reporter import reporter

code_na = 'N/A'


# генерация элементов прайса dataframe в таблицу
# build_df - формирование Dataframe для областей прайса, 
# build_periods - формирование прайса в формате goodwin, результат в df

class Gen_price:

    def __init__(self, price_structure, dic, index_dic, meal_dic):  # структура которая формализует строки прайса

        self.price = price_structure
        self.dic = dic
        self.index_dic = index_dic
        self.meal_dic = meal_dic
        self.meals = []
        self.meals_label = []
        self.meals_count = 0
        self._get_meal_list2()

        self.data_test = {}
        # минимальный возраст ребенка по отдельным категориям, 
        self.df = pd.DataFrame({})
        # весь прайс для анализа возрастов детей
        self.full_df = pd.DataFrame({})
        # справочник для допов v2
        self.DopPrice=collections.namedtuple('DopPrice', 'age0 age1 place name desc')
        self.dops=[]

    def build_df(self, sheet):  # формирование DataFrame объектов для областей таблиц прайса
        for periods in self.price:
            for catroom in periods['rooms']:
                # раскрытие области прайса в dataframe
                # df записывается обратно в структуру price
                catroom['df'] = self._buid_line_df(catroom, sheet)
                # побочно: формирование полной таблицы для анализа
                self.full_df = self.full_df.append(catroom['df'], ignore_index=True)
        # определение минимального возраста
        self.age0_min = self.full_df['age0'].min()
        # создание справочника по допам
        self._build_dops()
        
        return
    
    # создание справочника по допам
    # обработка self.full_df
    def _build_dops(self):
        # определение уникальных пар по возрастам
        age_group = self.full_df.groupby(['age0', 'age1']).mean().reset_index()

        # формирование справочника по допам
        # на каждый возраст два варианта
        for age_row in dataframe_to_rows(age_group, index=False, header=False):
            age0 = int(age_row[0])
            age1 = int(age_row[1]) +1 # максимальный возраст увеличивается на 1 
            name1 = 'dbl_child_main_{}_{}'.format(age0, age1)
            desc1 = 'DBL (1 adult main + 1 child main) ({}-{})'.format(age0, age1)
            name2 = 'dbl_child_extra_{}_{}'.format(age0, age1)
            desc2 = 'DBL (2 adult main + 1 child exb) ({}-{})'.format(age0, age1)

            self.dops.append(self.DopPrice(age0, age1, 1, name1, desc1))
            self.dops.append(self.DopPrice(age0, age1, 2, name2, desc2))
        return
        

    # раскрытие области прайса в dataframe
    def _buid_line_df(self, catroom, sheet):
        (t1, t2) = catroom['price']
        c0 = sheet[t1]
        c1 = sheet[t2]
        gen_cells = sheet.iter_cols(min_row=c0.row,
                                    min_col=column_index_from_string(c0.column),
                                    max_row=c1.row,
                                    max_col=column_index_from_string(c1.column))
        price_series = {}
        for cols in gen_cells:
            l = []
            for cell in cols:
                l.append(cell.value)
            field = self.index_dic[column_index_from_string(cell.column)]
            price_series[field] = pd.Series(l)
            df = pd.DataFrame(price_series)
        return df

    # формироваие области прайса на категорию номера
    def build_periods(self):

        # секции прайса по периодам
        for section_periods in self.price:
            # список периодов    
            periods = section_periods['period'][0:1]  # one periods [0:1]
            # прайс по типу номеров
            ddf = pd.DataFrame({})
            for catroom_price in section_periods['rooms'][:]:
                #periods можно будет убрать
                part_df=self._build_line(catroom_price['catroom'],
                                 catroom_price['troom'],
                                 catroom_price['df'])
                ddf = ddf.append(part_df)
                
            reporter.print('Gen_price->build_periods')
            reporter.print('Dataframe for period ')
            reporter.pprint(periods)
            reporter.print('Dataframe ')
            reporter.pprint(ddf)
            # размножить по периодам
            self._multi_period(periods, ddf)
 
        # вычислить мин и создать серию
        self._add_sr_age(self.age0_min)
        # чистка допов (вертикал)
        self._check_none_dop()
        # замена незначихся
        self.df = self.df.replace(nan, 'N/A')
        # чистка пустых строк (горизонт)
        self.df = self.df.query('not(sgl == "N/A" and dbl == "N/A" and trp == "N/A" and dsu == "N/A")')

        return self.df
    
    # Чистка серий доп размещений которые все null
    def _check_none_dop(self):
        dop_checked = []
        for dop in self.dops:
            # все нули Nan
            if self.df[dop.name].isnull().values.all():
                # удаление столбцов
                self.df.drop([dop.name], axis=1)
            else:
                dop_checked.append(dop)
        self.dops = dop_checked

    # сборка всех питаний которые дейтсвуют для прайса
    def _get_meal_list2(self):
        for meal in self.meal_dic:
            if meal['pos'] > 0:
                self.meals.append(meal['pat'])
                self.meals_label.append(meal['pat'])
        self.meals_count = len(self.meals)
        return

    # создание части прайса по прайсу категории номера (data)
    # прайс состоит из мини серий
    # впоследствии части прийса соединяются
    def _build_line(self, catroom, troom, data):
        # ['catroom', 'meal', 'period', '1', '2', '3', 'age', 'dsu']
        name_cat = self._name_catroom(troom, catroom)

        # серия категорий номеров
        sr_catroom = pd.Series(list(itertools.repeat(name_cat, self.meals_count)))
        # серия типов питания
        sr_meal = pd.Series(self.meals_label)
        
        data['taccom'] = data['taccom'].str.strip()
        data['yaccom'] = data['yaccom'].str.strip()

        sr_sgl = Accom(troom, self.meals, data).build_series2(['SGL'], 
                      'place == 1 and taccom == "BASE" and yaccom == "ADULT"')
        # одно место в dbl
        sr_dbl1 = Accom(troom, self.meals, data).build_series2(['DBL', 'SUITE', 'TRIPLE', 'APARTMENT'],
                       'place == 2 and taccom == "BASE" and yaccom == "ADULT"')
        sr_dbl2 = sr_dbl1 * 2

        # одно место в four
        sr_four1 = Accom(troom, self.meals, data).build_series2(['SUITE'],
                       'place == 4 and taccom == "BASE" and yaccom == "ADULT"')
        sr_four4 = sr_four1 * 4

        sr_extra = Extra_accom(troom, self.meals, data).build_series2(['DBL', 'SUITE', 'TRIPLE', 'APARTMENT'],
                        'taccom == "EXTRA" and yaccom == "ADULT"')
        sr_triple = sr_dbl2 + sr_extra
        sr_dsu = Accom(troom, self.meals, data).build_series2(['DBL', 'SUITE', 'TRIPLE', 'APARTMENT'],
                         'place == 1 and taccom == "BASE" and yaccom == "ADULT"' )

        df_main = pd.DataFrame({'catroom': sr_catroom,
                            'meal': sr_meal,
                            'sgl': sr_sgl,
                            'dbl': sr_dbl2,
                            'trp': sr_triple,
                            'dsu': sr_dsu,
                            'four': sr_four4,
                            })
        df_dop = self._child_dop2(sr_dbl1, troom, self.meals, data)
        ddf = pd.concat([df_main, df_dop], axis=1)

        return ddf
    
    # создание общего блока прайса и размножение по периодам 
    def _multi_period(self, periods, ddf):
        for (p1, p2) in periods:
            interval = '{}-{}'.format(p1, p2)  # даты не анализируются
            (num_row, _) = ddf.shape
            sr_period = pd.Series(list(itertools.repeat(interval, num_row)))
            df_period = pd.DataFrame({'period': sr_period})
            # переиндексация
            ddf.reset_index(inplace=True, drop=True)
            # склейка по вертикали
            ddf = pd.concat([ddf, df_period], axis=1)
            # склейка по горизонтали
            self.df = self.df.append(ddf)
        return
        
    # переработка под dops
    def _child_dop2(self, series_dbl_1, troom, meals, data):
        data = self._update_taccom(data)
        df = pd.DataFrame({})
        for dop in self.dops:
            place = 'BASE' if dop.place == 1 else 'EXTRA'
            criteria = 'yaccom == "CHILD" '
            criteria += 'and age0 == "{}" '.format(dop.age0)
            criteria += 'and taccom == "{}" '.format(place)
            sr_child = Accom(troom, meals, data).build_series2(['DBL', 'SUITE'], criteria)
            
            # при нулях расчитываются нули, а что должно?
            sr_adult_child = (series_dbl_1 * dop.place) + sr_child
            ddf = pd.DataFrame({dop.name: sr_adult_child})
            df = pd.concat([df, ddf], axis=1)
        return df
    
    # дополнение полей для размещения возразтов 
    #  если не указано берется из предыдущего поля
    def _update_taccom(self, data):
        index = data[(data.taccom.isnull()) & (data.yaccom == 'CHILD') ].index
        for i in index:
            #taccom = data['taccom'][i-1]
            # data['taccom'][i] = taccom
            data.loc[i, 'taccom'] = data.loc[i - 1, 'taccom']
        return data
    
    def _add_sr_age(self, age_min):
        if pd.isnull(age_min):
            age_min = 3
        v = '{}-{}'.format(int(0), int(age_min))
        (num_row, _) = self.df.shape
        self.df['age'] = pd.Series(list(itertools.repeat(v, num_row)))

    def _name_catroom(self, troom, catroom):
        cr = 'STANDARD' if catroom == '-' else catroom
        tr = 'SUITE ' + cr if troom == 'SUITE' else cr
        return tr


# класс формирует серию для типа размещения (Sgl, Dbl, Dsu, Extra)
class Accom:

    def __init__(self, troom, meals, data):
        self.troom = troom
        self.meals = meals
        self.meals_count = len(meals)
        self.data = data
        self.accom = []
        self.criteria = ''
        self.empty_sr = pd.Series([])
        self.na_sr = pd.Series(list(itertools.repeat(nan, self.meals_count)))

    def build_series(self):
        if self.troom not in self.accom:
            return self.empty_series()
        df = self.data.query(self.criteria)
        (rows, _) = df.shape
        cols = []
        if rows != 1:  # найдена одна строка, строгое соответствие
            return self.empty_series()
        for meal in self.meals:
            cols.append(self.calc_rate(df, meal))
        return pd.Series(cols)

    def build_series2(self, accom, criteria):
        self.accom = accom
        self.criteria = criteria
        return self.build_series()

    def calc_rate(self, df, meal):
        ret = df[meal].values[0]
        #        return code_na if ret == '-' else ret
        return None if ret == '-' else ret

    def empty_series(self):
        return self.na_sr

class Extra_accom(Accom):

    def __init__(self, troom, meals, data):
        Accom.__init__(self, troom, meals, data)

    def empty_series(self):
        return self.empty_sr
