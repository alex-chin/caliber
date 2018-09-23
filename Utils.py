# -*- coding: utf-8 -*-
"""
Created on Thu Apr 26 19:23:00 2018

@author: Alex
"""

# найти строку в ячейках
def pos_begin(cells, str=''):
    for rs in cells:
        for c1 in rs:
            if (c1.value == str) or (str == '' and c1.value != None):
                return True, c1
    return False, None