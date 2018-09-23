# -*- coding: utf-8 -*-
"""
Created on Sun Jun 10 18:46:14 2018

@author: Alex
"""
import pprint
import globals_param


class reporter:
    def print(message):
        if globals_param.trace_log:
            print(message)
        
    def pprint(structure):
        if globals_param.trace_log:
            pp = pprint.PrettyPrinter(indent=4)
            pp.pprint(structure)
        