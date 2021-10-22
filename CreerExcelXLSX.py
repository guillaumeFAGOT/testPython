#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Fri Oct 22 09:12:16 2021

@author: guillaume
"""
from openpyxl import load_workbook

def CreerClasseur():
    from openpyxl import Workbook
    wb = Workbook()
    ws = wb.active
    ws1 = wb.create_sheet("Mysheet")
    wb.save('balances.xlsx')
    
def AllerChercherDonneesClasseur():
#Aller chercher les valeurs de la feuille1 "Sheet"
    wb = load_workbook('balances.xlsx')
    ws = wb["Sheet"]
    for row in ws.iter_rows(min_row=1, max_col=6, max_row=6):
        for cell in row:
            if cell.value is not None:
                print (cell.value)
            
                
AllerChercherDonneesClasseur()
    
    
    
