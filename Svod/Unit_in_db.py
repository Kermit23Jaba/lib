from array import array
from multiprocessing.sharedctypes import RawArray
from time import sleep
from typing import final
# from openpyxl import load_workbook
import openpyxl
import os
import glob

def dell_slo(word):
    a = word.split()
    for i in a:
        if '№' in list(i):
            c = i.replace(',', '')
            return c

def dell_slo_o(a):
    b = a.split()
    c = list(b[0])
    d = c[:-1]
    g =''.join(d)
    return g
#####################################
# import sqlite3

# with sqlite3.connect('C:/Users/user/Desktop/svod_e.db') as con:

#     cur = con.cursor()


#######################################
import sqlite3
 
con = sqlite3.connect('C:/Users/user/Desktop/svod_e.db')
 
def sql_insert(con, entities):
 
    cursorObj = con.cursor()
    
    cursorObj.execute('INSERT INTO svod_avk_elements VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)', entities)
    
    con.commit()
 





############################################
lines = os.listdir('//ALYANS/ovk/Arhiv/')
linsize = len(lines)


for name_line in lines:
    all_papki_prover = os.listdir(f'//ALYANS/ovk/Arhiv/{name_line}/')

    for i in all_papki_prover:

        if 'Elements' in all_papki_prover:

            all_papki = os.listdir(f'//ALYANS/ovk/Arhiv/{name_line}/Elements/')

            for name_papka in all_papki:
                all_file_in_papka = os.listdir(f'//ALYANS/ovk/Arhiv/{name_line}/Elements/{name_papka}/') 

                for name_exel in all_file_in_papka:
                    if name_exel.endswith(".xlsx"):

                        #print(name_exel)

                        arr_name_exels = []
                        arr_name_exels.append(name_exel)

                        for name_exel in arr_name_exels:

                            way_to_exel = f'//ALYANS/ovk/Arhiv/{name_line}/Elements/{name_papka}/{name_exel}'

                            wb = openpyxl.reader.excel.load_workbook(filename=way_to_exel)

                            wb.active
                            sheet = wb.active

                            for i in range(27, 57) :
                                a = sheet['C' + str(i)].value

                                if a == 'Наименование и номер документа, удостоверяющего качество: ':
                                    break
                                else:
                                    for i2 in range(27, 98):
                                        if sheet['G' + str(i2)].value == sheet['R' + str(i)].value:
                                        
                                            zaivka = dell_slo(sheet['C23'].value)
                                            por_act = dell_slo_o(name_papka)

                                            entities = [name_line, por_act, sheet['C2'].value, zaivka, sheet['C' + str(i)].value, sheet['F' + str(i)].value,sheet['R' + str(i)].value,sheet['U' + str(i)].value,sheet['X' + str(i)].value, sheet['K' + str(i2)].value, sheet['X' + str(i2)].value]
                                            # tab = [name_line, por_act, sheet['C2'].value]
                                            
                                            sql_insert(con, entities)
                                           # cur.execute("INSERT INTO svod_avk_elements VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)", tab)
