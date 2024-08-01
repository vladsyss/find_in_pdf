import pandas as pd
import re
from contextlib import redirect_stdout
import os
import sys

if not os.path.exists('Расшифровка.pdf'):
    print ('\nОтсутствует файл - Расшифровка.pdf\n')
    input ("Нажмите OK для выхода из программы.\n")
    sys.exit()

if not os.path.exists('telefons.ini'):
    print ('\nОтсутствует файл - telefons.ini\n')
    input ("Нажмите OK для выхода из программы.\n")
    sys.exit()

if os.path.exists('output.txt'):
    os.remove('output.txt')
'''
df_orders = pd.read_excel('ТЕЛЕФОННЫЙ СПРАВОЧНИК.xls', sheet_name = 'Лист1', dtype='str')

df_orders.fillna('', inplace= True )

telefons = df_orders.iloc[:,2].tolist()

telefons = list(filter(None, telefons))

#print(telefons)
'''
with open('telefons.ini', 'r') as file:
    telefons = file.readlines()

a = []

for item in telefons:
    numbers = re.findall(r'\b\d+\b', str(item))
    numbers = ''.join(numbers)
    if len (numbers) == 5:
        #print(numbers, len(numbers))
        #print(numbers)
        #print (item)
        a.append(numbers)
    else:
        #print(numbers)
        b = re.findall(r'\d\d\d\d\d', numbers)
        for x in b:
            #print(x)
            a.append(x)

a.sort()
print ('\nОтсортированный список телефонов.\n')
print (a, len (a))

from itertools import groupby

new_x = [el for el, _ in groupby(a)]

print ('\nОтсортированный список телефонов без дубликатов.\n')
print(new_x, len(new_x))
print ('\n')

# Вывод списка телефонов в файл.
#with open('my_file.txt', 'w') as f:
#    f.writelines(f"{item}\n" for item in new_x)


# https://www.cyberforum.ru/python-beginners/thread3078200.html
import PyPDF2
 
''' 
Функция поиска номеров страниц в PDF документах где встречается фраза
 
Принимает 2 аргумента:
    pdf_filenames: массив строк с именами pdf файлов.
    word: строку, которую нужно искать в pdf файлах.
 
Возвращает:
    Структуру map, где ключом является имя PDF файла,
     а значением множество set содержащее номера страниц pdf файла,
     где встречается слово word.
'''
def find_string_in_pdfs(pdf_filenames, line):
    result = {}
    for pdf_filename in pdf_filenames:
        try:
            with open(pdf_filename, "rb") as f:
                pdf_file = PyPDF2.PdfReader(f)
                pages = set()
                for page_num in range(len(pdf_file.pages)):
                    page = pdf_file.pages[page_num]
                    if line in page.extract_text():
                        pages.add(page_num+1)
                if (len(pages)):
                    result[pdf_filename] = pages
        # Исключение - Если файл не найден
        except FileNotFoundError:
            pass
    return result
 
 
''' Пример использования '''
#pdf_filenames = ["ваш_документ1.pdf", "ваш_документ2.pdf", "ваш_документ3.pdf"] # или если 1 документ, то просто ["ваш_документ1.pdf"]
pdf_filenames = ["Расшифровка.pdf"]
#line = '78339'

for x in new_x:
    line = x

#for x in lines:
#    line = x

    # Поиск строки в PDF документах
    result = find_string_in_pdfs(pdf_filenames, line)
    # Выводим результат, но прежде проверяем на пустоту
    if (len(result)):
        # Вывод 'print' отправляется в 'output.txt'
        print('Номер [', line, '] встретился в:')
        with open('output.txt', 'a') as f, redirect_stdout(f):
            print('Номер [', line, '] встретился в:')
        for pdf_filename, pages in result.items():
            print("--[", pdf_filename, "] на страницах:", pages)
            with open('output.txt', 'a') as f, redirect_stdout(f):
                print("--[", pdf_filename, "] на страницах:", pages)
    else:
        #print('Ничего не найдено')
        pass

print ('\nВывод списка в файл output.txt завершён!\n')
input ("Нажмите OK для выхода из программы.\n")
