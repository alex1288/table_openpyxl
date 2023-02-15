# This is a sample Python script.
import openpyxl
import os
from openpyxl.workbook import Workbook

import pandas as pd
from xlrd import open_workbook

def main():


   def convert_file():
      wb = open_workbook('1_xls.xls')

      data = pd.read_excel('1_xls.xls')

      data.to_excel('convert/filename.xlsx', index=False)
      print("good")
   def table_1():
      path = '1_xls.xls'
      number_of_col = 5
      # path = input("Введи полный путь до файла:")
      # number_of_col = int(input("Введи номер столбца по которому будет проходить фильтрация:"))



      # открывает файл экселя который мы потом сортируем
      wb_obj = openpyxl.load_workbook(path)
      sheet_obj = wb_obj.active
      row = sheet_obj.max_row
      columns = sheet_obj.max_column

      print(f'В файле: строк {row} и столбцов {columns}\n')
      # получаем список категорий (в данном случае) и упаковываем их в массив
      category_arr = []
      for i in range(3 , row+1):
         cell_obj = sheet_obj.cell( row=i, column = number_of_col)
         clear_value = cell_obj.value
         if not clear_value in category_arr:
            category_arr.append(clear_value)

      print(f"Список объектов {category_arr}")
      # создаем общую папку для всех проектов
      if not os.path.isdir("object"):
         os.mkdir("object")

      # обрабатываем таблицу и формируем папки
      for category in category_arr:

         sub_folder_name = str(category)
         print(f'Начало распределения объекта {sub_folder_name} ')
         # создаем папку для определенного проекта
         dir_project = "object/" + sub_folder_name
         if not os.path.isdir(dir_project):
            os.mkdir(dir_project)

         # заново перебираем ячейки с целью отфильтровать нужную категорию в файл

         # Создаем новый файл таблицы
         wb_category = Workbook()
         resultRow = 3
         wb_category_sheet = wb_category.active

         # Создаем первую строку с названием столбцов
         for i in range(1, columns+1):
            title_value = sheet_obj.cell(row=1, column= i).value
            wb_category_sheet.cell(row=1, column= i).value = title_value
            title_value = sheet_obj.cell(row=2, column=i).value
            wb_category_sheet.cell(row=2, column=i).value = title_value
         # Ищем название категории

         try:lookin_for = int(sub_folder_name)
         except:
            print(f'!!! Значение: {sub_folder_name}. Обработка цикла прекращена !!!')
            break
         for i in range(2, row):
            # Здесь мы определяемм значение ящейки в которой находится код
            value = sheet_obj.cell(row=i, column=number_of_col).value
            # print(value)
            try:value = int(value)
            except:pass
            # Сравниваем на совпаденение полученый код с тем что мы ищем в цикле
            if value == lookin_for:
               # Запускаем цикл в котором заполняем ячейки в новом файле одна за одной
               for j in range(1, columns + 1):
                  value = sheet_obj.cell(row=i, column=j).value
                  wb_category_sheet.cell(row=resultRow, column=j).value = value
                  # print(f'perezapisan {value}')
               resultRow += 1

         # print(f'Распределение объекта {lookin_for} завершено')

         # Сохраняем файл таблицы в определенную папку с определенным именем. Имя задается названием категории
         wb_category.save(dir_project + '/' + sub_folder_name + '_таблица_1.xlsx')
         print(f'   >>>Распределение объекта {lookin_for} завершена.\n   Файл сохранен: {sub_folder_name}.xlsx')
      print('Обработка данных завершена')


   def table_6():
      # path = 'тест 2xlsx.xlsx'
      number_of_col = 2
      path = input("Введи полный путь до файла:")
      # number_of_col = int(input("Введи номер столбца по которому будет проходить фильтрация:"))




      # открывает файл экселя который мы потом сортируем
      wb_obj = openpyxl.load_workbook(path)
      sheet_obj = wb_obj.active
      row = sheet_obj.max_row
      columns = sheet_obj.max_column

      print(f'В файле: строк {row} и столбцов {columns}\n')
      # получаем список категорий (в данном случае) и упаковываем их в массив
      category_arr = []
      for i in range(4 , row+1):
         cell_obj = sheet_obj.cell( row=i, column = number_of_col)
         clear_value = cell_obj.value
         if not clear_value in category_arr:
            category_arr.append(clear_value)

      print(f"Список объектов {category_arr}")
      # создаем общую папку для всех проектов
      if not os.path.isdir("object"):
         os.mkdir("object")

      # обрабатываем таблицу и формируем папки
      for category in category_arr:

         sub_folder_name = str(category)
         print(f'Начало распределения объекта {sub_folder_name} ')
         # создаем папку для определенного проекта
         sub_folder_name_dir = sub_folder_name.split(' ')[0]
         dir_project = "object/" + sub_folder_name_dir
         if not os.path.isdir(dir_project):
            os.mkdir(dir_project)

         # заново перебираем ячейки с целью отфильтровать нужную категорию в файл

         # Создаем новый файл таблицы
         wb_category = Workbook()
         resultRow = 4
         wb_category_sheet = wb_category.active

         # Создаем первую строку с названием столбцов
         for i in range(1, columns+1):
            title_value = sheet_obj.cell(row=1, column= i).value
            wb_category_sheet.cell(row=1, column= i).value = title_value
            title_value = sheet_obj.cell(row=2, column=i).value
            wb_category_sheet.cell(row=2, column=i).value = title_value
            title_value = sheet_obj.cell(row=3, column=i).value
            wb_category_sheet.cell(row=3, column=i).value = title_value
         # Ищем название категории

         lookin_for = sub_folder_name

         for i in range(4, row):
            # Здесь мы определяемм значение ящейки в которой находится код
            value = sheet_obj.cell(row=i, column=number_of_col).value

            # Сравниваем на совпаденение полученый код с тем что мы ищем в цикле
            if value == lookin_for:
               # Запускаем цикл в котором заполняем ячейки в новом файле одна за одной
               for j in range(1, columns + 1):
                  value = sheet_obj.cell(row=i, column=j).value
                  wb_category_sheet.cell(row=resultRow, column=j).value = value
               resultRow += 1


         # print(f'Распределение объекта {lookin_for} завершено')

         # Сохраняем файл таблицы в определенную папку с определенным именем. Имя задается названием категории
         wb_category.save(dir_project + '/' + sub_folder_name_dir + '_таблица_6.xlsx')
         print(f'   >>>Распределение объекта {lookin_for} завершена.\n   Файл сохранен: {sub_folder_name}.xlsx')
      print('Обработка данных завершена')


   def table_number():
      while True:
         table_number = int(input('Какую таблицу обрабатываем?'
                                  '\n если таблица номер 6 то введи цифру 6'
                                  '\n если таблица номер 1 то введи цифру 1: \n'))
         if table_number == 1:
            table_1()
            return True
         elif table_number == 6:
            table_6()
            return True
         else:
            print("Не понял. Давай еще раз")

   # Запускаем обработку функций здесь
   # table_number()
   convert_file()
   # table_1()
   # table_6()

# Press the green button in the gutter to run the script.
if __name__ == '__main__':
   main()


