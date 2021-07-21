# -*- coding: utf-8 -*-
#|1

import os
import pythoncom
from win32com.client import Dispatch, gencache
from tkinter import filedialog
import tkinter as tk

#  Подключим константы API Компас
const = gencache.EnsureModule("{75C9F5D0-B5B8-4526-8681-9903C567D2ED}", 0, 1, 0).constants
const_3d = gencache.EnsureModule("{2CAF168C-7961-4B90-9DA2-701419BEEFE3}", 0, 1, 0).constants

#  Подключим описание интерфейсов API5
KAPI = gencache.EnsureModule("{0422828C-F174-495E-AC5D-D31014DBBE87}", 0, 1, 0)
iKompasObject = KAPI.KompasObject(Dispatch("Kompas.Application.5")._oleobj_.QueryInterface(KAPI.KompasObject.CLSID, pythoncom.IID_IDispatch))

#  Подключим описание интерфейсов API7
KAPI7 = gencache.EnsureModule("{69AC2981-37C0-4379-84FD-5DD2F3C0A520}", 0, 1, 0)
application = KAPI7.IApplication(Dispatch("Kompas.Application.7")._oleobj_.QueryInterface(KAPI7.IApplication.CLSID, pythoncom.IID_IDispatch))

#########################################
# 				ФУНКЦИИ					#
#########################################

# Функция получает список позиций из объектов iMarkInsideForm
def get_positions_list():	
	mark_list = []
	positions_list = []
	for i in range(len(iViews)):
		iView = iViews.View(i) #Интерфейс вида графического документа
		iBuildingContainer = KAPI7.IBuildingContainer(iView) # Контейнер объектов СПДС
		iMarks = iBuildingContainer.Marks
		for i in range(len(iMarks)):
			iMark = iMarks.Mark(i)
			if iMark.Type == 13012:
				mark_list.append(iMark)
	
	for i in range(len(mark_list)):
		iText = mark_list[i].TextBefore
		iText = iText.Str
		positions_list.append(iText)
	return positions_list

#########################################
# 				MAIN					#
#########################################


# Для выбора каталога со всеми файлами:
# root = tk.Tk()
# root.withdraw()
# folder_selected = filedialog.askdirectory()
# for root, dirs, files in os.walk(folder_selected):
#         for file in files:
#             if file.endswith(".cdw"):
#                 path_file = os.path.join(root,file)
#                 print(path_file)    

# Запускаем окно tkinter для открытия файла
root = tk.Tk()
root.withdraw()
root.filename = filedialog.askopenfilename(initialdir=os.getcwd(), title="Select file",
                                           filetypes=[("Компас-чертеж", "*.cdw")])

#  Получим документ
iDocuments = application.Documents
iDocument = iDocuments.Open(root.filename, False, True) # Путь, невидимый, только для чтения
iKompasDocument2D = KAPI7.IKompasDocument2D(iDocument)

# Получаем Виды
iViewsAndLayersManager = iKompasDocument2D.ViewsAndLayersManager # Менеджер слоев и видов графического документа
iViews = iViewsAndLayersManager.Views # Интерфейс коллекции видов графического документа

positions_list = get_positions_list()
print(positions_list)

