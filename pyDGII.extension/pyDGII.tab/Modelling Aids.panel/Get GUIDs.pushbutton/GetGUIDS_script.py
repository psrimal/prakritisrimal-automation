# -*- coding: utf-8 -*-
'''Write GUID to excel'''

__title__ = "Get GUIDs"
__author__ = "prakritisrimal"

from pyrevit import script, forms, revit
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import *
from Autodesk.Revit.UI.Selection import Selection, ObjectType, ISelectionFilter
from System.Collections.Generic import List
import os
import csv
output = script.get_output()
ui_doc = __revit__.ActiveUIDocument
doc     = __revit__.ActiveUIDocument.Document # Get the Active Document
app     = __revit__.Application # Returns the Revit Application Object

model_groups = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_IOSModelGroups).WhereElementIsNotElementType().ToElements()
all_write_datas = []
for group in model_groups:
    group_member_ids = group.GetMemberIds()

    for element_id in group_member_ids:
        element = group.Document.GetElement(element_id)
        element_name = element.Name
        guid = element.UniqueId
        # print (element_name)
        # print (guid)
        # print ("---")
        write_data = [element_name, guid]
        all_write_datas.append(write_data)

csv_file_path = "C:\Users\psrimal\Desktop\Diryah Automation\prakritisrimal-automation\pyDGII.extension\pyDGII.tab\GUID.csv"
try:
    # Append data to the CSV file
    with open(csv_file_path, mode='ab') as file:  # 'ab' for append and binary mode in Python 2.7
        writer = csv.writer(file)

        # If the file is new and empty, add the header row
        if file.tell() == 0:  # Checks if the file is empty
            writer.writerow(["Element Name", "GUID"])

        # Write the new data row
        for data in all_write_datas:
            writer.writerow(data)
except:
    pass
