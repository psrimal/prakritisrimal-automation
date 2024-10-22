# -*- coding: utf-8 -*-
'''Write Co-ordinates of Selected Units in an excel'''

__title__ = "Write Data"
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



unique_room_names = set()
rooms = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Rooms).WhereElementIsNotElementType().ToElements()
for room in rooms:
    room_name = room.LookupParameter("Name").AsString()
    if room_name:
        unique_room_names.add(room_name)

if not unique_room_names:
    script.exit()

sorted_room_names = sorted(unique_room_names)
selected_room_names = forms.SelectFromList.show(sorted_room_names, multiselect = True, title = 'Select Rooms to Write Base Point Location')

if not selected_room_names:
    script.exit()

target_rooms = []
for room in rooms:
    room_name = room.LookupParameter("Name").AsString()
    if room_name in selected_room_names:
        target_rooms.append(room)

# t = Transaction (doc, "Draw Room Edges")
# t.Start()
options = Options()
boundary_options = SpatialElementBoundaryOptions()
curves = []
#loop = CurveLoop()
edge_points = []
for room in target_rooms:
    room_name = room.LookupParameter("Name").AsString()
    boundaries = room.GetBoundarySegments(boundary_options)
    min_point = None
    max_point = None
    if boundaries:
        # Iterate through each boundary list (usually one per room level)
        for boundary_list in boundaries:
            # Each segment contains a Curve object that defines a boundary segment
            for boundary_segment in boundary_list:
                curve = boundary_segment.GetCurve()
                start_point = curve.GetEndPoint(0)
                end_point = curve.GetEndPoint(1)
                if min_point is None:
                    min_point = XYZ(start_point.X, start_point.Y, start_point.Z)
                    #max_point = XYZ(start_point.X, start_point.Y, start_point.Z)

                # Update min and max points by comparing with start and end points
                for point in [start_point, end_point]:
                    min_point = XYZ(min(min_point.X, point.X), min(min_point.Y, point.Y), min(min_point.Z, point.Z))
                    #max_point = XYZ(max(max_point.X, point.X), max(max_point.Y, point.Y), max(max_point.Z, point.Z))
    min_point_str = "{},{},{}".format(min_point.X, min_point.Y, min_point.Z)
    write_data = [room_name, ' ', min_point_str]
    csv_file_path = "C:\Users\psrimal\Desktop\Diryah Automation\prakritisrimal-automation\pyDGII.extension\pyDGII.tab\Room Location.csv"
    try:
        # Append data to the CSV file
        with open(csv_file_path, mode='ab') as file:  # 'ab' for append and binary mode in Python 2.7
            writer = csv.writer(file)

            # If the file is new and empty, add the header row
            if file.tell() == 0:  # Checks if the file is empty
                writer.writerow(["Room Name ", "File Location", "Room Location"])

            # Write the new data row
            writer.writerow(write_data)
    except:
        pass

#t.Commit()

    # print("Minimum point: X={}, Y={}, Z={}".format(min_point.X, min_point.Y, min_point.Z))
    # print("Maximum point: X={}, Y={}, Z={}".format(max_point.X, max_point.Y, max_point.Z))

# model_group = []
# #Pre-Selected Groups
# selection = ui_doc.Selection.GetElementIds()
# if len(selection) > 0:
#     for id in selection:
#         element = doc.GetElement(id)
#         try:
#             if isinstance(element, Group): 
#                model_group.append(element)

#         except:
#             continue

# for group in model_group:
#     #group_type = group.GroupType
#     group_name = group.Name
#     group_id = group.Id
#     try:
#         location = group.Location.Point
#         location_str = "{},{},{}".format(location.X, location.Y, location.Z)
#     except:
#         continue

#     write_data = [group_name,' ', location_str]
#     csv_file_path = "C:\Users\psrimal\Desktop\Diryah Automation\prakritisrimal-automation\pyDGII.extension\pyDGII.tab\Model Group Location.csv"
#     try:
#         # Append data to the CSV file
#         with open(csv_file_path, mode='ab') as file:  # 'ab' for append and binary mode in Python 2.7
#             writer = csv.writer(file)

#             # If the file is new and empty, add the header row
#             if file.tell() == 0:  # Checks if the file is empty
#                 writer.writerow(["Model Group Name", "File Location", "Model Group Location"])

#             # Write the new data row
#             writer.writerow(write_data)
#     except:
#         pass
