# -*- coding: utf-8 -*-
'''Copy and Paste Units from Another File'''

__title__ = "Copy Units"
__author__ = "prakritisrimal"

from pyrevit import script, forms, revit
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import *
from System.Collections.Generic import List
import os
import xlrd
output = script.get_output()
ui_doc = __revit__.ActiveUIDocument
doc     = __revit__.ActiveUIDocument.Document # Get the Active Document
app     = __revit__.Application # Returns the Revit Application Object

script_dir = os.path.dirname(__file__)
parent_dir = os.path.abspath(os.path.join(script_dir, "..", ".."))
excel_filename = "Room Location.xlsx"
excel_path = os.path.join(parent_dir, excel_filename)

class MyCopyHandler(IDuplicateTypeNamesHandler):
    def OnDuplicateTypeNamesFound(self, args):
        return DuplicateTypeAction.UseDestinationTypes

unique_room_names = set()
rooms = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Rooms).WhereElementIsNotElementType().ToElements()
for room in rooms:
    room_name = room.LookupParameter("Name").AsString()
    if room_name:
        unique_room_names.add(room_name)

if not unique_room_names:
    script.exit()

sorted_room_names = sorted(unique_room_names)
selected_room_name = forms.SelectFromList.show(sorted_room_names, multiselect = False, title = 'Select Unit Type that has to be placed')

if not selected_room_name:
    script.exit()

t = Transaction (doc, ("Copy Unit Typologies"))
t.Start()

excel_workbook = xlrd.open_workbook(excel_path)
excel_worksheet = excel_workbook.sheet_by_index(0)
excel_worksheet_file_location = excel_workbook.sheet_by_index(1)
excel_typology_names = []
file_location = None
levels = FilteredElementCollector(doc).OfClass(Level).ToElements()
for row in range (1, excel_worksheet_file_location.nrows):
    excel_typology_names.append(excel_worksheet_file_location.cell_value(row,0))
    #print (name)
for row, name in enumerate(excel_typology_names, start=1):
    if name == selected_room_name:
        file_location = excel_worksheet_file_location.cell_value(row,1)
        print (file_location)
        model_path = ModelPathUtils.ConvertUserVisiblePathToModelPath(file_location)
excel_typology_names = []               
for row in range(1, excel_worksheet.nrows):
    excel_typology_names.append(excel_worksheet.cell_value(row,0))
for row, name in enumerate(excel_typology_names, start=1):
    if name == selected_room_name:
        # Setup options to open the file
        open_options = OpenOptions()
        if open_options.DetachFromCentralOption:
            open_options.DetachFromCentralOption = None
            open_options.Audit = False
        # Open the file
        revit_doc = app.OpenDocumentFile(model_path, open_options)
        if not revit_doc:
            script.exit()
        # Get Model Group from the revit file 
        model_groups = FilteredElementCollector(revit_doc).OfCategory(BuiltInCategory.OST_IOSModelGroups).WhereElementIsNotElementType().ToElements()
        #print(len(model_groups))
        group_ids =  List[ElementId]()
        #name = excel_worksheet.cell_value(row, 0)
        for group in model_groups:
            #print (group.Name)
            if group.Name in name:
                group_to_copy = group
                group_location = group_to_copy.Location.Point
                group_ids.Add(group_to_copy.Id)     
                
        #print (group_location)
        #print (len(group_ids))
        # translation_vector = XYZ(0, 0, 0) - group_location
        # translation_transform = Transform.CreateTranslation(translation_vector)
        options = CopyPasteOptions()
        options.SetDuplicateTypeNamesHandler(MyCopyHandler())
        copied_element_ids = ElementTransformUtils.CopyElements(revit_doc, group_ids, doc, Transform.Identity, options)

        target_point_strs = []
        target_point_strs.append(excel_worksheet.cell_value(row,1))
        for target_point_str in target_point_strs:
            #print (target_point_str)
            target_point_coords = list(map(float, target_point_str.split(',')))
            target_point = XYZ(target_point_coords[0], target_point_coords[1], target_point_coords[2])
            target_z_value = target_point.Z
        if copied_element_ids:
            copied_element_id = copied_element_ids[0]
            copied_group = doc.GetElement(copied_element_id)
            current_level_param = copied_group.LookupParameter("Reference Level")
            if current_level_param:
                current_level_id = current_level_param.AsElementId()
                current_level = doc.GetElement(current_level_id)
                current_level_elevation = current_level.Elevation
                print("Current level: {}, Elevation: {}".format(current_level.Name, current_level_elevation))
            target_level = None
            for level in levels:
                if abs(level.Elevation - target_z_value) < 0.001:  # Adjust the tolerance as needed
                    target_level = level
                    break
            group_location = copied_group.Location.Point
            print ("Copied Group Location Point{}".format(group_location))
            translation_vector = target_point
            ElementTransformUtils.MoveElement(doc, copied_element_id, translation_vector)
            group_location_after_move = copied_group.Location.Point
            if target_level:
                print("Target level: {}, Elevation: {}".format(target_level.Name, target_level.Elevation))
                if current_level_param:
                    current_level_param.Set(target_level.Id)
                level_offset = copied_group.LookupParameter("Origin Level Offset")
                if level_offset:
                    level_offset.Set(0)
                print("Group moved to new level: {}".format(target_level.Name))
            else:
                print("No target level found for Z-value: {}".format(target_z_value))
        print ("Model Group {} (ID:{}) copied to {}".format(copied_group,output.linkify(copied_element_id),group_location_after_move))
        # copied_group.UngroupMembers()
t.Commit()



#  #Prompt user to select the file with the Unit Type
# forms.alert("Please select the Unit Typology File", title="File Selection", warn_icon=False)
# unit_file = forms.pick_file(file_ext='rvt', title="Select the Unit Typology file")
# if not unit_file:
#     script.exit()

# model_path = ModelPathUtils.ConvertUserVisiblePathToModelPath(unit_file)




# unique_line_style_name = set()
# model_lines = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Lines).WhereElementIsNotElementType().ToElements()
# for model_line in model_lines:
#     if isinstance(model_line, ModelCurve):
#         # Get the line style of the model line
#         line_style = model_line.LineStyle
#         if line_style:
#             line_style_name = line_style.Name
#             if line_style_name:
#                 unique_line_style_name.add(line_style_name)
            
# if not unique_line_style_name:
#     script.exit()

# sorted_line_names = sorted(unique_line_style_name)

# selected_line_style = forms.SelectFromList.show(sorted_line_names, multiselect = False, title = 'Select Line Style of Model Lines')
# if not selected_line_style:
#     script.exit()












    




