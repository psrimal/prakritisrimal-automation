# -*- coding: utf-8 -*-
'''Mirror Units'''

__title__ = "Mirror Units"
__author__ = "prakritisrimal"

from pyrevit import script, forms, revit
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import *
from Autodesk.Revit.UI.Selection import *
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
reference_room_name = forms.SelectFromList.show(sorted_room_names, multiselect = False, title = 'Select the Unit Type that needs to be Mirrored')
if not reference_room_name:
    script.exit()
target_room_name = forms.SelectFromList.show(sorted_room_names, multiselect = False, title = 'Select the Mirrored Unit Type Name')
if not target_room_name:
    script.exit()

axes = ["X-axis", "Y-axis"]
axis_choice = forms.SelectFromList.show(axes, multiselect = False, title = "Choose the axes for mirroring" )
if not axis_choice:
    script.exit()
elif axis_choice == "X-axis":
    plane = Plane.CreateByNormalAndOrigin(XYZ(0,1,0), XYZ(0,0,0))
elif axis_choice == "Y-axis":
    plane = Plane.CreateByNormalAndOrigin(XYZ(1,0,0), XYZ(0,0,0))
# elif axis_choice == "Custom Line":
#     forms.alert("Ensure that you have a pre-drawn Detail Line")
#     reference_element = ui_doc.Selection.PickObject(ObjectType.Element, "Select the Line along which the Unit has to be Mirrored")
#     reference_line = doc.GetElement(reference_element)
#     location_curve = reference_line.Location
#     if isinstance(location_curve, LocationCurve):
#         curve = location_curve.Curve 
#     start_point = curve.GetEndPoint(0)
#     plane = Plane.CreateByNormalAndOrigin(curve.Direction, start_point)


try:
    success = True
    t = Transaction (doc, ("Mirror Unit Typologies"))
    t.Start()
    skipped_data =[]
    moved_data = []
    excel_workbook = xlrd.open_workbook(excel_path)
    excel_worksheet = excel_workbook.sheet_by_index(0)
    excel_worksheet_file_location = excel_workbook.sheet_by_index(1)
    excel_typology_names = []
    file_location = None
    levels = FilteredElementCollector(doc).OfClass(Level).ToElements()
    for row in range (1, excel_worksheet_file_location.nrows):
        excel_typology_names.append(excel_worksheet_file_location.cell_value(row,0))
        #print (name)

    copy_counter = 0
    for row, name in enumerate(excel_typology_names, start=1):
        if name == reference_room_name:
            file_location = excel_worksheet_file_location.cell_value(row,1)
            #print (file_location)
            if file_location:
                model_path = ModelPathUtils.ConvertUserVisiblePathToModelPath(file_location)
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
                options = CopyPasteOptions()
                options.SetDuplicateTypeNamesHandler(MyCopyHandler())
                initial_copy_ids  = ElementTransformUtils.CopyElements(revit_doc, group_ids, doc, Transform.Identity, options)
                if initial_copy_ids:
                    initial_copy_id = initial_copy_ids[0]
                    ElementTransformUtils.MirrorElement(doc,initial_copy_id,plane)
                    mirrored_group = doc.GetElement(initial_copy_id)
                    location = mirrored_group.Location
                    location_point = location.Point
                    translation_vector = XYZ (0,0,0) - location_point
                    ElementTransformUtils.MoveElement(doc, initial_copy_id, translation_vector)
                else:
                    skipped_group_data = [reference_room_name, "UNABLE TO COPY"]
                    skipped_data.append(skipped_group_data)
            else:
                skipped_group_data = [reference_room_name, "UNIT TYPOLOGY FILE NOT FOUND"]
                skipped_data.append(skipped_group_data)

    excel_typology_names = []
    target_point_strs = []
    for row in range (1, excel_worksheet.nrows):
        excel_typology_names.append(excel_worksheet.cell_value(row,0))
    for row, name in enumerate(excel_typology_names, start=1):
        if name == target_room_name:
            target_point_strs.append(excel_worksheet.cell_value(row,1))
    if len(target_point_strs)>0: 
        for target_point_str in target_point_strs:
            #print (target_point_str)
            target_point_coords = list(map(float, target_point_str.split(',')))
            target_point = XYZ(target_point_coords[0], target_point_coords[1], target_point_coords[2])
            target_z_value = target_point.Z
            if initial_copy_id:
                copied_group = doc.GetElement(initial_copy_id)
                group_location = copied_group.Location.Point
                #print ("Mirrored and Moved Group Location Point{}".format(group_location))
                translation_vector = target_point
                copied_at_target_ids  = ElementTransformUtils.CopyElement(doc, initial_copy_id, translation_vector)
                #print ("Copied at target number".format(len(copied_at_target_ids)))
                for copied_at_target_id in copied_at_target_ids:
                    copied_at_target = doc.GetElement(copied_at_target_id)
                    group_location_after_move = copied_at_target.Location.Point
                    current_level_param = copied_at_target.LookupParameter("Reference Level")
                    if current_level_param:
                        current_level_id = current_level_param.AsElementId()
                        current_level = doc.GetElement(current_level_id)
                        current_level_elevation = current_level.Elevation
                        #print("Current level: {}, Elevation: {}".format(current_level.Name, current_level_elevation))
                    target_level = None
                    for level in levels:
                        if abs(level.Elevation - target_z_value) < 0.001:  # Adjust the tolerance as needed
                            target_level = level
                            break
                    if target_level:
                        #print("Target level: {}, Elevation: {}".format(target_level.Name, target_level.Elevation))
                        if current_level_param:
                            current_level_param.Set(target_level.Id)
                    level_offset = copied_at_target.LookupParameter("Origin Level Offset")
                    if level_offset:
                        level_offset.Set(0)

                    moved_group_data = [target_room_name, target_level.Name]
                    moved_data.append(moved_group_data)
                    #     print("Group moved to new level: {}".format(target_level.Name))
                    # else:
                    #     print("No target level found for Z-value: {}".format(target_z_value))
            else:
                skipped_group_data = [reference_room_name, "UNABLE TO COPY"]
                skipped_data.append(skipped_group_data)

                print ("Model Group {} (ID:{}) copied to {}".format(copied_group,output.linkify(copied_at_target_id),group_location_after_move))
                copied_at_target.UngroupMembers()
    else:
        skipped_group_data = [target_room_name, "UNIT LOCATION NOT FOUND"]
        skipped_data.append(skipped_group_data)
    doc.Delete(initial_copy_id)
    #print (copy_counter)
    t.Commit()
except Exception as e:
    success = False  
    output.print_md("##⚠️ Error occurred: {}".format(e))
    t.RollBack()

if success:
    if moved_data:
        output.print_md("##⚠️ {} Completed.😊 ".format(__title__))
        output.print_md("---")
        output.print_md("✅ Units Copied. Refer to the **Table Report** below for reference")
        output.print_table(table_data=moved_data, columns=["UNIT NAME", "LEVEL"])
        output.print_md("---")

    if skipped_data:
        output.print_md("##⚠️ {} Completed. Issues Found ☹️".format(__title__))
        output.print_md("---")
        output.print_md("❌ Some Unit Typologies were not Copied. Refer to the **Table Report** below for reference")
        output.print_table(table_data=skipped_data, columns=["UNIT NAME","ERROR CODE"])
        output.print_md("---")
        output.print_md("***✅ ERROR CODE REFERENCE***")
        output.print_md("---")
        output.print_md("**UNIT TYPOLOGY FILE NOT FOUND** - Ensure that the Excel file has the Unit's File Location.  \n")
        output.print_md("**UNABLE TO COPY** - Unit Typology was not copied. Check manually. \n")
        output.print_md("**UNIT LOCATION NOT FOUND** - Unable to read the Unit's Coordinates. Ensure that it is entered in the correct format. \n")
        output.print_md("---")











    




