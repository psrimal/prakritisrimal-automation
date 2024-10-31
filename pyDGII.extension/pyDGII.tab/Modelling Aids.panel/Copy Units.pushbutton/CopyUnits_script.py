# -*- coding: utf-8 -*-
'''Copy and Paste Units from Another File'''

__title__ = "Copy Units"
__author__ = "prakritisrimal"

from pyrevit import script, forms, revit
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import *
from System.Collections.Generic import List
import os
import xlrd, csv
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

# try:
success = True
t = Transaction (doc, ("Copy Unit Typologies"))
t.Start()
skipped_data=[]
excel_workbook = xlrd.open_workbook(excel_path)
excel_worksheet = excel_workbook.sheet_by_index(0)
excel_worksheet_file_location = excel_workbook.sheet_by_index(1)
excel_typology_names = []
file_location = None
all_unit_file_data = []
all_id_datas = []
levels = FilteredElementCollector(doc).OfClass(Level).ToElements()
for row in range (1, excel_worksheet_file_location.nrows):
    excel_typology_names.append(excel_worksheet_file_location.cell_value(row,0))
    #print (name)

copy_counter = 0
for row, name in enumerate(excel_typology_names, start=1):
    if name == selected_room_name:
        file_location = excel_worksheet_file_location.cell_value(row,1)
        #print (file_location)
        if file_location:
            model_path = ModelPathUtils.ConvertUserVisiblePathToModelPath(file_location)
            # Setup options to open the file
            open_options = OpenOptions()
            # if open_options.DetachFromCentralOption:
            #     open_options.DetachFromCentralOption = DetachFromCentralOption.DetachAndPreserveWorksets
            #     open_options.Audit = False
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
                    group_member_ids = group.GetMemberIds()

                    for element_id in group_member_ids:
                        element = group.Document.GetElement(element_id)
                        element_name = element.Name
                        guid = element.UniqueId
                        category = element.Category
                        if element.Category.Name == 'Walls' and element.Category is not None:
                            wall = element
                            wall_name = wall.Name
                            wall_location = wall.Location
                            wall_location_curve = wall_location.Curve
                            wall_start = wall_location_curve.GetEndPoint(0)
                            wall_end = wall_location_curve.GetEndPoint(1)
                            wall_orientation = wall.Orientation
                            wall_fire = wall.LookupParameter ("Fire_Rating")
                            wall_fire_rating = wall_fire.AsValueString()
                            wall_stc = wall.LookupParameter ("STC_Rating")
                            wall_stc_rating = wall_stc.AsValueString()
                            wall_structural =  wall.LookupParameter("Structural").AsInteger()
                            unit_file_data = [guid, element_id, element.Category.Name, wall_name, wall_start, 
                                         wall_end, wall_orientation, wall_fire_rating, wall_stc_rating, wall_structural]
                            all_unit_file_data.append(unit_file_data)


            csv_file_path = "C:\\Users\\psrimal\\Desktop\\Diryah Automation\\prakritisrimal-automation\\pyDGII.extension\\pyDGII.tab\\" + name + "\\" + name + ".csv"
            try:
                # Append data to the CSV file
                with open(csv_file_path, mode='ab') as file:  # 'ab' for append and binary mode in Python 2.7
                    writer = csv.writer(file)

                    # If the file is new and empty, add the header row
                    if file.tell() == 0:  # Checks if the file is empty
                        writer.writerow(["GUID", "Element ID", "Category", "Wall Type Name", "Wall Start Point", "Wall End Point", "Wall Orientation", "Fire_Rating",
                                         "STC_Rating", "Structural Property"])

                    # Write the new data row
                    for data in all_unit_file_data:
                        writer.writerow(data)
            except:
                pass
                 
            #print (group_location)
            #print (len(group_ids))
            # translation_vector = XYZ(0, 0, 0) - group_location
            # translation_transform = Transform.CreateTranslation(translation_vector)
            
            options = CopyPasteOptions()
            options.SetDuplicateTypeNamesHandler(MyCopyHandler())
            initial_copy_ids  = ElementTransformUtils.CopyElements(revit_doc, group_ids, doc, Transform.Identity, options)

        else:
            skipped_group_data = [selected_room_name, "UNIT TYPOLOGY FILE NOT FOUND"]
            skipped_data.append(skipped_group_data)

moved_data =[]
excel_typology_names = []
target_point_strs = []
all_copied_file_data = []
copied_at_target_ids = []
for row in range (1, excel_worksheet.nrows):
    excel_typology_names.append(excel_worksheet.cell_value(row,0))
for row, name in enumerate(excel_typology_names, start=1):
    if name == selected_room_name:
        target_point_strs.append(excel_worksheet.cell_value(row,1))        
if len(target_point_strs)>0: 
    counter = 0
    for target_point_str in target_point_strs:
        #print (target_point_str)
        target_point_coords = list(map(float, target_point_str.split(',')))
        target_point = XYZ(target_point_coords[0], target_point_coords[1], target_point_coords[2])
        target_z_value = target_point.Z
        if initial_copy_ids :
            initial_copy_id = initial_copy_ids [0]
            copied_group = doc.GetElement(initial_copy_id)
            group_location = copied_group.Location.Point
        # print ("Copied Group Location Point{}".format(group_location))
            translation_vector = target_point
            copied_at_target_ids  = ElementTransformUtils.CopyElement(doc, initial_copy_id, translation_vector)
            #counter = 0
            if copied_at_target_ids:
                if counter == 0:
                    copied_id = copied_at_target_ids
                    counter +=1
                    if copied_id:
                        copied_id = copied_id [0]
                        #print (copied_id)
                        copied_group = doc.GetElement (copied_id)
                        copied_element_ids = copied_group.GetMemberIds()
                        for copied_element_id in copied_element_ids:
                            #print (copied_element_id)
                            copied_element = copied_group.Document.GetElement(copied_element_id)
                            copied_element_name = copied_element.Name
                            copied_guid = copied_element.UniqueId
                            copied_category = copied_element.Category
                            if copied_element.Category.Name == 'Walls' and copied_element.Category is not None:
                                copied_wall = copied_element
                                copied_wall_name = copied_wall.Name
                                copied_wall_location = copied_wall.Location
                                copied_wall_location_curve = copied_wall_location.Curve
                                copied_wall_start = copied_wall_location_curve.GetEndPoint(0)
                                copied_wall_end = copied_wall_location_curve.GetEndPoint(1)
                                copied_wall_orientation = copied_wall.Orientation
                                copied_wall_fire = copied_wall.LookupParameter ("Fire_Rating")
                                copied_wall_fire_rating = copied_wall_fire.AsValueString()
                                copied_wall_stc = copied_wall.LookupParameter ("STC_Rating")
                                copied_wall_stc_rating = copied_wall_stc.AsValueString()
                                copied_wall_structural =  copied_wall.LookupParameter("Structural").AsInteger()
                                copied_file_data = [ copied_guid, copied_element_id, copied_element.Category.Name, copied_wall_name, copied_wall_start, 
                                                copied_wall_end, copied_wall_orientation, copied_wall_fire_rating, copied_wall_stc_rating, copied_wall_structural]
                                all_copied_file_data.append(copied_file_data)

                        csv_file_path = "C:\\Users\\psrimal\\Desktop\\Diryah Automation\\prakritisrimal-automation\\pyDGII.extension\\pyDGII.tab\\" + copied_group.Name + "\\Copied "  + copied_group.Name + ".csv"
                        try:
                            # Append data to the CSV file
                            with open(csv_file_path, mode='ab') as file:  # 'ab' for append and binary mode in Python 2.7
                                writer = csv.writer(file)

                                # If the file is new and empty, add the header row
                                if file.tell() == 0:  # Checks if the file is empty
                                    writer.writerow(["Copied GUID", "Copied Element ID", "Copied Category", "Copied Wall Type Name", "Copied Wall Start Point", 
                                                    "Copied Wall End Point", "Copied Wall Orientation", "Copied Fire_Rating",
                                                    "Copied STC_Rating", "Copied Structural Property"])

                                # Write the new data row
                                for data in all_copied_file_data:
                                    writer.writerow(data)
                        except:
                            pass

            for id in copied_at_target_ids:
                copied_at_target = doc.GetElement(id)
                group_location_after_move = copied_at_target.Location.Point
                current_level_param = copied_at_target.LookupParameter("Reference Level")
                if current_level_param:
                    current_level_id = current_level_param.AsElementId()
                    current_level = doc.GetElement(current_level_id)
                    current_level_elevation = current_level.Elevation
                # print("Current level: {}, Elevation: {}".format(current_level.Name, current_level_elevation))
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
                #     print("Group moved to new level: {}".format(target_level.Name))
                # else:
                #     print("No target level found for Z-value: {}".format(target_z_value))
                copied_at_target.UngroupMembers()
                moved_group_data = [selected_room_name, target_level.Name]
                moved_data.append(moved_group_data)


        else:
            skipped_group_data = [selected_room_name, "UNABLE TO COPY"]
            skipped_data.append(skipped_group_data)
else:
    skipped_group_data = [selected_room_name, "UNIT LOCATION NOT FOUND"]
    skipped_data.append(skipped_group_data)


for initial_copy_id in initial_copy_ids:
    doc.Delete (initial_copy_id)

t.Commit()


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










    




