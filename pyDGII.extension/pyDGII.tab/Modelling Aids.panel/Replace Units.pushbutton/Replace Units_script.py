# -*- coding: utf-8 -*-
'''Replace Units'''

__title__ = "Replace Units"
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

file_path = "C:\Users\psrimal\Desktop\Diryah Automation\prakritisrimal-automation\pyDGII.extension\pyDGII.tab\Room Names.csv"
unit_names = []
try:
    with open (file_path) as csvfile:
        reader = csv.DictReader(csvfile)
        for row in reader:
            unit_names.append(row["Room Name"])
except:
    pass

selected_unit_name = forms.SelectFromList.show(unit_names, multiselect = False, title = 'Select Unit Type that has to be replaced')
if not selected_unit_name:
    script.exit()

master_file_path = "C:\\Users\\psrimal\\Desktop\\Diryah Automation\\prakritisrimal-automation\\pyDGII.extension\\pyDGII.tab\\" + selected_unit_name + '\\' +  selected_unit_name + ".csv"
master_guids = []
master_ids = []
master_categorys = []
master_type_names = []
master_start_points = []
master_end_points = []
master_orientations = []
master_fire_ratings = []
master_stc_ratings = []
master_structurals = []

try:
    with open (master_file_path) as csvfile:
        reader = csv.DictReader(csvfile)
        for row in reader:
            master_guids.append(row["GUID"])
            master_ids.append(row["Element ID"])
            master_categorys.append(row["Category"])
            master_type_names.append(row["Wall Type Name"])
            master_start_points.append(row["Wall Start Point"])
            master_end_points.append(row["Wall End Point"])
            master_orientations.append(row["Wall Orientation"])
            master_fire_ratings.append(row["Fire_Rating"])
            master_stc_ratings.append(row["STC_Rating"])
            master_structurals.append(row["Structural Property"])
except Exception as e:
    print ("Error reading master file:{}".format(e))
    pass

copy_file_path = "C:\\Users\\psrimal\\Desktop\\Diryah Automation\\prakritisrimal-automation\\pyDGII.extension\\pyDGII.tab\\" + selected_unit_name + '\\Copied ' +  selected_unit_name + ".csv"
copy_guids = []
copy_ids = []
copy_categorys = []
copy_type_names = []
copy_start_points = []
copy_end_points = []
copy_orientations = []
copy_fire_ratings = []
copy_stc_ratings = []
copy_structurals = []

try:
    with open (copy_file_path) as csvfile:
        reader = csv.DictReader(csvfile)
        for row in reader:
            copy_guids.append(row["Copied GUID"])
            copy_ids.append(row["Copied Element ID"])
            copy_categorys.append(row["Copied Category"])
            copy_type_names.append(row["Copied Wall Type Name"])
            copy_start_points.append(row["Copied Wall Start Point"])
            copy_end_points.append(row["Copied Wall End Point"])
            copy_orientations.append(row["Copied Wall Orientation"])
            copy_fire_ratings.append(row["Copied Fire_Rating"])
            copy_stc_ratings.append(row["Copied STC_Rating"])
            copy_structurals.append(row["Copied Structural Property"])
except Exception as e:
    print ("Error reading copied file:{}".format(e))
    pass

coordinates_file_path = "C:\Users\psrimal\Desktop\Diryah Automation\prakritisrimal-automation\pyDGII.extension\pyDGII.tab\Room Location.csv"
coordinates = []
room_names = []

try:
    with open (coordinates_file_path) as csvfile:
        reader = csv.DictReader(csvfile)
        for row in reader:
            room_names.append(row["Room Name "])
            for room_name in room_names:
                if room_name == selected_unit_name:
                    coordinates.append(row["Room Location"])
except Exception as e:
    print ("Error reading coordinates:{}".format(e))
    pass

mismatched_data = []
skipped_data = []
unique_wall_names = set()
for i, (master_category, copy_category) in enumerate(zip(master_categorys, copy_categorys)):
    if master_category != copy_category:
        element_id = ElementId(int(copy_ids[i]))
        mismatched_category_data = [output.linkify(element_id), master_category, copy_category, "CATEGORY MISMATCH"]
        mismatched_data.append(mismatched_category_data)

for i, (master_type_name, copy_type_name) in enumerate(zip(master_type_names, copy_type_names)):
    if master_type_name != copy_type_name:
        element_id = ElementId(int(copy_ids[i]))
        element = doc.GetElement (element_id)
        category_name = element.Category.Name
        element_name = element.Name
        if category_name == "Walls":
            wall_type = element.WallType
            wall_collector = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Walls).WhereElementIsNotElementType().ToElements()
            master_wall_type = None
            for wall in wall_collector:
                if wall.Name == master_type_name:
                    master_wall_type = wall.WallType
                    break
            if master_wall_type is not None:
                t = Transaction (doc, "Change Wall Type")
                t.Start()
                try: 
                    element.WallType = master_wall_type
                    t.Commit()
                except Exception as e:
                    print ("Wall Type Could not be changed : {}".format(e))

        # elif category_name == "Doors":
        #     category = BuiltInCategory.OST_Doors
        # all_family_types = FilteredElementCollector(doc).OfCategory(category).OfClass(FamilySymbol).ToElements()
        # print(len(all_family_types))
        # for family_type in all_family_types:
        #     print (family_type.Name)
        # element_family = element.Symbol.Family
        # family_name = element_family.Name
            else:
                skipped_type_data = [output.linkify(element_id), master_type_name, copy_type_name, "MASTER WALL TYPE NOT FOUND"]
                skipped_data.append(skipped_type_data)

        mismatched_type_data = [output.linkify(element_id), master_type_name, copy_type_name, "TYPE NAME MISMATCH"]
        mismatched_data.append(mismatched_type_data)

tolerance = 1e-10
difference_start_values = []
for i, (master_start_point, copy_start_point) in enumerate(zip(master_start_points, copy_start_points)):
    coordinate_str = coordinates[0]
    coordinate_coords = list(map(float, coordinate_str.split(',')))
    coordinate = XYZ(coordinate_coords[0], coordinate_coords[1], coordinate_coords[2])
    #coordinate_value = XYZ(coordinate.X, coordinate.Y, coordinate.Z)
    master_start_point = master_start_point.replace('(', '').replace(')', '')
    master_start_coords = list(map(float, master_start_point.split(',')))
    master_value = XYZ(master_start_coords[0], master_start_coords[1], master_start_coords[2])
    #copy_sp_value = XYZ(copy_start_point.X, copy_start_point.Y, copy_start_point.Z)
    copy_start_point = copy_start_point.replace('(', '').replace(')', '')
    copy_start_coords = list(map(float, copy_start_point.split(',')))
    copy_value = XYZ(copy_start_coords[0], copy_start_coords[1], copy_start_coords[2])
    copy_start_value = XYZ((copy_value.X - coordinate.X), (copy_value.Y - coordinate.Y), (copy_value.Z - coordinate.Z))
    if ((master_value.X - copy_start_value.X) > tolerance) or ((master_value.Y - copy_start_value.Y) > tolerance):
        # print("Master Start X : {}, Master Start Y : {}".format(master_value.X, master_value.Y))
        # print("Copy Start Y : {}, Copy Start Y : {}" .format (copy_start_value.X, copy_start_value.Y))
        difference_start = XYZ ((master_value.X - copy_start_value.X), (master_value.Y - copy_start_value.Y), copy_value.Z)
        element_id = ElementId(int(copy_ids[i]))
        element = doc.GetElement (element_id)
        category_name = element.Category.Name
        element_name = element.Name
        if category_name == "Walls":
            wall = element
            wall_curve = wall.Location.Curve
            wall_old_start_point = wall_curve.GetEndPoint(0)
            wall_old_start_point = wall_old_start_point.replace('(', '').replace(')', '') 
            wall_old_start_coords = list(map(float, wall_old_start_point.split(',')))
            wall_old_start = XYZ(wall_old_start_coords[0], wall_old_start_coords[1], wall_old_start_coords[2])
            wall_end_point = wall_curve.GetEndPoint(1)
            wall_end_point = wall_end_point.replace('(', '').replace(')', '') 
            wall_end_coords = list(map(float, wall_end_point.split(',')))
            wall_end = XYZ(wall_end_coords[0], wall_end_coords[1], wall_end_coords[2])
            wall_new_start = XYZ((wall_old_start.X + difference_start.X), (wall_old_start.Y + difference_start.Y), (wall_old_start.Z + difference_start.Z))
            t = Transaction (doc, "Change Wall Type")
            t.Start()
            try: 
                element.WallType = master_wall_type
                t.Commit()
            except Exception as e:
                print ("Wall Type Could not be changed : {}".format(e))

            
        difference_start_values.append(difference_start)
        
        mismatched_start_data = [output.linkify(element_id), master_start_point, copy_start_point, "START POINT MISMATCH"]
        mismatched_data.append(mismatched_start_data)

difference_end_values = []
for i, (master_end_point, copy_end_point) in enumerate(zip(master_end_points, copy_end_points)):
    coordinate_str = coordinates[0]
    coordinate_coords = list(map(float, coordinate_str.split(',')))
    coordinate = XYZ(coordinate_coords[0], coordinate_coords[1], coordinate_coords[2])
    master_end_point = master_end_point.replace('(', '').replace(')', '')
    master_end_coords = list(map(float, master_end_point.split(',')))
    master_value = XYZ(master_end_coords[0], master_end_coords[1], master_end_coords[2])
    copy_end_point = copy_end_point.replace('(', '').replace(')', '')
    copy_end_coords = list(map(float, copy_end_point.split(',')))
    copy_value = XYZ(copy_end_coords[0], copy_end_coords[1], copy_end_coords[2])
    copy_end_value = XYZ((copy_value.X - coordinate.X), (copy_value.Y - coordinate.Y), (copy_value.Z - coordinate.Z))
    if ((master_value.X - copy_end_value.X) > tolerance) or ((master_value.Y - copy_end_value.Y) > tolerance):
        # print("Master End X : {}, Master End Y : {}".format(master_value.X, master_value.Y))
        # print("Copy End X : {}, Copy End Y : {}" .format (copy_end_value.X, copy_end_value.Y))
        difference_end = XYZ ((master_value.X - copy_end_value.X), (master_value.Y - copy_end_value.Y), copy_value.Z)
        difference_end_values.append(difference_end)
        element_id = ElementId(int(copy_ids[i]))
        mismatched_end_data = [output.linkify(element_id), master_end_point, copy_end_point, "END POINT MISMATCH"]
        mismatched_data.append(mismatched_end_data)



# for i, (master_orientation, copy_orientation) in enumerate(zip(master_orientations, copy_orientations)):
#     if master_orientation != copy_orientation:
#         element_id = ElementId(int(copy_ids[i]))
#         mismatched_orientation_data = [output.linkify(element_id), master_orientation, copy_orientation, "ORIENTATION MISMATCH"]
#         mismatched_data.append(mismatched_orientation_data)

# for i, (master_fire_rating, copy_fire_rating) in enumerate(zip(master_fire_ratings, copy_fire_ratings)):
#     if master_fire_rating != copy_fire_rating:
#         element_id = ElementId(int(copy_ids[i]))
#         mismatched_fire_data = [output.linkify(element_id), master_fire_rating, copy_fire_rating, "FIRE RATING MISMATCH"]
#         mismatched_data.append(mismatched_fire_data)

# for i, (master_stc_rating, copy_stc_rating) in enumerate(zip(master_stc_ratings, copy_stc_ratings)):
#     if master_stc_rating != copy_stc_rating:
#         element_id = ElementId(int(copy_ids[i]))
#         mismatched_stc_data = [output.linkify(element_id), master_stc_rating, copy_stc_rating, "STC RATING MISMATCH"]
#         mismatched_data.append(mismatched_stc_data)

# for i, (master_structural, copy_structural) in enumerate(zip(master_structurals, copy_structurals)):
#     if master_structural != copy_structural:
#         element_id = ElementId(int(copy_ids[i]))
#         mismatched_structural_data = [output.linkify(element_id), master_structural, copy_structural, "STRUCTURAL PARAMETER MISMATCH"]
#         mismatched_data.append(mismatched_structural_data)

# if mismatched_data:
#     output.print_md("##⚠️ {} Completed. Issues Found ☹️".format(__title__))
#     output.print_md("---")
#     output.print_md("❌ Some Data is Mismatched Refer to the **Table Report** below for reference")
#     output.print_table(table_data=mismatched_data, columns=["COPIED ELEMENT ID","VALUE IN UNIT TYPOLOGY FILE", "VALUE IN MAIN FILE", "ERROR CODE"])
#     output.print_md("---")
#     output.print_md("***✅ ERROR CODE REFERENCE***")
#     output.print_md("---")
#     output.print_md("**CATEGORY MISMATCH** - Element's Category does not match  \n")
#     output.print_md("**TYPE NAME MISMATCH** - Element's Type Name does not match  \n")
#     output.print_md("**START POINT MISMATCH** - Element's Start Point does not match  \n")
#     output.print_md("**END POINT MISMATCH** - Element's End Point does not match  \n")
#     output.print_md("**ORIENTATION MISMATCH** - Element's Orientation does not match  \n")
#     output.print_md("**FIRE RATING MISMATCH** - Element's Fire Rating does not match  \n")
#     output.print_md("**STC RATING MISMATCH** - Element's STC Rating does not match  \n")
#     output.print_md("**STRUCTURAL PARAMETER MISMATCH** - Element's Structural Parameter does not match  \n")
#     output.print_md("---")

# else:
#     output.print_md("##✅ {} Completed. No Issues Found 😃".format(__title__))
#     output.print_md("---")







