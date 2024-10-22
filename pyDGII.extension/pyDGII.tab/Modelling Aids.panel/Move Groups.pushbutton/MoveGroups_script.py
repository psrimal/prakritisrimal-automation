# -*- coding: utf-8 -*-
'''Move Groups to internal origin'''

__title__ = "Move Groups"
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
openings = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_SWallRectOpening).WhereElementIsNotElementType().ToElements()
levels = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Levels).WhereElementIsNotElementType().ToElements()
for level in levels:
    level_elevation = level.Elevation
    sorted_levels = sorted(levels, key=lambda lvl: lvl.Elevation)
    lowest_level = sorted_levels[0].Id
t = Transaction (doc, "Move Group")
t.Start()
for group in model_groups:
    group_id = group.Id
    location = group.Location
    location_point = location.Point
    translation_vector = XYZ (0,0,0) - location_point
    ElementTransformUtils.MoveElement(doc, group_id, translation_vector)
    print ("Model group {} moved from {} to {}". format(group.Name, location_point, translation_vector))
    current_level_param = group.LookupParameter("Reference Level")
    if current_level_param:
        current_level_param.Set(lowest_level)
    level_offset = group.LookupParameter("Origin Level Offset")
    if level_offset:
        level_offset.Set(0)

t.Commit()
