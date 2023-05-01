from os import path
from sys import argv, exit
import arcpy

# Set overwrite
arcpy.env.overwriteOutput = True
        
# Inputs
source_clu = arcpy.GetParameterAsText(0)
source_soils = arcpy.GetParameterAsText(1).split(';')

# Variables
temp_soil = path.join(path.dirname(argv[0]), 'HEL.mdb', 'temp_soil')
merged_soil = path.join(path.dirname(argv[0]), 'HEL.mdb', 'Merged_HEL_Soil')

# Make sure at least 2 datasets to be merged were entered
if len(source_soils) < 2:
    arcpy.AddError('Only one input soil layer selected. If you need multiple layers, please run again and select multiple soil layers... Exiting')
    exit()

# List of valid HEL soil layer schema for the tool (in lower case for comparative purposes)
schema = ['areasymbol', 'spatialver', 'musym', 'muname', 'muhelcl', 't', 'k', 'r']
# Check the input soil layers to make sure they contain fields with the same field names
x = 0
for layer in source_soils:
    field_names = [f.name.lower() for f in arcpy.ListFields(source_soils[x])]
    for s in schema:
        if s not in field_names:
            arcpy.AddMessage(f"The layer {str(source_soils[x])} is missing field {str(s)}... Exiting")
            exit()
    x += 1

# Clip out the soils that were entered
arcpy.AddMessage('Clipping inputs...')
x = 0
del_list = []
# Start an empty list that will be used to clean up the temporary clips after merge is done
while x < len(source_soils):
    current_soil = source_soils[x].replace("'", '')
    out_clip = f"{temp_soil}_{str(x)}"
    try:
        arcpy.Clip_analysis(current_soil, source_clu, out_clip)
    except:
        arcpy.AddError('The input fields may not cover the input soil layers. Clip & Merge failed... Exiting')
        exit()
    if x == 0:
        # Start list of layers to merge
        merge_list = str(out_clip)
    else:
        # Append to list
        merge_list = f"{merge_list};{str(out_clip)}"
    # Append name of temporary output to the list of temp soil layers to be deleted
    del_list.append(str(out_clip))
    x += 1

# Merge Clipped Datasets
arcpy.AddMessage('Merging inputs...')
if arcpy.Exists(merged_soil):
    arcpy.Delete_management(merged_soil)
arcpy.Merge_management(merge_list, merged_soil, '')

# Delete temporary soils
arcpy.AddMessage('Cleaning up...')
for lyr in del_list:
    arcpy.Delete_management(lyr)

# Add resulting data to map
arcpy.AddMessage('Adding layer to map...')
arcpy.SetParameterAsText(2, merged_soil)
