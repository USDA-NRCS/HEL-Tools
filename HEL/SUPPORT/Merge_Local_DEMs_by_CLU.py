from os import path
from sys import argv, exit

from arcpy import AddError, AddMessage, Describe, env, Exists, GetParameterAsText, SetParameterAsText
from arcpy.analysis import Buffer
from arcpy.management import CreateFileGDB, Delete, GetRasterProperties, MosaicToNewRaster
from arcpy.sa import ExtractByMask


# Tool Inputs
source_clu = GetParameterAsText(0)
source_DEMs = GetParameterAsText(1).split(';')
number_of_DEMs = len(source_DEMs)

# Paths to SCRATCH.gdb features
scratch_gdb = path.join(path.dirname(argv[0]), 'SCRATCH.gdb')
clu_buffer = path.join(scratch_gdb, 'CLU_Buffer')
merged_DEM = path.join(scratch_gdb, 'Merged_DEM')

# Create SCRATCH.gdb if needed, clear any existing features otherwise
if not Exists(scratch_gdb):
    try:
        CreateFileGDB(path.dirname(argv[0]), 'SCRATCH.gdb')
    except Exception:
        AddError('Failed to create SCRATCH.gdb in install location... Exiting')
        exit()
else:
    scratch_features = [clu_buffer, merged_DEM]
    for feature in scratch_features:
        if Exists(feature):
            Delete(feature)

# Geoprocessing Environment Settings
env.workspace = scratch_gdb
env.overwriteOutput = True
env.geographicTransformations = 'WGS_1984_(ITRF00)_To_NAD_1983'
env.resamplingMethod = 'BILINEAR'
env.pyramid = 'PYRAMIDS -1 BILINEAR DEFAULT 75 NO_SKIP'
env.outputCoordinateSystem = Describe(source_clu).SpatialReference

# Dictionary for code values returned by the getRasterProperties tool (keys)
# Values represent the pixel_type inputs for the mosaic to new raster tool
pixelTypeDict = {0:'1_BIT',1:'2_BIT',2:'4_BIT',3:'8_BIT_UNSIGNED',4:'8_BIT_SIGNED',5:'16_BIT_UNSIGNED',6:'16_BIT_SIGNED',7:'32_BIT_UNSIGNED',8:'32_BIT_SIGNED',9:'32_BIT_FLOAT',10:'64_BIT'}
numOfBands = ''
pixelType = ''

# Evaluate every input raster to be merged
AddMessage(f"Checking {str(number_of_DEMs)} input raster layers")
x = 0
while x < number_of_DEMs:
    raster = source_DEMs[x].replace("'", '')
    desc = Describe(raster)
    sr = desc.SpatialReference
    units = sr.LinearUnitName
    bandCount = desc.bandCount

    # Check for Projected Coordinate System
    if sr.Type != 'Projected':
        AddError(f"The {str(raster)} input must have a projected coordinate system... Exiting")
        exit()

    # Check for linear Units
    if units == 'Meter':
        tolerance = 3
    elif units == 'Foot':
        tolerance = 9.84252
    elif units == 'Foot_US':
        tolerance = 9.84252
    else:
        AddError(f"Horizontal units of {str(desc.baseName)} must be in feet or meters... Exiting")
        exit()

    # Check for cell size; Reject if greater than 3m
    if desc.MeanCellWidth > tolerance:
        AddError(f"The cell size of the {str(raster)} input exceeds 3 meters or 9.84252 feet which cannot be used in the NRCS HEL Determination Tool... Exiting")
        exit()

    # Check for consistent bit depth
    cellValueCode = int(GetRasterProperties(raster, 'VALUETYPE').getOutput(0))
    bitDepth = pixelTypeDict[cellValueCode]  # Convert the code to pixel depth using dictionary

    if pixelType == '':
        pixelType = bitDepth
    else:
        if pixelType != bitDepth:
            AddError(f"Cannot Mosaic different pixel types: {bitDepth} & {pixelType}")
            AddError('Pixel Types must be the same for all input rasters...')
            AddError('Contact your state GIS Coordinator to resolve this issue... Exiting')
            exit()

    # Check for consistent band count - highly unlikely more than 1 band is input
    if numOfBands == '':
        numOfBands = bandCount
    else:
        if numOfBands != bandCount:
            AddError(f"Cannot mosaic rasters with multiple raster bands: {str(numOfBands)} & {str(bandCount)}")
            AddError('Number of bands must be the same for all input rasters...')
            AddError('Contact your state GIS Coordinator to resolve this issue... Exiting')
            exit()
    x += 1

# Buffer the selected CLUs
AddMessage('Buffering input CLU fields...')
Buffer(source_clu, clu_buffer, '410 Meters', 'FULL', '', 'ALL', '')

# Clip out the DEMs that were entered
AddMessage('Clipping Raster Layers...')
del_list = [] # Start an empty list that will be used to clean up the temporary clips after merge is done
mergeRasters = ''
x = 0
while x < number_of_DEMs:
    current_dem = source_DEMs[x].replace("'", '')
    out_clip = f"temp_dem_{str(x)}"

    try:
        AddMessage(f"Clipping {current_dem} {str(x+1)} of {str(number_of_DEMs)}")
        extractedDEM = ExtractByMask(current_dem, clu_buffer)
        extractedDEM.save(out_clip)
    except:
        AddError('The input CLU fields may not cover the input DEM files. Clip & Merge failed... Exiting')
        exit()

    # Create merge statement
    if x == 0:
        # Start list of layers to merge
        mergeRasters = str(out_clip)
    else:
        # Append to list
        mergeRasters = f"{mergeRasters};{str(out_clip)}"

    # Append name of temporary output to the list of temp soil layers to be deleted
    del_list.append(str(out_clip))
    x += 1

# Merge Clipped Datasets
AddMessage('Merging clipped DEMs...')
MosaicToNewRaster(mergeRasters, scratch_gdb, path.basename(merged_DEM), '#', pixelType, 3, numOfBands, 'MEAN', '#')

for lyr in del_list:
    Delete(lyr)
Delete(clu_buffer)

# Add final DEM layer to derived output parameter
AddMessage('Adding DEM layer to map...')
SetParameterAsText(2, merged_DEM)
