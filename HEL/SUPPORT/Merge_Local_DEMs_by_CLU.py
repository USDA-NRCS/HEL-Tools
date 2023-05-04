from arcpy import AddError, AddMessage, CheckExtension, CheckOutExtension, CreateScratchName, \
    Describe, env, Exists, GetParameter, GetParameterAsText, SetParameterAsText
from arcpy.analysis import Buffer
from arcpy.management import CopyFeatures, Delete, GetRasterProperties, MosaicToNewRaster
from arcpy.sa import ExtractByMask

from os import path
from sys import argv, exit


# Check out Spatial Analyst License
if CheckExtension('Spatial') == 'Available':
    CheckOutExtension('Spatial')
else:
    AddError('Spatial Analyst Extension must be enabled to use this tool... Exiting')
    exit()

# Tool Inputs
source_clu = GetParameter(0)
source_dems = GetParameterAsText(1).split(';')

# Geoprocessing Environment Settings
scratch_gdb = path.join(path.dirname(argv[0]), 'SCRATCH.gdb')
env.workspace = scratch_gdb
env.overwriteOutput = True
env.geographicTransformations = 'WGS_1984_(ITRF00)_To_NAD_1983'
env.resamplingMethod = 'BILINEAR'
env.pyramid = 'PYRAMIDS -1 BILINEAR DEFAULT 75 NO_SKIP'
env.outputCoordinateSystem = Describe(source_clu).SpatialReference

# Make sure at least 2 datasets to be merged were entered
datasets = len(source_dems)
if datasets < 2:
    AddError('Only one input DEM layer selected. If you need multiple layers, please run again and select multiple DEM files... Exiting')
    exit()

temp_dem = CreateScratchName('temp_dem', data_type='RasterDataset')
merged_dem = CreateScratchName('merged_dem', data_type='RasterDataset')
clu_selected = path.join('in_memory', path.basename(CreateScratchName('clu_selected', data_type='FeatureClass')))
clu_buffer = path.join('in_memory', path.basename(CreateScratchName('clu_buffer', data_type='FeatureClass')))

# Make sure CLU fields are selected
cluDesc = Describe(source_clu)
if cluDesc.FIDset == '':
    AddError('Please select fields from the CLU Layer. Exiting')
    exit()
else:
    source_clu = CopyFeatures(source_clu, clu_selected)

AddMessage(f"Number of CLU fields selected: {str(len(cluDesc.FIDset.split(';')))}")

# Dictionary for code values returned by the getRasterProperties tool (keys)
# Values represent the pixel_type inputs for the mosaic to new raster tool
pixelTypeDict = {0:'1_BIT',1:'2_BIT',2:'4_BIT',3:'8_BIT_UNSIGNED',4:'8_BIT_SIGNED',5:'16_BIT_UNSIGNED',6:'16_BIT_SIGNED',7:'32_BIT_UNSIGNED',8:'32_BIT_SIGNED',9:'32_BIT_FLOAT',10:'64_BIT'}
numOfBands = ''
pixelType = ''

# Evaluate every input raster to be merged
AddMessage(f"Checking {str(datasets)} input raster layers")
x = 0
while x < datasets:
    raster = source_dems[x].replace("'", '')
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
    cellValueCode =  int(GetRasterProperties(raster, 'VALUETYPE').getOutput(0))
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

# Buffer the selected CLUs by 400m
AddMessage('Buffering input CLU fields...')
# Use 410 meter radius so that you have a bit of extra area for the HEL Determination tool to clip against
# This distance is designed to minimize problems of no data crashes if the HEL Determiation tool's resampled 3-meter DEM doesn't perfectly snap with results from this tool.
Buffer(source_clu, clu_buffer, '410 Meters', 'FULL', '', 'ALL', '')

# Clip out the DEMs that were entered
AddMessage('Clipping Raster Layers...')
del_list = [] # Start an empty list that will be used to clean up the temporary clips after merge is done
mergeRasters = ''
x = 0
while x < datasets:
    current_dem = source_dems[x].replace("'", '')
    out_clip = f"{temp_dem}_{str(x)}"

    try:
        AddMessage(f"Clipping {current_dem} {str(x+1)} of {str(datasets)}")
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
AddMessage('Merging inputs...')

if Exists(merged_dem):
    Delete(merged_dem)

MosaicToNewRaster(mergeRasters, scratch_gdb, path.basename(merged_dem), '#', pixelType, 3, numOfBands, 'MEAN', '#')

for lyr in del_list:
    Delete(lyr)
Delete(clu_buffer)

# Add resulting data to map
AddMessage(f"Adding {path.basename(merged_dem)} to map...")
SetParameterAsText(2, merged_dem)
