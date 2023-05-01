from os import path
from sys import argv, exit
import arcpy

# Inputs
source_clu = arcpy.GetParameter(0)
source_dems = arcpy.GetParameterAsText(1).split(';')

# Set environmental variables
scratch_gdb = path.join(path.dirname(argv[0]), 'scratch.gdb')
arcpy.env.workspace = scratch_gdb
arcpy.env.overwriteOutput = True
arcpy.env.geographicTransformations = 'WGS_1984_(ITRF00)_To_NAD_1983'
arcpy.env.resamplingMethod = 'BILINEAR'
arcpy.env.pyramid = 'PYRAMIDS -1 BILINEAR DEFAULT 75 NO_SKIP'
arcpy.env.outputCoordinateSystem = arcpy.Describe(source_clu).SpatialReference

# Check out Spatial Analyst License
if arcpy.CheckExtension('Spatial') == 'Available':
    arcpy.CheckOutExtension('Spatial')
else:
    arcpy.AddError('Spatial Analyst Extension not enabled. Please enable Spatial analyst from the Tools/Extensions menu... Exiting')
    exit()

# Make sure at least 2 datasets to be merged were entered
datasets = len(source_dems)
if datasets < 2:
    arcpy.AddError('Only one input DEM layer selected. If you need multiple layers, please run again and select multiple DEM files... Exiting')
    exit()

temp_dem = arcpy.CreateScratchName('temp_dem', data_type='RasterDataset')
merged_dem = arcpy.CreateScratchName('merged_dem', data_type='RasterDataset')
clu_selected = path.join('in_memory', path.basename(arcpy.CreateScratchName('clu_selected', data_type='FeatureClass')))
clu_buffer = path.join('in_memory', path.basename(arcpy.CreateScratchName('clu_buffer', data_type='FeatureClass')))

# Make sure CLU fields are selected
cluDesc = arcpy.Describe(source_clu)
if cluDesc.FIDset == '':
    arcpy.AddError('Please select fields from the CLU Layer. Exiting')
    exit()
else:
    source_clu = arcpy.CopyFeatures_management(source_clu, clu_selected)

arcpy.AddMessage(f"Number of CLU fields selected: {str(len(cluDesc.FIDset.split(';')))}")

# Dictionary for code values returned by the getRasterProperties tool (keys)
# Values represent the pixel_type inputs for the mosaic to new raster tool
pixelTypeDict = {0:'1_BIT',1:'2_BIT',2:'4_BIT',3:'8_BIT_UNSIGNED',4:'8_BIT_SIGNED',5:'16_BIT_UNSIGNED',6:'16_BIT_SIGNED',7:'32_BIT_UNSIGNED',8:'32_BIT_SIGNED',9:'32_BIT_FLOAT',10:'64_BIT'}
numOfBands = ''
pixelType = ''

# Evaluate every input raster to be merged
arcpy.AddMessage(f"Checking {str(datasets)} input raster layers")
x = 0
while x < datasets:
    raster = source_dems[x].replace("'", '')
    desc = arcpy.Describe(raster)
    sr = desc.SpatialReference
    units = sr.LinearUnitName
    bandCount = desc.bandCount

    # Check for Projected Coordinate System
    if sr.Type != 'Projected':
        arcpy.AddError(f"The {str(raster)} input must have a projected coordinate system... Exiting")
        exit()

    # Check for linear Units
    if units == 'Meter':
        tolerance = 3
    elif units == 'Foot':
        tolerance = 9.84252
    elif units == 'Foot_US':
        tolerance = 9.84252
    else:
        arcpy.AddError(f"Horizontal units of {str(desc.baseName)} must be in feet or meters... Exiting")
        exit()

    # Check for cell size; Reject if greater than 3m
    if desc.MeanCellWidth > tolerance:
        arcpy.AddError(f"The cell size of the {str(raster)} input exceeds 3 meters or 9.84252 feet which cannot be used in the NRCS HEL Determination Tool... Exiting")
        exit()

    # Check for consistent bit depth
    cellValueCode =  int(arcpy.GetRasterProperties_management(raster, 'VALUETYPE').getOutput(0))
    bitDepth = pixelTypeDict[cellValueCode]  # Convert the code to pixel depth using dictionary

    if pixelType == '':
        pixelType = bitDepth
    else:
        if pixelType != bitDepth:
            arcpy.AddError(f"Cannot Mosaic different pixel types: {bitDepth} & {pixelType}")
            arcpy.AddError('Pixel Types must be the same for all input rasters...')
            arcpy.AddError('Contact your state GIS Coordinator to resolve this issue... Exiting')
            exit()

    # Check for consistent band count - highly unlikely more than 1 band is input
    if numOfBands == '':
        numOfBands = bandCount
    else:
        if numOfBands != bandCount:
            arcpy.AddError(f"Cannot mosaic rasters with multiple raster bands: {str(numOfBands)} & {str(bandCount)}")
            arcpy.AddError('Number of bands must be the same for all input rasters...')
            arcpy.AddError('Contact your state GIS Coordinator to resolve this issue... Exiting')
            exit()
    x += 1

# Buffer the selected CLUs by 400m
arcpy.AddMessage('Buffering input CLU fields...')
# Use 410 meter radius so that you have a bit of extra area for the HEL Determination tool to clip against
# This distance is designed to minimize problems of no data crashes if the HEL Determiation tool's resampled 3-meter DEM doesn't perfectly snap with results from this tool.
arcpy.Buffer_analysis(source_clu, clu_buffer, '410 Meters', 'FULL', '', 'ALL', '')

# Clip out the DEMs that were entered
arcpy.AddMessage('Clipping Raster Layers...')
del_list = [] # Start an empty list that will be used to clean up the temporary clips after merge is done
mergeRasters = ''
x = 0
while x < datasets:
    current_dem = source_dems[x].replace("'", '')
    out_clip = f"{temp_dem}_{str(x)}"

    try:
        arcpy.AddMessage(f"Clipping {current_dem} {str(x+1)} of {str(datasets)}")
        extractedDEM = arcpy.sa.ExtractByMask(current_dem, clu_buffer)
        extractedDEM.save(out_clip)
    except:
        arcpy.AddError('The input CLU fields may not cover the input DEM files. Clip & Merge failed... Exiting')
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
arcpy.AddMessage('Merging inputs...')

if arcpy.Exists(merged_dem):
    arcpy.Delete_management(merged_dem)

arcpy.MosaicToNewRaster_management(mergeRasters, scratch_gdb, path.basename(merged_dem), '#', pixelType, 3, numOfBands, 'MEAN', '#')

for lyr in del_list:
    arcpy.Delete_management(lyr)
arcpy.Delete_management(clu_buffer)

# Add resulting data to map
arcpy.AddMessage(f"Adding {path.basename(merged_dem)} to map...")
arcpy.SetParameterAsText(2, merged_dem)
