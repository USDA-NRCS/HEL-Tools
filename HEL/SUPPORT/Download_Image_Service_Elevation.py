from arcpy import AddError, AddMessage, CheckExtension, CheckOutExtension, Describe, env, Exists, \
    GetParameterAsText, SetParameterAsText, SpatialReference
from arcpy.analysis import Buffer
from arcpy.management import Clip, Delete, Project, ProjectRaster

from os import path
from sys import argv, exit


# Check out Spatial Analyst License
if CheckExtension('Spatial') == 'Available':
    CheckOutExtension('Spatial')
else:
    AddError('Spatial Analyst Extension must be enabled to use this tool... Exiting')
    exit()

# Geoprocessing Environment Settings
scratch_gdb = path.join(path.dirname(argv[0]), 'scratch.gdb')
env.scratchWorkspace = scratch_gdb
env.overwriteOutput = True
env.geographicTransformations = 'WGS_1984_(ITRF00)_To_NAD_1983'
env.resamplingMethod = 'BILINEAR'
env.pyramid = 'PYRAMIDS -1 BILINEAR DEFAULT 75 NO_SKIP'

# Tool Inputs
source_clu = GetParameterAsText(0)
source_service = GetParameterAsText(1)

# Variables
wgs84_DEM = path.join(scratch_gdb, 'WGS84_DEM')
final_DEM = path.join(scratch_gdb, 'Downloaded_DEM')
clu_buffer = path.join(scratch_gdb, 'CLU_Buffer')
clu_buff_wgs = path.join(scratch_gdb, 'CLU_Buffer_WGS')

# Delete posible temp datasets if they already exist
if Exists(wgs84_DEM):
    Delete(wgs84_DEM)
if Exists(final_DEM):
    Delete(final_DEM)
if Exists(clu_buff_wgs):
    Delete(clu_buff_wgs)

# Buffer the selected CLUs by 400m
AddMessage('Buffering input CLU fields...')
# Use 410 meter radius so that you have a bit of extra area for the HEL Determination tool to clip against
# This distance is designed to minimize problems of no data crashes if the HEL Determiation tool's resampled 3-meter DEM doesn't perfectly snap with results from this tool.
Buffer(source_clu, clu_buffer, '410 Meters', 'FULL', '', 'ALL', '')

# Re-project the AOI to WGS84 Geographic (EPSG WKID: 4326)
AddMessage('Converting CLU Buffer to WGS 1984...')
Project(clu_buffer, clu_buff_wgs, SpatialReference(4326))

# Use the WGS 1984 AOI to clip/extract the DEM from the service
AddMessage('Downloading Data...')
aoi_ext = Describe(clu_buff_wgs).extent
clip_ext = f"{str(aoi_ext.XMin)} {str(aoi_ext.YMin)} {str(aoi_ext.XMax)} {str(aoi_ext.YMax)}"
Clip(source_service, clip_ext, wgs84_DEM, '', '', '', 'NO_MAINTAIN_EXTENT')

# Project the WGS 1984 DEM to the coordinate system of the input CLU layer
AddMessage('Projecting data to match input CLU...')
final_CS = Describe(source_clu).spatialReference.factoryCode
ProjectRaster(wgs84_DEM, final_DEM, final_CS, 'BILINEAR', 3)

# Delete temporary data
AddMessage('Cleaning up...')
Delete(wgs84_DEM)
Delete(clu_buff_wgs)
Delete(clu_buffer)

# Add resulting data to map
AddMessage('Adding layer to map...')
SetParameterAsText(2, final_DEM)
