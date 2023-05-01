from os import path
from sys import argv, exit
import arcpy

# Set arcpy environments
scratch_gdb = path.join(path.dirname(argv[0]), 'scratch.gdb')
arcpy.env.scratchWorkspace = scratch_gdb
arcpy.env.overwriteOutput = True
arcpy.env.geographicTransformations = 'WGS_1984_(ITRF00)_To_NAD_1983'
arcpy.env.resamplingMethod = 'BILINEAR'
arcpy.env.pyramid = 'PYRAMIDS -1 BILINEAR DEFAULT 75 NO_SKIP'

# Check out Spatial Analyst License
if arcpy.CheckExtension('Spatial') == 'Available':
    arcpy.CheckOutExtension('Spatial')
else:
    arcpy.AddError('Spatial Analyst Extension not enabled. Please enable Spatial analyst from the Tools/Extensions menu... Exiting')
    exit()

# Inputs
source_clu = arcpy.GetParameterAsText(0)
source_service = arcpy.GetParameterAsText(1)

# Variables
wgs84_DEM = path.join(scratch_gdb, 'WGS84_DEM')
final_DEM = path.join(scratch_gdb, 'Downloaded_DEM')
clu_buffer = path.join(scratch_gdb, 'CLU_Buffer')
clu_buff_wgs = path.join(scratch_gdb, 'CLU_Buffer_WGS')

# Delete posible temp datasets if they already exist
if arcpy.Exists(wgs84_DEM):
    arcpy.Delete_management(wgs84_DEM)
if arcpy.Exists(final_DEM):
    arcpy.Delete_management(final_DEM)
if arcpy.Exists(clu_buff_wgs):
    arcpy.Delete_management(clu_buff_wgs)

# Buffer the selected CLUs by 400m
arcpy.AddMessage('Buffering input CLU fields...')
# Use 410 meter radius so that you have a bit of extra area for the HEL Determination tool to clip against
# This distance is designed to minimize problems of no data crashes if the HEL Determiation tool's resampled 3-meter DEM doesn't perfectly snap with results from this tool.
arcpy.Buffer_analysis(source_clu, clu_buffer, '410 Meters', 'FULL', '', 'ALL', '')

# Re-project the AOI to WGS84 Geographic (EPSG WKID: 4326)
arcpy.AddMessage('Converting CLU Buffer to WGS 1984...')
wgs_CS = arcpy.SpatialReference(4326)
arcpy.Project_management(clu_buffer, clu_buff_wgs, wgs_CS)

# Use the WGS 1984 AOI to clip/extract the DEM from the service
arcpy.AddMessage('Downloading Data...')
aoi_ext = arcpy.Describe(clu_buff_wgs).extent
clip_ext = f"{str(aoi_ext.XMin)} {str(aoi_ext.YMin)} {str(aoi_ext.XMax)} {str(aoi_ext.YMax)}"
arcpy.Clip_management(source_service, clip_ext, wgs84_DEM, '', '', '', 'NO_MAINTAIN_EXTENT')

# Project the WGS 1984 DEM to the coordinate system of the input CLU layer
arcpy.AddMessage('Projecting data to match input CLU...')
final_CS = arcpy.Describe(source_clu).spatialReference.factoryCode
arcpy.ProjectRaster_management(wgs84_DEM, final_DEM, final_CS, 'BILINEAR', 3)

# Delete temporary data
arcpy.AddMessage('Cleaning up...')
arcpy.Delete_management(wgs84_DEM)
arcpy.Delete_management(clu_buff_wgs)
arcpy.Delete_management(clu_buffer)

# Add resulting data to map
arcpy.AddMessage('Adding layer to map...')
arcpy.SetParameterAsText(2, final_DEM)
