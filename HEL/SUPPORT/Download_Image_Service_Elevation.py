from os import path
from pathlib import Path
from sys import argv

from arcpy import AddError, AddMessage, Describe, env, Exists, GetParameterAsText, SetParameterAsText, SpatialReference
from arcpy.analysis import Buffer
from arcpy.management import Clip, CreateFileGDB, Delete, Project, ProjectRaster


# Tool Inputs
source_clu = GetParameterAsText(0)
source_service = GetParameterAsText(1)

# Paths to SCRATCH.gdb features
scratch_gdb = path.join(path.dirname(argv[0]), 'SCRATCH.gdb')
clu_buffer = path.join(scratch_gdb, 'CLU_Buffer')
wgs84_clu_buffer = path.join(scratch_gdb, 'WGS84_CLU_Buffer')
wgs84_DEM = path.join(scratch_gdb, 'WGS84_DEM')

# Output to project base data GDB
base_data_gdb = Path(Describe(source_clu).catalogPath).parent.parent
final_DEM = path.join(base_data_gdb, 'Downloaded_DEM')

# Project Base Data GDB validation
if not Exists(base_data_gdb):
    AddError('\Failed to locate the project Base Data GDB... Exiting')
    exit()

# Create SCRATCH.gdb if needed, clear any existing features otherwise
if not Exists(scratch_gdb):
    try:
        CreateFileGDB(path.dirname(argv[0]), 'SCRATCH.gdb')
    except:
        AddError('Failed to create SCRATCH.gdb in install location... Exiting')
        exit()
else:
    scratch_features = [clu_buffer, wgs84_clu_buffer, wgs84_DEM]
    for feature in scratch_features:
        if Exists(feature):
            Delete(feature)

# Geoprocessing Environment Settings
env.workspace = scratch_gdb
env.overwriteOutput = True
env.geographicTransformations = 'WGS_1984_(ITRF00)_To_NAD_1983'
env.resamplingMethod = 'BILINEAR'
env.pyramid = 'PYRAMIDS -1 BILINEAR DEFAULT 75 NO_SKIP'

# Buffer the selected CLUs
AddMessage('Buffering input CLU fields...')
try:
    Buffer(source_clu, clu_buffer, '410 Meters', 'FULL', '', 'ALL', '')
except:
    AddError('Failed to buffer selected CLU(s)... Exiting')
    exit()

# Re-project the CLU Buffer to WGS84 Geographic (EPSG WKID: 4326)
AddMessage('Projecting CLU Buffer to WGS 1984...')
try:
    Project(clu_buffer, wgs84_clu_buffer, SpatialReference(4326))
except:
    AddError('Failed to project CLU Buffer to WGS84... Exiting')
    exit()

# Use the WGS84 CLU Buffer to clip DEM service
AddMessage('Clipping DEM service to area of interest...')
try:
    aoi_ext = Describe(wgs84_clu_buffer).extent
    clip_ext = f"{str(aoi_ext.XMin)} {str(aoi_ext.YMin)} {str(aoi_ext.XMax)} {str(aoi_ext.YMax)}"
    Clip(source_service, clip_ext, wgs84_DEM, '', '', '', 'NO_MAINTAIN_EXTENT')
except:
    AddError('Failed to clip DEM service to area of interest... Exiting')
    exit()

# Project the WGS84 DEM to the coordinate system of the input CLU layer
AddMessage('Projecting DEM to match input CLU...')
try:
    final_CS = Describe(source_clu).spatialReference.factoryCode
    ProjectRaster(wgs84_DEM, final_DEM, final_CS, 'BILINEAR', 3)
except:
    AddError('Failed to project clipped DEM to CLU coordinate system... Exiting')
    exit()

# Clear SCRATCH.gdb except final DEM
scratch_features = [clu_buffer, wgs84_clu_buffer, wgs84_DEM]
for feature in scratch_features:
    if Exists(feature):
        Delete(feature)

# Add final DEM layer to derived output parameter
AddMessage('Adding DEM layer to map...')
SetParameterAsText(2, final_DEM)
