from getpass import getuser
from os import path
from sys import argv
from time import ctime

from arcpy import CheckExtension, CheckOutExtension, Describe, env, Exists, GetParameterAsText, ListDatasets, \
    SetParameterAsText, SetProgressorLabel
from arcpy.analysis import Buffer
from arcpy.management import Clip as Clip_m, Compact, CopyRaster, Delete, MosaicToNewRaster, Project, ProjectRaster
from arcpy.mp import ArcGISProject
from arcpy.da import Editor
from arcpy.sa import ExtractByMask, Hillshade

from hel_utils import AddMsgAndPrint, errorMsg, removeScratchLayers


def logBasicSettings(textFilePath, userWorkspace, inputDEMs, zUnits):
    with open(textFilePath,'a+') as f:
        f.write('\n######################################################################\n')
        f.write('Executing Tool: Prepare Site DEM\n')
        f.write(f"User Name: {getuser()}\n")
        f.write(f"Date Executed: {ctime()}\n")
        f.write('User Parameters:\n')
        f.write(f"\tWorkspace: {userWorkspace}\n")
        f.write(f"\tInput DEMs: {str(inputDEMs)}\n")
        if len (zUnits) > 0:
            f.write(f"\tElevation Z-units: {zUnits}\n")
        else:
            f.write('\tElevation Z-units: NOT SPECIFIED\n')


### Initial Tool Validation ###
try:
    aprx = ArcGISProject('CURRENT')
    map = aprx.listMaps('HEL Determination')[0]
except:
    AddMsgAndPrint('This tool must be run from an ArcGIS Pro project that was developed from the template distributed with this toolbox. Exiting!', 2)
    exit()

if CheckExtension('Spatial') == 'Available':
    CheckOutExtension('Spatial')
else:
    AddMsgAndPrint('Spatial Analyst Extension not enabled. Please enable Spatial Analyst from Project, Licensing, Configure licensing options. Exiting...\n', 2)
    exit()


### ESRI Environment Settings ###
env.overwriteOutput = True
env.resamplingMethod = 'BILINEAR'
env.pyramid = 'PYRAMIDS -1 BILINEAR DEFAULT 75 NO_SKIP'


### Input Parameters ###
sourceCLU = GetParameterAsText(0)
demFormat = GetParameterAsText(1)
inputDEMs = GetParameterAsText(2).split(';')
DEMcount = len(inputDEMs)
sourceService = GetParameterAsText(3)
sourceCellsize = GetParameterAsText(4)
zUnits = GetParameterAsText(5)
demSR = GetParameterAsText(6)
cluSR = GetParameterAsText(7)
transform = GetParameterAsText(8)


try:
    #### Set base path
    sourceCLU_path = Describe(sourceCLU).CatalogPath
    if sourceCLU_path.find('.gdb') > 0 and sourceCLU_path.find('Determinations') > 0 and sourceCLU_path.find('Site_CLU') > 0:
        basedataGDB_path = sourceCLU_path[:sourceCLU_path.find('.gdb')+4]
    else:
        AddMsgAndPrint('\nSelected Site CLU layer is not from a Determinations project folder. Exiting...', 2)
        exit()


    #### Do not run if an unsaved edits exist in the target workspace
    # Pro opens an edit session when any edit has been made and stays open until edits are committed with Save Edits.
    # Check for uncommitted edits and exit if found, giving the user a message directing them to Save or Discard them.
    workspace = basedataGDB_path
    edit = Editor(workspace)
    if edit.isEditing:
        AddMsgAndPrint('\nYou have an active edit session. Please Save or Discard edits and run this tool again. Exiting...', 2)
        exit()


    #### Define Variables
    scratchGDB = path.join(path.dirname(argv[0]), 'SCRATCH.gdb')
    basedataGDB_name = path.basename(basedataGDB_path)
    basedataFD_name = 'Layers'
    basedataFD = path.join(basedataGDB_path, basedataFD_name)
    userWorkspace = path.dirname(basedataGDB_path)
    projectName = path.basename(userWorkspace).replace(' ', '_')
    wetDir = path.join(userWorkspace, 'Wetlands')

    projectTract = path.join(basedataFD, 'Site_Tract')
    projectAOI = path.join(basedataFD, 'Site_AOI')
    projectAOI_B = path.join(basedataFD, 'project_AOI_B')
    projectExtent = path.join(basedataFD, 'Request_Extent')
    bufferDist = '500 Feet'
    bufferDistPlus = '550 Feet'

    projectDEM = path.join(basedataGDB_path, 'Site_DEM')
    projectHillshade = path.join(basedataGDB_path, 'Site_Hillshade')

    wgs_AOI = path.join(scratchGDB, 'AOI_WGS84')
    WGS84_DEM = path.join(scratchGDB, 'WGS84_DEM')
    tempDEM = path.join(scratchGDB, 'tempDEM')
    tempDEM2 = path.join(scratchGDB, 'tempDEM2')
    DEMagg = path.join(scratchGDB, 'aggDEM')
    DEMsmooth = path.join(scratchGDB, 'DEMsmooth')
    ContoursTemp = path.join(scratchGDB, 'ContoursTemp')
    extendedContours = path.join(scratchGDB, 'extendedContours')
    Temp_DEMbase = path.join(scratchGDB, 'Temp_DEMbase')
    Fill_DEMaoi = path.join(scratchGDB, 'Fill_DEMaoi')
    FilMinus = path.join(scratchGDB, 'FilMinus')

    demOut = 'Site_DEM'
    hillshadeOut = 'Site_Hillshade'

    # Temp layers list for cleanup at the start and at the end
    tempLayers = [wgs_AOI, WGS84_DEM, tempDEM, tempDEM2, DEMagg, DEMsmooth, ContoursTemp, extendedContours, Temp_DEMbase, Fill_DEMaoi, FilMinus]
    AddMsgAndPrint('Deleting Temp layers...')
    SetProgressorLabel('Deleting Temp layers...')
    removeScratchLayers(tempLayers)


    #### Set up log file path and start logging
    textFilePath = path.join(userWorkspace, f"{projectName}_log.txt")
    logBasicSettings(textFilePath, userWorkspace, inputDEMs, zUnits)


    #### Create the projectAOI and projectAOI_B layers based on the choice selected by user input
    AddMsgAndPrint('\nBuffering selected extent...',0)
    SetProgressorLabel('Buffering selected extent...')
    Buffer(projectTract, projectAOI, bufferDist, 'FULL', '', 'ALL', '')
    Buffer(projectTract, projectAOI_B, bufferDistPlus, 'FULL', '', 'ALL', '')


    #### Remove existing project DEM related layers from the Pro maps
    AddMsgAndPrint('\nRemoving layers from project maps, if present...\n',0)
    SetProgressorLabel('Removing layers from project maps, if present...')

    # Set starting layers to be removed
    mapLayersToRemove = [demOut, hillshadeOut]

    # Remove the layers in the list
    try:
        for maps in aprx.listMaps():
            for lyr in maps.listLayers():
                if lyr.longName in mapLayersToRemove:
                    maps.removeLayer(lyr)
    except:
        pass


    #### Remove existing DEM related layers from the geodatabase
    AddMsgAndPrint('\nRemoving layers from project database, if present...\n',0)
    SetProgressorLabel('Removing layers from project database, if present...')

    # Set starting datasets to remove
    datasetsToRemove = [projectDEM, projectHillshade]

    # Remove the datasets in the list
    for dataset in datasetsToRemove:
        if Exists(dataset):
            try:
                Delete(dataset)
            except:
                pass


    #### Process the input DEMs
    AddMsgAndPrint('\nProcessing the input DEM(s)...',0)
    SetProgressorLabel('Processing the input DEM(s)...')

    # Extract and process the DEM if it's an image service
    if demFormat == 'NRCS Image Service':
        if sourceCellsize == '':
            AddMsgAndPrint('\nAn output DEM cell size was not specified. Exiting...',2)
            exit()
        else:
            AddMsgAndPrint('\nProjecting AOI to match input DEM...',0)
            SetProgressorLabel('Projecting AOI to match input DEM...')
            wgs_CS = demSR
            Project(projectAOI_B, wgs_AOI, wgs_CS)
            
            AddMsgAndPrint('\nDownloading DEM data...',0)
            SetProgressorLabel('Downloading DEM data...')
            aoi_ext = Describe(wgs_AOI).extent
            xMin = aoi_ext.XMin
            yMin = aoi_ext.YMin
            xMax = aoi_ext.XMax
            yMax = aoi_ext.YMax
            clip_ext = str(xMin) + ' ' + str(yMin) + ' ' + str(xMax) + ' ' + str(yMax)
            Clip_m(sourceService, clip_ext, WGS84_DEM, '', '', '', 'NO_MAINTAIN_EXTENT')

            AddMsgAndPrint('\nProjecting downloaded DEM...',0)
            SetProgressorLabel('Projecting downloaded DEM...')
            ProjectRaster(WGS84_DEM, tempDEM, cluSR, 'BILINEAR', sourceCellsize)

    # Else, extract the local file DEMs
    else:
        # Manage spatial references
        env.outputCoordinateSystem = cluSR
        if transform != '':
            env.geographicTransformations = transform
        
        # Clip out the DEMs that were entered
        AddMsgAndPrint('\tExtracting input DEM(s)...',0)
        SetProgressorLabel('Extracting input DEM(s)...')
        x = 0
        DEMlist = []
        while x < DEMcount:
            raster = inputDEMs[x].replace("'", '')
            desc = Describe(raster)
            raster_path = desc.CatalogPath
            sr = desc.SpatialReference
            units = sr.LinearUnitName
            if units == 'Meter':
                units = 'Meters'
            elif units == 'Foot':
                units = 'Feet'
            elif units == 'Foot_US':
                units = 'Feet'
            else:
                AddMsgAndPrint('\nHorizontal units of one or more input DEMs do not appear to be feet or meters! Exiting...',2)
                exit()
            outClip = tempDEM + '_' + str(x)
            try:
                extractedDEM = ExtractByMask(raster_path, projectAOI_B)
                extractedDEM.save(outClip)
            except:
                AddMsgAndPrint('\nOne or more input DEMs may have a problem! Please verify that the input DEMs cover the tract area and try to run again. Exiting...',2)
                exit()
            if x == 0:
                mosaicInputs = '' + str(outClip) + ''
            else:
                mosaicInputs = mosaicInputs + ';' + str(outClip)
            DEMlist.append(str(outClip))
            x += 1
            del sr

        cellsize = 0
        # Determine largest cell size
        for raster in DEMlist:
            desc = Describe(raster)
            sr = desc.SpatialReference
            cellwidth = desc.MeanCellWidth
            if cellwidth > cellsize:
                cellsize = cellwidth
            del sr

        # Merge the DEMs
        if DEMcount > 1:
            AddMsgAndPrint('\nMerging multiple input DEM(s)...',0)
            SetProgressorLabel('Merging multiple input DEM(s)...')
            MosaicToNewRaster(mosaicInputs, scratchGDB, 'tempDEM', '#', '32_BIT_FLOAT', cellsize, '1', 'MEAN', '#')

        # Else just convert the one input DEM to become the tempDEM
        else:
            AddMsgAndPrint('\nOnly one input DEM detected. Carrying extract forward for final DEM processing...',0)
            firstDEM = DEMlist[0]
            CopyRaster(firstDEM, tempDEM)

        # Delete clippedDEM files
        AddMsgAndPrint('\nDeleting temp DEM file(s)...',0)
        SetProgressorLabel('Deleting temp DEM file(s)...')
        for raster in DEMlist:
            Delete(raster)

        
    # Gather info on the final temp DEM
    desc = Describe(tempDEM)
    sr = desc.SpatialReference
    # linear units should now be meters, since outputs were UTM zone specified
    units = sr.LinearUnitName

    if sr.Type == 'Projected':
        if zUnits == 'Meters':
            Zfactor = 1
        elif zUnits == 'Meter':
            Zfactor = 1
        elif zUnits == 'Feet':
            Zfactor = 0.3048
        elif zUnits == 'Inches':
            Zfactor = 0.0254
        elif zUnits == 'Centimeters':
            Zfactor = 0.01
        else:
            AddMsgAndPrint('\nZunits were not selected at runtime....Exiting!',2)
            exit()

        AddMsgAndPrint('\tDEM Projection Name: ' + sr.Name,0)
        AddMsgAndPrint('\tDEM XY Linear Units: ' + units,0)
        AddMsgAndPrint('\tDEM Elevation Values (Z): ' + zUnits,0)
        AddMsgAndPrint('\tZ-factor for Slope Modeling: ' + str(Zfactor),0)
        AddMsgAndPrint('\tDEM Cell Size: ' + str(desc.MeanCellWidth) + ' x ' + str(desc.MeanCellHeight) + ' ' + units,0)

    else:
        AddMsgAndPrint('\n\n\t' + path.basename(tempDEM) + ' is not in a projected Coordinate System! Exiting...',2)
        exit()

    # Clip out the DEM with extended buffer for temp processing and standard buffer for final DEM display
    AddMsgAndPrint('\nClipping project DEM to buffered extent...',0)
    SetProgressorLabel('Clipping project DEM to buffered extent...')
    Clip_m(tempDEM, '', projectDEM, projectAOI, '', 'ClippingGeometry')


    #### Create Hillshade and Depth Grid
    cZfactor = 0
    if zUnits == 'Meters':
        cZfactor = 3.28084
    elif zUnits == 'Centimeters':
        cZfactor = 0.0328084
    elif zUnits == 'Inches':
        cZfactor = 0.0833333
    else:
        cZfactor = 1
    AddMsgAndPrint('\nCreating Hillshade...',0)
    SetProgressorLabel('Creating Hillshade...')
    outHillshade = Hillshade(projectDEM, '315', '45', '#', Zfactor)
    outHillshade.save(projectHillshade)
    AddMsgAndPrint('\tSuccessful',0)


    #### Delete temp data
    AddMsgAndPrint('\nDeleting temp data...' ,0)
    SetProgressorLabel('Deleting temp data...')
    removeScratchLayers(tempLayers)


    #### Add layers to Pro Map
    AddMsgAndPrint('\nAdding layers to map...',0)
    SetProgressorLabel('Adding layers to map...')
    SetParameterAsText(9, projectDEM)
    SetParameterAsText(10, projectHillshade)


    #### Clean up
    # Look for and delete anything else that may remain in the installed SCRATCH.gdb
    startWorkspace = env.workspace
    env.workspace = scratchGDB
    dss = []
    for ds in ListDatasets('*'):
        dss.append(path.join(scratchGDB, ds))
    for ds in dss:
        if Exists(ds):
            try:
                Delete(ds)
            except:
                pass
    env.workspace = startWorkspace
    del startWorkspace


    #### Compact FGDB
    try:
        AddMsgAndPrint('\nCompacting File Geodatabase...' ,0)
        SetProgressorLabel('Compacting File Geodatabase...')
        Compact(basedataGDB_path)
        AddMsgAndPrint('\tSuccessful',0)
    except:
        pass


except SystemExit:
    pass

except:
    errorMsg('Prepare Site DEM')
