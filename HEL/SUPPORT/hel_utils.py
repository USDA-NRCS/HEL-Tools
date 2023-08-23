from datetime import date
from getpass import getuser
from json import dump as json_dump, loads as json_loads
from os import makedirs, path, startfile
from sys import argv, exc_info
from time import ctime
from traceback import format_exception
from urllib import parse, request

from arcpy import AddError, AddFieldDelimiters, AddMessage, AddWarning, CreateScratchName, Describe, env, Exists, \
    ListFields, ParseFieldName, SetParameterAsText, SetProgressor, SetProgressorLabel, SpatialReference

from arcpy.analysis import Buffer
from arcpy.conversion import JSONToFeatures
from arcpy.da import SearchCursor, UpdateCursor

from arcpy.management import AddField, CalculateField, Clip as Clip_m, Delete, Dissolve, MakeRasterLayer, Project, \
    ProjectRaster, SelectLayerByAttribute

# TODO: All arcpy.mapping must be converted to arcpy.mp (MapDocument --> ArcGISProject)
# from arcpy.mapping import AddLayer, Layer, ListDataFrames, ListLayoutElements, ListLayers, MapDocument, RefreshActiveView, \
#     RefreshTOC, RemoveLayer, UpdateLayer
# from arcpy.mp import ArcGISProject, LayerFile


def AddMsgAndPrint(msg, severity=0, textFilePath=None):
    """ Adds tool message to the geoprocessor. Split the message on \n first, so a GPMessage will be added for each line."""
    try:
        if textFilePath:
            f = open(textFilePath, 'a+')
            f.write(f"{msg}\n")
            f.close
            del f
        if severity == 0:
            AddMessage(msg)
        elif severity == 1:
            AddWarning(msg)
        elif severity == 2:
            AddError(msg)
    except:
        pass


def errorMsg():
    """ Print traceback exceptions. If sys.exit was trapped by default exception then ignore traceback message."""
    try:
        exc_type, exc_value, exc_traceback = exc_info()
        theMsg = f"\t{format_exception(exc_type, exc_value, exc_traceback)[1]}\n\t{format_exception(exc_type, exc_value, exc_traceback)[-1]}"
        if theMsg.find('sys.exit') > -1:
            AddMsgAndPrint('\n\n')
            pass
        else:
            AddMsgAndPrint('\n\tNRCS HEL Tool Error: -------------------------', 2)
            AddMsgAndPrint(theMsg, 2)
    except:
        AddMsgAndPrint('Unhandled error in errorMsg method', 2)
        pass


def logBasicSettings(textFilePath, sourceState, sourceCounty, tractNumber, owFlag):
    with open(textFilePath, 'a+') as f:
        f.write('\n######################################################################\n')
        f.write('Executing Create Project Folder tool...\n')
        f.write(f"User Name: {getuser()}\n")
        f.write(f"Date Executed: {ctime()}\n")
        f.write('User Parameters:\n')
        f.write(f"\tAdmin State Selected: {sourceState}\n")
        f.write(f"\tAdmin County Selected: {sourceCounty}\n")
        f.write(f"\tTract Entered: {str(tractNumber)}\n")
        f.write(f"\tOverwrite CLU: {str(owFlag)}\n")


def createTextFile(tract, farm):
    """ This function sets up the text file to begin recording all messages
        reported to the console.  The text file will be created in a folder
        called 'HEL_Text_Files' in argv[0].  The text file will have the prefix
        "NRCS_HEL_Determination" and the Tract, Farm and field numbers appended
        to the end.  Basic information will be collected and logged to the
        text file.  The function will return the full path to the text file."""
    try:
        # Set log file
        helTextNotesDir = path.join(path.dirname(argv[0]), 'HEL_Text_Files')
        if not path.isdir(helTextNotesDir):
           makedirs(helTextNotesDir)
        textFileName = f"NRCS_HEL_Determination_TRACT({str(tract)})_FARM({str(farm)}).txt"
        textPath = path.join(helTextNotesDir, textFileName)
        f = open(textPath,'a+')
        f.write('#' * 80 + "\n")
        f.write("NRCS HEL Determination Tool\n")
        f.write(f"User Name: {getuser()}\n")
        f.write(f"Date Executed: {ctime()}\n")
        f.close
        return textPath
    except:
        errorMsg()


def removeScratchLayers(scratchLayers):
    """ This function is the last task that is executed or gets invoked in an except clause. Removes all temporary scratch layers."""
    try:
        for lyr in scratchLayers:
            try:
                Delete(lyr)
            except:
                AddMsgAndPrint(f"\n\tDeleting Layer: {str(lyr)} failed.", 1)
                continue
    except:
        pass


def FindField(layer, chkField):
    """ Check table or featureclass to see if specified field exists. If fully qualified name is found, return that name;
        otherwise return Set workspace before calling FindField."""
    try:
        if Exists(layer):
            theDesc = Describe(layer)
            theFields = theDesc.fields
            theField = theFields[0]
            for theField in theFields:
                # Parses a fully qualified field name into its components (database, owner name, table name, and field name)
                parseList = ParseFieldName(theField.name) # (null), (null), (null), MUKEY
                # choose the last component which would be the field name
                theFieldname = parseList.split(',')[len(parseList.split(','))-1].strip()  # MUKEY
                if theFieldname.upper() == chkField.upper():
                    return theField.name
            return False
        else:
            AddMsgAndPrint('\tInput layer not found')
            return False
    except:
        errorMsg()
        return False


def extractDEMfromImageService(demSource, fieldDetermination, scratchWS, cluLayer, zFactorList, unitLookUpDict, zUnits):
    """ This function will extract a DEM from a Web Image Service that is in WGS. The CLU will be buffered to 410 meters
        and set to WGS84 GCS in order to clip the DEM. The clipped DEM will then be projected to the same coordinate system as the CLU.
        Eventually code will be added to determine the approximate cell size  of the image service using y-distances from the center of the cells.
        Cell size from a WGS84 service is difficult to calculate. Clip is the fastest however it doesn't honor cellsize so a project is required.
        Original Z-factor on WGS84 service cannot be calculated b/c linear units are unknown. Assume linear units and z-units are the same.
        Returns a clipped DEM and new Z-Factor"""
    try:
        desc = Describe(demSource)
        sr = desc.SpatialReference
        outputCellsize = 3

        AddMsgAndPrint(f"\nInput DEM Image Service: {desc.baseName}")
        AddMsgAndPrint(f"\tGeographic Coordinate System: {sr.Name}")
        AddMsgAndPrint(f"\tUnits (XY): {sr.AngularUnitName}")

        # Set output env variables to WGS84 to project clu
        env.geographicTransformations = 'WGS_1984_(ITRF00)_To_NAD_1983'
        env.outputCoordinateSystem = SpatialReference(4326)

        # Buffer CLU by 410 Meters. Output buffer will be in GCS
        cluBuffer = path.join('in_memory', path.basename(CreateScratchName('cluBuffer_GCS', data_type='FeatureClass', workspace=scratchWS)))
        Buffer(fieldDetermination, cluBuffer, '410 Meters', 'FULL', '', 'ALL', '')

        # Use the WGS 1984 AOI to clip/extract the DEM from the service
        cluExtent = Describe(cluBuffer).extent
        clipExtent = f"{str(cluExtent.XMin)} {str(cluExtent.YMin)} {str(cluExtent.XMax)} {str(cluExtent.YMax)}"

        SetProgressorLabel(f"Downloading DEM from {desc.baseName} Image Service")
        AddMsgAndPrint(f"\n\tDownloading DEM from {desc.baseName} Image Service")

        demClip = path.join('in_memory', path.basename(CreateScratchName('demClipIS', data_type='RasterDataset', workspace=scratchWS)))
        Clip_m(demSource, clipExtent, demClip, '', '', '', 'NO_MAINTAIN_EXTENT')

        # Project DEM subset from WGS84 to CLU coord system
        outputCS = Describe(cluLayer).SpatialReference
        env.outputCoordinateSystem = outputCS
        demProject = path.join('in_memory', path.basename(CreateScratchName('demProjectIS', data_type='RasterDataset', workspace=scratchWS)))
        ProjectRaster(demClip, demProject, outputCS, 'BILINEAR', outputCellsize)
        Delete(demClip)

        # Report new DEM properties
        desc = Describe(demProject)
        newSR = desc.SpatialReference
        newLinearUnits = newSR.LinearUnitName
        newCellSize = desc.MeanCellWidth

        # if zUnits not populated assume it is the same as linearUnits
        if not zUnits: zUnits = newLinearUnits
        newZfactor = zFactorList[unitLookUpDict.get(newLinearUnits)][unitLookUpDict.get(zUnits)]

        AddMsgAndPrint(f"\t\tNew Projection Name: {newSR.Name}")
        AddMsgAndPrint(f"\t\tLinear Units (XY): {newLinearUnits}")
        AddMsgAndPrint(f"\t\tElevation Units (Z): {zUnits}")
        AddMsgAndPrint(f"\t\tCell Size: {str(newCellSize)} {newLinearUnits}")
        AddMsgAndPrint(f"\t\tZ-Factor: {str(newZfactor)}")

        return newLinearUnits, newZfactor, demProject

    except:
        errorMsg()


def extractDEM(cluLayer, inputDEM, fieldDetermination, scratchWS, zFactorList, unitLookUpDict, zUnits):
    """ This function will return a DEM that has the same extent as the CLU selected fields buffered to 410 Meters. The DEM can be a local
        raster layer or a web image server. Datum must be in WGS84 or NAD83 and linear units must be in Meters or Feet otherwise it will exit.
        If the cell size is finer than 3M then the DEM will be resampled. The resampling will happen using the Project Raster tool regardless
        of an actual coordinate system change. If the cell size is 3M then the DEM will be clipped using the buffered CLU. Environment settings
        are used to control the output coordinate system. Function does not check the SR of the CLU. Assumes the CLU is in a projected coordinate
        system and in meters. Should probably verify before reprojecting to a 3M cell.
        Returns a clipped DEM and new Z-Factor"""
    try:
        # Set environment variables
        env.geographicTransformations = 'WGS_1984_(ITRF00)_To_NAD_1983'
        env.resamplingMethod = 'BILINEAR'
        env.outputCoordinateSystem = Describe(cluLayer).SpatialReference
        bImageService = False
        bResample = False
        outputCellSize = 3
        desc = Describe(inputDEM)
        sr = desc.SpatialReference
        linearUnits = sr.LinearUnitName
        cellSize = desc.MeanCellWidth

        if desc.format == 'Image Service':
            if sr.Type == 'Geographic':
                newLinearUnits, newZfactor, demExtract = extractDEMfromImageService(inputDEM, fieldDetermination, scratchWS, cluLayer, zFactorList, unitLookUpDict, zUnits)
                return newLinearUnits, newZfactor, demExtract
            bImageService = True

        # Check DEM properties
        if bImageService:
            AddMsgAndPrint(f"\nInput DEM Image Service: {desc.baseName}")
        else:
            AddMsgAndPrint(f"\nInput DEM Information: {desc.baseName}")

        # DEM must be in a project coordinate system to continue unless DEM is an image service.
        if not bImageService and sr.type != 'Projected':
            AddMsgAndPrint(f"\n\t{str(desc.name)} Must be in a projected coordinate system, Exiting", 2)
            AddMsgAndPrint('\tContact your State GIS Coordinator to resolve this issue', 2)
            return False, False

        # Linear units must be in Meters or Feet
        if not linearUnits:
            AddMsgAndPrint('\n\tCould not determine linear units of DEM....Exiting!', 2)
            AddMsgAndPrint('\tContact your State GIS Coordinator to resolve this issue', 2)
            return False, False

        if linearUnits == 'Meter':
            tolerance = 3.1
        elif linearUnits == 'Foot':
            tolerance = 10.1706
        elif linearUnits == 'Foot_US':
            tolerance = 10.1706
        else:
            AddMsgAndPrint(f"\n\tHorizontal units of {str(desc.baseName)} must be in feet or meters... Exiting!")
            AddMsgAndPrint('\tContact your State GIS Coordinator to resolve this issue', 2)
            return False, False

        # Cell size must be the equivalent of 3 meters.
        if cellSize > tolerance:
            AddMsgAndPrint('\n\tThe cell size of the input DEM must be 3 Meters (9.84252 FT) or less to continue... Exiting!', 2)
            AddMsgAndPrint(f"\t{str(desc.baseName)} has a cell size of {str(desc.MeanCellWidth)} {linearUnits}", 2)
            AddMsgAndPrint('\tContact your State GIS Coordinator to resolve this issue', 2)
            return False, False
        elif cellSize < tolerance:
            bResample = True
        else:
            bResample = False

        # if zUnits not populated assume it is the same as linearUnits
        bZunits = True
        if not zUnits: zUnits = linearUnits; bZunits = False

        # look up zFactor based on units (z,xy -- I should've reversed the table)
        zFactor = zFactorList[unitLookUpDict.get(linearUnits)][unitLookUpDict.get(zUnits)]

        # Print Input DEM properties
        AddMsgAndPrint(f"\tProjection Name: {sr.Name}")
        AddMsgAndPrint(f"\tLinear Units (XY): {linearUnits}")
        AddMsgAndPrint(f"\tElevation Units (Z): {zUnits}")
        if not bZunits:
            AddMsgAndPrint(f"\t\tZ-units were auto set to: {linearUnits}")
        AddMsgAndPrint(f"\tCell Size: {str(desc.MeanCellWidth)} {linearUnits}")
        AddMsgAndPrint(f"\tZ-Factor: {str(zFactor)}")

        # Extract DEM
        SetProgressorLabel('Buffering AOI by 410 Meters')
        cluBuffer = path.join('in_memory', path.basename(CreateScratchName('cluBuffer', data_type='FeatureClass', workspace=scratchWS)))
        Buffer(fieldDetermination, cluBuffer, '410 Meters', 'FULL', 'ROUND')
        env.extent = cluBuffer

        # CLU clip extents
        cluExtent = Describe(cluBuffer).extent
        clipExtent = f"{str(cluExtent.XMin)} {str(cluExtent.YMin)} {str(cluExtent.XMax)} {str(cluExtent.YMax)}"

        # Cell Resolution needs to change; Clip and Project
        if bResample:
            SetProgressorLabel(f"Changing resolution from {str(cellSize)} {linearUnits} to 3 Meters")
            AddMsgAndPrint(f"\n\tChanging resolution from {str(cellSize)} {linearUnits} to 3 Meters")
            demClip = path.join('in_memory', path.basename(CreateScratchName('demClip_resample', data_type='RasterDataset', workspace=scratchWS)))
            Clip_m(inputDEM, clipExtent, demClip, '', '', '', 'NO_MAINTAIN_EXTENT')
            demExtract = path.join('in_memory', path.basename(CreateScratchName('demClip_project', data_type='RasterDataset', workspace=scratchWS)))
            ProjectRaster(demClip, demExtract, env.outputCoordinateSystem, env.resamplingMethod, outputCellSize, '#', '#', sr)
            Delete(demClip)

        # Resolution is correct; Clip the raster
        else:
            SetProgressorLabel('Clipping DEM using buffered CLU')
            AddMsgAndPrint('\n\tClipping DEM using buffered CLU')
            demExtract = path.join('in_memory', path.basename(CreateScratchName('demClip', data_type='RasterDataset', workspace=scratchWS)))
            Clip_m(inputDEM, clipExtent, demExtract, '', '', '', 'NO_MAINTAIN_EXTENT')

        # Report any new DEM properties
        desc = Describe(demExtract)
        newSR = desc.SpatialReference
        newLinearUnits = newSR.LinearUnitName
        newCellSize = desc.MeanCellWidth
        newZfactor = zFactorList[unitLookUpDict.get(newLinearUnits)][unitLookUpDict.get(zUnits)]

        if newSR.name != sr.Name:
            AddMsgAndPrint(f"\t\tNew Projection Name: {newSR.Name}")
        if newCellSize != cellSize:
            AddMsgAndPrint(f"\t\tNew Cell Size: {str(newCellSize)} {newLinearUnits}")
        if newZfactor != zFactor:
            AddMsgAndPrint(f"\t\tNew Z-Factor: {str(newZfactor)}")

        Delete(cluBuffer)
        return newLinearUnits, newZfactor, demExtract

    except:
        errorMsg()
        return False, False, False
            

def queryIntersect(ws, temp_dir, fc, RESTurl, outFC, portalToken):
    """ This function uses a REST API query to retrieve geometry from that overlap an input feature class from a
        hosted feature service.
        Relies on a global variable of portalToken to exist and be active (checked before running this function)
        ws is a file geodatabase workspace to store temp files for processing
        fc is the input feature class. Should be a polygon feature class, but technically shouldn't fail if other types
        RESTurl is the url for the query where the target hosted data resides
        Example: https://gis-testing.usda.net/server/rest/services/Hosted/CWD_Training/FeatureServer/0/query
        outFC is the output feature class path/name that is return if the function succeeds AND finds data
        Otherwise False is returned """
    
    # Set variables
    query_url = f"{RESTurl}/query"
    jfile = path.join(temp_dir, 'jsonFile.json')
    wmas_fc = path.join(ws, 'wmas_fc')
    wmas_dis = path.join(ws, 'wmas_dis_fc')
    wmas_sr = SpatialReference(3857)

    # Convert the input feature class to Web Mercator and to JSON
    Project(fc, wmas_fc, wmas_sr)
    Dissolve(wmas_fc, wmas_dis, "", "", "MULTI_PART", "")
    jsonPolygon = [row[0] for row in SearchCursor(wmas_dis, ['SHAPE@JSON'])][0]

    # Setup parameters for query
    params = parse.urlencode({
        'f': 'json',
        'geometry':jsonPolygon,
        'geometryType':'esriGeometryPolygon',
        'spatialRelationship':'esriSpatialRelIntersects',
        'returnGeometry':'true',
        'outFields':'*',
        'token': portalToken['token']
    })

    INparams = params.encode('ascii')
    resp = request.urlopen(query_url, INparams)

    responseStatus = resp.getcode()
    jsonString = resp.read()

    if responseStatus > 200:
        AddMsgAndPrint(f"\nHost Feature Service {RESTurl} may be inaccessible or query may be invalid.", 1)
        AddMsgAndPrint('\nReturning to mainline functions...', 1)
        return False

    results = json_loads(jsonString)

    # Check for error in results and exit with message if found.
    if 'error' in results.keys():
        if results['error']['message'] == 'Invalid Token':
            AddMsgAndPrint('\nSign-in token expired. Sign-out and sign-in to the portal again and then re-run. Exiting...', 2)
            exit()
        else:
            AddMsgAndPrint(f"\nUnknown error encountered. Host Feature Service {RESTurl} may be inaccessible or query may be invalid. Continuing...", 1)
            AddMsgAndPrint(f"\nResponse status code: {str(responseStatus)}", 1)
            return False
    else:
        # Convert results to a feature class
        if not len(results['features']):
            return False
        else:
            with open(jfile, 'w') as outfile:
                json_dump(results, outfile)

            JSONToFeatures(jfile, outFC)
            Delete(jfile)
            Delete(wmas_fc)
            Delete(wmas_dis)
            return outFC


def addOutputLayers(lidarHEL, helSummary, finalHELSummary, fieldDetermination):
    SetParameterAsText(5, lidarHEL)
    SetParameterAsText(6, helSummary)
    SetParameterAsText(7, finalHELSummary)
    SetParameterAsText(8, fieldDetermination)
    return True


# TODO: Needs to be converted to Pro
# def AddLayersToArcMap():
#     """ Adds necessary layers to ArcMap. If no PHEL Values were present only 2 layers will be added: Initial HEL Summary
#         and Field Determination. If PHEL values were processed than all 4 layers will be added. This function does not
#         utilize the setParameterastext function to add layers to arcmap through the toolbox."""
#     try:
#         # Add 3 fields to the field determination layer and populate them
#         # from ogCLUinfoDict and 4 fields to the Final HEL Summary layer that
#         # otherwise would've been added after geoprocessing was successful.
#         if bNoPHELvalues or bSkipGeoprocessing:
#             # TODO: Need to include AddFields
#             # Add 3 fields to fieldDetermination layer
#             fieldList = ['HEL_YES', 'HEL_Acres', 'HEL_Pct']
#             for field in fieldList:
#                 if not FindField(fieldDetermination, field):
#                     if field == 'HEL_YES':
#                         AddField(fieldDetermination, field, 'TEXT', '', '', 5)
#                     else:
#                         AddField(fieldDetermination, field, 'FLOAT')
#             fieldList.append(cluNumberFld)

#             # Update new fields using ogCLUinfoDict
#             with UpdateCursor(fieldDetermination, fieldList) as cursor:
#                 for row in cursor:
#                     row[0] = ogCLUinfoDict.get(row[3])[0]   # "HEL_YES" value
#                     row[1] = ogCLUinfoDict.get(row[3])[1]   # "HEL_Acres" value
#                     row[2] = ogCLUinfoDict.get(row[3])[2]   # "HEL_Pct" value
#                     cursor.updateRow(row)

#             # Add 4 fields to Final HEL Summary layer
#             # TODO: Need to include
#             newFields = ['Polygon_Acres', 'Final_HEL_Value', 'Final_HEL_Acres', 'Final_HEL_Percent']
#             for fld in newFields:
#                 if not len(ListFields(finalHELSummary, fld)) > 0:
#                     if fld == 'Final_HEL_Value':
#                         AddField(finalHELSummary, 'Final_HEL_Value', 'TEXT', '', '', 5)
#                     else:
#                         AddField(finalHELSummary, fld, 'DOUBLE')
#             newFields.append(helFld)
#             newFields.append(cluNumberFld)
#             newFields.append('SHAPE@AREA')

#             # [polyAcres,finalHELvalue,finalHELacres,finalHELpct,MUHELCL,'CLUNBR',"SHAPE@AREA"]
#             # TODO: Need to include
#             with UpdateCursor(finalHELSummary, newFields) as cursor:
#                 for row in cursor:
#                     # Calculate polygon acres;
#                     # TODO: change to GIS calc acres, remove dict
#                     row[0] = row[6] / acreConversionDict.get(Describe(finalHELSummary).SpatialReference.LinearUnitName)
#                     # Final_HEL_Value will be set to the initial HEL value
#                     row[1] = row[4]
#                     # set Final HEL Acres to 0 for PHEL and NHEL; othewise set to polyAcres
#                     if row[4] in ('NHEL', 'PHEL'):
#                         row[2] = 0.0
#                     else:
#                         row[2] = row[0]
#                     # Calculate percent of polygon relative to CLU
#                     cluAcres = ogCLUinfoDict.get(row[5])[1]
#                     pct = (row[0] / cluAcres) * 100
#                     if pct > 100.0: pct = 100.0
#                     row[3] = pct
#                     del cluAcres,pct
#                     cursor.updateRow(row)
#             del cursor

#         # Put this section in a try-except. It will fail if run from ArcCatalog
#         # TODO: use derived output parameters 
#         mxd = MapDocument('CURRENT')
#         #aprx = ArcGISProject('CURRENT') #pro
#         df = ListDataFrames(mxd)[0]
#         #maps = aprx.listMaps()[0] #pro

#         # Workaround:  ListLayers returns a list of layer objects. Need to create a list of layer name Strings
#         currentLayersObj = ListLayers(mxd)
#         #currentLayersObj = maps.listLayers() #pro
#         currentLayersStr = [str(x) for x in ListLayers(mxd)]
#         #currentLayersStr = [str(x) for x in maps.listLayers()] #pro

#         # List of layers to add to Arcmap (layer path, arcmap layer name)
#         if bNoPHELvalues or bSkipGeoprocessing:
#             addToArcMap = [(helSummary, 'Initial HEL Summary'), (fieldDetermination, 'Field Determination')]
#             # Remove these layers from arcmap if they are present since they were not produced
#             if 'LiDAR HEL Summary' in currentLayersStr:
#                 RemoveLayer(df, currentLayersObj[currentLayersStr.index('LiDAR HEL Summary')])
#                 #LayerFile.removeLayer(currentLayersObj[currentLayersStr.index('LiDAR HEL Summary')]) #pro
#             if 'Final HEL Summary' in currentLayersStr:
#                 RemoveLayer(df, currentLayersObj[currentLayersStr.index('Final HEL Summary')])
#                 #LayerFile.removeLayer(currentLayersObj[currentLayersStr.index('Final HEL Summary')]) #pro
#         else:
#             addToArcMap = [(lidarHEL, 'LiDAR HEL Summary'), (helSummary, 'Initial HEL Summary'), (finalHELSummary, 'Final HEL Summary'), (fieldDetermination, 'Field Determination')]

#         for layer in addToArcMap:
#             # remove layer from ArcMap if it exists
#             if layer[1] in currentLayersStr:
#                 RemoveLayer(df, currentLayersObj[currentLayersStr.index(layer[1])])
#                 #LayerFile.removeLayer(currentLayersObj[currentLayersStr.index(layer[1])]) #pro
#             # Raster Layers need to be handled differently than vector layers
#             if layer[1] == 'LiDAR HEL Summary':
#                 rasterLayer = MakeRasterLayer(layer[0], layer[1])
#                 tempLayer = rasterLayer.getOutput(0)
#                 AddLayer(df, tempLayer, 'TOP')
#                 #maps.addLayer(tempLayer, 'TOP') #pro
#                 # define the symbology layer and convert it to a layer object
                
#                 #This section below to UpdateLayer I don't think is necessary because of how pro updates.
#                 updateLayer = ListLayers(mxd, layer[1], df)[0]
#                 #updateLayer = maps.listLayers(layer[1]) #pro
#                 symbologyLyr = path.join(path.dirname(argv[0]), layer[1].lower().replace(' ', '') + '.lyr')
#                 sourceLayer = Layer(symbologyLyr) #this should work the same in pro after removing the .mapping imports
#                 UpdateLayer(df, updateLayer, sourceLayer) #I don't think this is necessary in Pro
#             else:
#                 # add layer to arcmap
#                 symbologyLyr = path.join(path.dirname(argv[0]), layer[1].lower().replace(' ', '') + '.lyr')
#                 AddLayer(df, Layer(symbologyLyr.strip("'")), 'TOP')
#                 #maps.addLayer(Layer(symbologyLyr.strip("'")), 'TOP') #pro

#             # This layer should be turned on if no PHEL values were processed. Symbology should also be updated to reflect current values.
#             if layer[1] in ('Initial HEL Summary') and bNoPHELvalues:
#                 for lyr in ListLayers(mxd, layer[1]):
#                 #for lyr in maps.listLayers(layer[1]): #pro
#                     lyr.visible = True

#             # these 2 layers should be turned off by default if full processing happens
#             if layer[1] in ('Initial HEL Summary', 'LiDAR HEL Summary') and not bNoPHELvalues:
#                 for lyr in ListLayers(mxd, layer[1]):
#                 #for lyr in ListLayers(layer[1]): #pro
#                     lyr.visible = False

#             AddMsgAndPrint(f"Added {layer[1]} to your ArcMap Session")

#         # Unselect CLU polygons; Looks goofy after processed layers have been added to ArcMap. Turn it off as well
#         for lyr in ListLayers(mxd, '*' + str(Describe(cluLayer).nameString).split('\\')[-1], df):
#         #for lyr in maps.listLayers('*' + str(Describe(cluLayer).nameString).split('\\')[-1]): #pro
#             SelectLayerByAttribute(lyr, 'CLEAR_SELECTION')
#             #TODO: clear selection on original input CLU Layer
#             lyr.visible = False

#         # Turn off the original HEL layer to put the outputs into focus
#         helLyr = ListLayers(mxd, Describe(helLayer).nameString, df)[0]
#         #helLyr = maps.listLayers(Describe(helLayer).nameString)[0] #pro
#         helLyr.visible = False
#         # set dataframe extent to the extent of the Field Determintation layer buffered by 50 meters.
#         # NOTE: no need to try and zoom to tract in map frame
#         fieldDeterminationBuffer = path.join('in_memory', path.basename(CreateScratchName('fdBuffer', data_type='FeatureClass', workspace=scratchWS)))
#         Buffer(fieldDetermination, fieldDeterminationBuffer, '50 Meters', 'FULL', '', 'ALL', '')
#         df.extent = Describe(fieldDeterminationBuffer).extent
#         Delete(fieldDeterminationBuffer)
#         RefreshTOC() # No longer needed as changes in Pro are immediate
#         RefreshActiveView() # No longer needed as changes in Pro are immediate

#     except:
#         errorMsg()


def populateForm(fieldDetermination, lu_table, dcSignature, input_cust, helDatabase):
    """ This function will prepare the 1026 form by adding 16 fields to the fieldDetermination feauture class. This function will
        still be invoked even if there was no geoprocessing executed due to no PHEL values to be computed. If bNoPHELvalues is true,
        3 fields will be added that otherwise would've been added after HEL geoprocessing. However, values to these fields will be
        populated from the ogCLUinfoDict and represent original HEL values primarily determined by the 33.33% or 50 acre rule."""
    try:
        AddMsgAndPrint('\nPreparing and Populating NRCS-CPA-026 Form')
        today = date.today()
        today = today.strftime('%b %d, %Y')
        stateCodeDict = {'WA': '53', 'DE': '10', 'DC': '11', 'WI': '55', 'WV': '54', 'HI': '15',
                        'FL': '12', 'WY': '56', 'PR': '72', 'NJ': '34', 'NM': '35', 'TX': '48',
                        'LA': '22', 'NC': '37', 'ND': '38', 'NE': '31', 'TN': '47', 'NY': '36',
                        'PA': '42', 'AK': '02', 'NV': '32', 'NH': '33', 'VA': '51', 'CO': '08',
                        'CA': '06', 'AL': '01', 'AR': '05', 'VT': '50', 'IL': '17', 'GA': '13',
                        'IN': '18', 'IA': '19', 'MA': '25', 'AZ': '04', 'ID': '16', 'CT': '09',
                        'ME': '23', 'MD': '24', 'OK': '40', 'OH': '39', 'UT': '49', 'MO': '29',
                        'MN': '27', 'MI': '26', 'RI': '44', 'KS': '20', 'MT': '30', 'MS': '28',
                        'SC': '45', 'KY': '21', 'OR': '41', 'SD': '46'}

        # Try to get the state using the field determination layer, Otherwise get the state from the computer user name
        try:
            stateCode = ([row[0] for row in SearchCursor(fieldDetermination, 'STATECD')])
            state = [stAbbrev for (stAbbrev, code) in list(stateCodeDict.items()) if code == stateCode[0]][0]
        except:
            state = getuser().replace('.', ' ').replace('\'', '')

        # Get the geographic and admin counties
        stFIPS = ''
        coFIPS = ''
        astFIPS = ''
        acoFIPS = ''

        # check that lookup table exists and search it for geographic and admin state and county codes
        if Exists(lu_table):
            with SearchCursor(fieldDetermination, ['statecd', 'countycd', 'admnstate', 'admncounty']) as cursor:
                for row in cursor:
                    stFIPS = row[0]
                    coFIPS = row[1]
                    astFIPS = row[2]
                    acoFIPS = row[3]
                    break

            # Lookup state and county names from recovered fips codes
            GeoCode = str(stFIPS) + str(coFIPS)
            expression1 = ("{} = '" + GeoCode + "'").format(AddFieldDelimiters(lu_table, 'GEOID'))
            with SearchCursor(lu_table, ['GEOID', 'NAME', 'STPOSTAL'], where_clause=expression1) as cursor:
                for row in cursor:
                    GeoCounty = row[1]
                    GeoState = row[2]
                    # We should only get one result if using installed lookup table from US Census Tiger table, so break
                    break

            AdminCode = str(astFIPS) + str(acoFIPS)
            expression2 = ("{} = '" + AdminCode + "'").format(AddFieldDelimiters(lu_table, 'GEOID'))
            with SearchCursor(lu_table, ['GEOID', 'NAME', 'STPOSTAL'], where_clause=expression2) as cursor:
                for row in cursor:
                    AdminCounty = row[1]
                    AdminState = row[2]
                    # We should only get one result if using installed lookup table from US Census Tiger table, so break
                    break

            GeoLocation = f"{GeoCounty}, {GeoState}"
            AdminLocation = f"{AdminCounty}, {AdminState}"

        else:
            GeoLocation = 'Not Found'
            AdminLocation = 'Not Found'

        # Add 20 Fields to the fieldDetermination feature class
        remarks_txt = r'This Highly Erodible Land determination was conducted offsite using the soil survey. If PHEL soil map units were present, they may have been evaluated using elevation data.'
        fieldDict = {'Signature':('TEXT',dcSignature,50),'SoilAvailable':('TEXT','Yes',5),'Completion':('TEXT','Office',10),
                        'SodbustField':('TEXT','No',5),'Delivery':('TEXT','Mail',10),'Remarks':('TEXT',remarks_txt,255),
                        'RequestDate':('DATE',''),'LastName':('TEXT','',50),'FirstName':('TEXT',input_cust,50),'Address':('TEXT','',50),
                        'City':('TEXT','',25),'ZipCode':('TEXT','',10),'Request_from':('TEXT','AD-1026',15),'HELFarm':('TEXT','0',5),
                        'Determination_Date':('DATE',today),'state':('TEXT',state,2),'SodbustTract':('TEXT','No',5),'Lidar':('TEXT','Yes',5),
                        'Location_County':('TEXT',GeoLocation,50),'Admin_County':('TEXT',AdminLocation,50)}

        SetProgressor('step', 'Preparing and Populating NRCS-CPA-026 Form', 0, len(fieldDict), 1)

        for field,params in fieldDict.items():
            SetProgressorLabel(f"Adding Field: {field} to 'Field Determination' layer")
            try:
                fldLength = params[2]
            except:
                fldLength = 0
                pass
            AddField(fieldDetermination, field, params[0], '#', '#', fldLength)
            if len(params[1]) > 0:
                expression = '\'' + params[1] + '\''
                CalculateField(fieldDetermination, field, expression, 'VB') #TODO: Change to Python expression

        AddMsgAndPrint('\tOpening NRCS-CPA-026 Form')
        try:
            startfile(helDatabase)
        except:
            AddMsgAndPrint('\tCould not locate the Microsoft Word', 1)
            AddMsgAndPrint('\tOpen Microsoft Word manually to access the NRCS-CPA-026 Form', 1)

        return True
    except:
        return False


# def configLayout(lu_table, fieldDetermination, input_cust):
#     """ This function will gather and update information for elements of the map layout"""
#     try:
#         # Confirm Map Doc existence
#         mxd = MapDocument('CURRENT')
#         #project = ArcGISProject('CURRENT')
#         # Set starting variables as empty for logical comparison later on, prior to updating layout
#         farmNum = ''
#         trNum = ''
#         county = ''
#         state = ''

#         # end function if lookup table from tool parameters does not exist
#         if not Exists(lu_table):
#             return False

#         # Get CLU information from first row of the cluLayer.
#         # All CLU records should have this info, so break after one record.
#         with SearchCursor(fieldDetermination, ['statecd', 'countycd', 'farmnbr', 'tractnbr']) as cursor:
#             for row in cursor:
#                 stCD = row[0]
#                 coCD = row[1]
#                 farmNum = row[2]
#                 trNum = row[3]
#                 break

#         # Lookup state and county name
#         stco_code = str(stCD) + str(coCD)
#         expression = ("{} = '" + stco_code + "'").format(AddFieldDelimiters(lu_table, 'GEOID'))
#         with SearchCursor(lu_table, ['GEOID', 'NAME', 'STPOSTAL'], where_clause=expression) as cursor:
#             for row in cursor:
#                 county = row[1]
#                 state = row[2]
#                 # We should only get one result if using installed lookup table from US Census Tiger table, so break
#                 break

#         # Find and hook map layout elements to variables
#         for elm in ListLayoutElements(mxd):
#         #for elm in project.listLayouts() #pro
#             if elm.name == 'farm_txt':
#                 farm_ele = elm
#             if elm.name == 'tract_txt':
#                 tract_ele = elm
#             if elm.name == 'customer_txt':
#                 customer_ele = elm
#             if elm.name == 'county_txt':
#                 county_ele = elm
#             if elm.name == 'state_txt':
#                 state_ele = elm #Not used anywhere?

#         # Configure the text boxes
#         # If any of the info is missing, the script still runs and boxes are reset to a manual entry prompt
#         if farmNum != '':
#             farm_ele.text = f"Farm: {str(farmNum)}"
#         else:
#             farm_ele.text = 'Farm: <dbl-click to enter>'
#         if trNum != '':
#             tract_ele.text = f"Tract: {str(trNum)}"
#         else:
#             tract_ele.text = 'Tract: <dbl-click to enter>'
#         if input_cust != '':
#             customer_ele.text = f"Customer(s): {str(input_cust)}"
#         else:
#             customer_ele.text = 'Customer(s): <dbl-click to enter>'
#         if county != '':
#             county_ele.text = f"County: {str(county)}, {str(state)}"
#         else:
#             county_ele.text = 'County: <dbl-click to enter>'

#         return True
#     except:
#         return False
