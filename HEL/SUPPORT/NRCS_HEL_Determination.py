from datetime import date
from getpass import getuser
from math import pi
from os import makedirs, path, remove, sep, startfile, system
from subprocess import Popen
from sys import argv, exc_info, exit, getwindowsversion
from time import ctime, sleep
from traceback import format_exception

from arcpy import AddError, AddFieldDelimiters, AddMessage, AddWarning, CheckExtension, CheckOutExtension, CreateScratchName, \
    Describe, env, Exists, GetInstallInfo, GetParameter, GetParameterAsText, ListFields, ParseFieldName, Reclassify_3d, \
    RefreshCatalog, SetProgressor, SetProgressorLabel, SpatialReference

from arcpy.analysis import Buffer, Clip as Clip_a, Intersect, Statistics
from arcpy.conversion import FeatureToRaster, RasterToPolygon
from arcpy.da import SearchCursor, UpdateCursor

from arcpy.management import AddField, CalculateField, Clip as Clip_m, CopyFeatures, CreateFileGDB, Delete, DeleteField, \
    Dissolve, JoinField, MakeRasterLayer, MultipartToSinglepart, PivotTable, ProjectRaster, SelectLayerByAttribute

# TODO: All arcpy.mapping must be converted to arcpy.mp (MapDocument --> ArcGISProject)
from arcpy.mapping import AddLayer, Layer, ListDataFrames, ListLayoutElements, ListLayers, MapDocument, RefreshActiveView, \
    RefreshTOC, RemoveLayer, UpdateLayer

from arcpy.sa import ATan, Con, Cos, Divide, Fill, FlowDirection, FlowLength, FocalStatistics, IsNull, NbrRectangle, \
    Power, SetNull, Slope, Sin, TabulateArea, Times


def AddMsgAndPrint(msg, severity=0):
    """ Adds tool message to the geoprocessor. Split the message on \n first, so a GPMessage will be added for each line."""
    try:
        if bLog:
            f = open(textFilePath,'a+')
            f.write(msg + " \n")
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
        theMsg = "\t" + format_exception(exc_type, exc_value, exc_traceback)[1] + "\n\t" + format_exception(exc_type, exc_value, exc_traceback)[-1]
        if theMsg.find("sys.exit") > -1:
            AddMsgAndPrint("\n\n")
            pass
        else:
            AddMsgAndPrint("\n\tNRCS HEL Tool Error: -------------------------",2)
            AddMsgAndPrint(theMsg,2)
    except:
        AddMsgAndPrint("Unhandled error in errorMsg method", 2)
        pass


def createTextFile(Tract, Farm):
    """ This function sets up the text file to begin recording all messages
        reported to the console.  The text file will be created in a folder
        called 'HEL_Text_Files' in argv[0].  The text file will have the prefix
        "NRCS_HEL_Determination" and the Tract, Farm and field numbers appended
        to the end.  Basic information will be collected and logged to the
        text file.  The function will return the full path to the text file."""
    try:
        # Set log file
        helTextNotesDir = path.dirname(argv[0]) + sep + 'HEL_Text_Files'
        if not path.isdir(helTextNotesDir):
           makedirs(helTextNotesDir)
        textFileName = "NRCS_HEL_Determination_TRACT(" + str(Tract) + ")_FARM(" + str(Farm) + ").txt"
        textPath = helTextNotesDir + sep + textFileName
        # Version check
        version = str(GetInstallInfo()['Version'])
        versionFlt = float(version[0:4])
        if versionFlt < 10.3:
            AddMsgAndPrint("\nThis tool has only been tested on ArcGIS version 10.4 or greater",1)
        f = open(textPath,'a+')
        f.write('#' * 80 + "\n")
        f.write("NRCS HEL Determination Tool\n")
        f.write("User Name: " + getuser() + "\n")
        f.write("Date Executed: " + ctime() + "\n")
        f.write("ArcGIS Version: " + str(version) + "\n")
        f.close
        return textPath
    except:
        errorMsg()


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
                theFieldname = parseList.split(",")[len(parseList.split(','))-1].strip()  # MUKEY
                if theFieldname.upper() == chkField.upper():
                    return theField.name
            return False
        else:
            AddMsgAndPrint("\tInput layer not found", 0)
            return False
    except:
        errorMsg()
        return False


def extractDEMfromImageService(demSource, zUnits):
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

        AddMsgAndPrint("\nInput DEM Image Service: " + desc.baseName)
        AddMsgAndPrint("\tGeographic Coordinate System: " + sr.Name)
        AddMsgAndPrint("\tUnits (XY): " + sr.AngularUnitName)

        # Set output env variables to WGS84 to project clu
        env.geographicTransformations = "WGS_1984_(ITRF00)_To_NAD_1983"
        env.outputCoordinateSystem = SpatialReference(4326)

        # Buffer CLU by 410 Meters. Output buffer will be in GCS
        cluBuffer = "in_memory" + sep + path.basename(CreateScratchName("cluBuffer_GCS",data_type="FeatureClass",workspace=scratchWS))
        Buffer(fieldDetermination, cluBuffer, "410 Meters", "FULL", "", "ALL", "")

        # Use the WGS 1984 AOI to clip/extract the DEM from the service
        cluExtent = Describe(cluBuffer).extent
        clipExtent = str(cluExtent.XMin) + " " + str(cluExtent.YMin) + " " + str(cluExtent.XMax) + " " + str(cluExtent.YMax)

        SetProgressorLabel("Downloading DEM from " + desc.baseName + " Image Service")
        AddMsgAndPrint("\n\tDownloading DEM from " + desc.baseName + " Image Service")

        demClip = "in_memory" + sep + path.basename(CreateScratchName("demClipIS",data_type="RasterDataset",workspace=scratchWS))
        Clip_m(demSource, clipExtent, demClip, "", "", "", "NO_MAINTAIN_EXTENT")

        # Project DEM subset from WGS84 to CLU coord system
        outputCS = Describe(cluLayer).SpatialReference
        env.outputCoordinateSystem = outputCS
        demProject = "in_memory" + sep + path.basename(CreateScratchName("demProjectIS",data_type="RasterDataset",workspace=scratchWS))
        ProjectRaster(demClip, demProject, outputCS, "BILINEAR", outputCellsize)
        Delete(demClip)

        # Report new DEM properties
        desc = Describe(demProject)
        newSR = desc.SpatialReference
        newLinearUnits = newSR.LinearUnitName
        newCellSize = desc.MeanCellWidth

        # if zUnits not populated assume it is the same as linearUnits
        if not zUnits: zUnits = newLinearUnits
        newZfactor = zFactorList[unitLookUpDict.get(newLinearUnits)][unitLookUpDict.get(zUnits)]

        AddMsgAndPrint("\t\tNew Projection Name: " + newSR.Name,0)
        AddMsgAndPrint("\t\tLinear Units (XY): " + newLinearUnits)
        AddMsgAndPrint("\t\tElevation Units (Z): " + zUnits)
        AddMsgAndPrint("\t\tCell Size: " + str(newCellSize) + " " + newLinearUnits )
        AddMsgAndPrint("\t\tZ-Factor: " + str(newZfactor))

        return newLinearUnits,newZfactor,demProject

    except:
        errorMsg()


def extractDEM(inputDEM, zUnits):
    """ This function will return a DEM that has the same extent as the CLU selected fields buffered to 410 Meters. The DEM can be a local
        raster layer or a web image server. Datum must be in WGS84 or NAD83 and linear units must be in Meters or Feet otherwise it will exit.
        If the cell size is finer than 3M then the DEM will be resampled. The resampling will happen using the Project Raster tool regardless
        of an actual coordinate system change. If the cell size is 3M then the DEM will be clipped using the buffered CLU. Environment settings
        are used to control the output coordinate system. Function does not check the SR of the CLU. Assumes the CLU is in a projected coordinate
        system and in meters. Should probably verify before reprojecting to a 3M cell.
        Returns a clipped DEM and new Z-Factor"""
    try:
        # Set environment variables
        env.geographicTransformations = "WGS_1984_(ITRF00)_To_NAD_1983"
        env.resamplingMethod = "BILINEAR"
        env.outputCoordinateSystem = Describe(cluLayer).SpatialReference
        bImageService = False
        bResample = False
        outputCellSize = 3
        desc = Describe(inputDEM)
        sr = desc.SpatialReference
        linearUnits = sr.LinearUnitName
        cellSize = desc.MeanCellWidth

        if desc.format == 'Image Service':
            if sr.Type == "Geographic":
                newLinearUnits,newZfactor,demExtract = extractDEMfromImageService(inputDEM,zUnits)
                return newLinearUnits,newZfactor,demExtract
            bImageService = True

        # Check DEM properties
        if bImageService:
            AddMsgAndPrint("\nInput DEM Image Service: " + desc.baseName)
        else:
            AddMsgAndPrint("\nInput DEM Information: " + desc.baseName)

        # DEM must be in a project coordinate system to continue unless DEM is an image service.
        if not bImageService and sr.type != 'Projected':
            AddMsgAndPrint("\n\t" + str(desc.name) + " Must be in a projected coordinate system, Exiting",2)
            AddMsgAndPrint("\tContact your State GIS Coordinator to resolve this issue",2)
            return False,False

        # Linear units must be in Meters or Feet
        if not linearUnits:
            AddMsgAndPrint("\n\tCould not determine linear units of DEM....Exiting!",2)
            AddMsgAndPrint("\tContact your State GIS Coordinator to resolve this issue",2)
            return False,False

        if linearUnits == "Meter":
            tolerance = 3.1
        elif linearUnits == "Foot":
            tolerance = 10.1706
        elif linearUnits == "Foot_US":
            tolerance = 10.1706
        else:
            AddMsgAndPrint("\n\tHorizontal units of " + str(desc.baseName) + " must be in feet or meters... Exiting!")
            AddMsgAndPrint("\tContact your State GIS Coordinator to resolve this issue",2)
            return False,False

        # Cell size must be the equivalent of 3 meters.
        if cellSize > tolerance:
            AddMsgAndPrint("\n\tThe cell size of the input DEM must be 3 Meters (9.84252 FT) or less to continue... Exiting!",2)
            AddMsgAndPrint("\t" + str(desc.baseName) + " has a cell size of " + str(desc.MeanCellWidth) + " " + linearUnits,2)
            AddMsgAndPrint("\tContact your State GIS Coordinator to resolve this issue",2)
            return False,False
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
        AddMsgAndPrint("\tProjection Name: " + sr.Name)
        AddMsgAndPrint("\tLinear Units (XY): " + linearUnits)
        AddMsgAndPrint("\tElevation Units (Z): " + zUnits)
        if not bZunits:
            AddMsgAndPrint("\t\tZ-units were auto set to: " + linearUnits,0)
        AddMsgAndPrint("\tCell Size: " + str(desc.MeanCellWidth) + " " + linearUnits)
        AddMsgAndPrint("\tZ-Factor: " + str(zFactor))

        # Extract DEM
        SetProgressorLabel("Buffering AOI by 410 Meters")
        cluBuffer = "in_memory" + sep + path.basename(CreateScratchName("cluBuffer",data_type="FeatureClass",workspace=scratchWS))
        Buffer(fieldDetermination,cluBuffer,"410 Meters","FULL","ROUND")
        env.extent = cluBuffer

        # CLU clip extents
        cluExtent = Describe(cluBuffer).extent
        clipExtent = str(cluExtent.XMin) + " " + str(cluExtent.YMin) + " " + str(cluExtent.XMax) + " " + str(cluExtent.YMax)

        # Cell Resolution needs to change; Clip and Project
        if bResample:
            SetProgressorLabel("Changing resolution from " + str(cellSize) + " " + linearUnits + " to 3 Meters")
            AddMsgAndPrint("\n\tChanging resolution from " + str(cellSize) + " " + linearUnits + " to 3 Meters")
            demClip = "in_memory" + sep + path.basename(CreateScratchName("demClip_resample",data_type="RasterDataset",workspace=scratchWS))
            Clip_m(inputDEM, clipExtent, demClip, "", "", "", "NO_MAINTAIN_EXTENT")
            demExtract = "in_memory" + sep + path.basename(CreateScratchName("demClip_project",data_type="RasterDataset",workspace=scratchWS))
            ProjectRaster(demClip,demExtract,env.outputCoordinateSystem,env.resamplingMethod,outputCellSize,"#","#",sr)
            Delete(demClip)

        # Resolution is correct; Clip the raster
        else:
            SetProgressorLabel("Clipping DEM using buffered CLU")
            AddMsgAndPrint("\n\tClipping DEM using buffered CLU")
            demExtract = "in_memory" + sep + path.basename(CreateScratchName("demClip",data_type="RasterDataset",workspace=scratchWS))
            Clip_m(inputDEM, clipExtent, demExtract, "", "", "", "NO_MAINTAIN_EXTENT")

        # Report any new DEM properties
        desc = Describe(demExtract)
        newSR = desc.SpatialReference
        newLinearUnits = newSR.LinearUnitName
        newCellSize = desc.MeanCellWidth
        newZfactor = zFactorList[unitLookUpDict.get(newLinearUnits)][unitLookUpDict.get(zUnits)]

        if newSR.name != sr.Name:
            AddMsgAndPrint("\t\tNew Projection Name: " + newSR.Name,0)
        if newCellSize != cellSize:
            AddMsgAndPrint("\t\tNew Cell Size: " + str(newCellSize) + " " + newLinearUnits )
        if newZfactor != zFactor:
            AddMsgAndPrint("\t\tNew Z-Factor: " + str(newZfactor))

        Delete(cluBuffer)
        return newLinearUnits,newZfactor,demExtract

    except:
        errorMsg()
        return False,False,False


def removeScratchLayers():
    """ This function is the last task that is executed or gets invoked in an except clause. Removes all temporary scratch layers."""
    try:
        for lyr in scratchLayers:
            try:
                Delete(lyr)
            except:
                AddMessage("Deleting Layer: " + str(lyr) + " failed.")
                continue
    except:
        pass


def AddLayersToArcMap():
    """ Adds necessary layers to ArcMap. If no PHEL Values were present only 2 layers will be added: Initial HEL Summary
        and Field Determination. If PHEL values were processed than all 4 layers will be added. This function does not
        utilize the setParameterastext function to add layers to arcmap through the toolbox."""
    try:
        # Add 3 fields to the field determination layer and populate them
        # from ogCLUinfoDict and 4 fields to the Final HEL Summary layer that
        # otherwise would've been added after geoprocessing was successful.
        if bNoPHELvalues or bSkipGeoprocessing:

            # Add 3 fields to fieldDetermination layer
            fieldList = ["HEL_YES","HEL_Acres","HEL_Pct"]
            for field in fieldList:
                if not FindField(fieldDetermination,field):
                    if field == "HEL_YES":
                        AddField(fieldDetermination,field,"TEXT","","",5)
                    else:
                        AddField(fieldDetermination,field,"FLOAT")
            fieldList.append(cluNumberFld)

            # Update new fields using ogCLUinfoDict
            with UpdateCursor(fieldDetermination,fieldList) as cursor:
                 for row in cursor:
                     row[0] = ogCLUinfoDict.get(row[3])[0]   # "HEL_YES" value
                     row[1] = ogCLUinfoDict.get(row[3])[1]   # "HEL_Acres" value
                     row[2] = ogCLUinfoDict.get(row[3])[2]   # "HEL_Pct" value
                     cursor.updateRow(row)

            # Add 4 fields to Final HEL Summary layer
            newFields = ['Polygon_Acres','Final_HEL_Value','Final_HEL_Acres','Final_HEL_Percent']
            for fld in newFields:
                if not len(ListFields(finalHELSummary,fld)) > 0:
                   if fld == 'Final_HEL_Value':
                      AddField(finalHELSummary,'Final_HEL_Value',"TEXT","","",5)
                   else:
                        AddField(finalHELSummary,fld,"DOUBLE")

            newFields.append(helFld)
            newFields.append(cluNumberFld)
            newFields.append("SHAPE@AREA")

            # [polyAcres,finalHELvalue,finalHELacres,finalHELpct,MUHELCL,'CLUNBR',"SHAPE@AREA"]
            with UpdateCursor(finalHELSummary,newFields) as cursor:
                for row in cursor:
                    # Calculate polygon acres;
                    row[0] = row[6] / acreConversionDict.get(Describe(finalHELSummary).SpatialReference.LinearUnitName)
                    # Final_HEL_Value will be set to the initial HEL value
                    row[1] = row[4]
                    # set Final HEL Acres to 0 for PHEL and NHEL; othewise set to polyAcres
                    if row[4] in ('NHEL','PHEL'):
                        row[2] = 0.0
                    else:
                        row[2] = row[0]
                    # Calculate percent of polygon relative to CLU
                    cluAcres = ogCLUinfoDict.get(row[5])[1]
                    pct = (row[0] / cluAcres) * 100
                    if pct > 100.0: pct = 100.0
                    row[3] = pct
                    del cluAcres,pct
                    cursor.updateRow(row)
            del cursor

        # Put this section in a try-except. It will fail if run from ArcCatalog
        mxd = MapDocument("CURRENT")
        df = ListDataFrames(mxd)[0]

        # Workaround:  ListLayers returns a list of layer objects. Need to create a list of layer name Strings
        currentLayersObj = ListLayers(mxd)
        currentLayersStr = [str(x) for x in ListLayers(mxd)]

        # List of layers to add to Arcmap (layer path, arcmap layer name)
        if bNoPHELvalues or bSkipGeoprocessing:
            addToArcMap = [(helSummary,"Initial HEL Summary"),(fieldDetermination,"Field Determination")]
            # Remove these layers from arcmap if they are present since they were not produced
            if 'LiDAR HEL Summary' in currentLayersStr:
                RemoveLayer(df,currentLayersObj[currentLayersStr.index('LiDAR HEL Summary')])
            if 'Final HEL Summary' in currentLayersStr:
                RemoveLayer(df,currentLayersObj[currentLayersStr.index('Final HEL Summary')])
        else:
            addToArcMap = [(lidarHEL,"LiDAR HEL Summary"),(helSummary,"Initial HEL Summary"),(finalHELSummary,"Final HEL Summary"),(fieldDetermination,"Field Determination")]

        for layer in addToArcMap:
            # remove layer from ArcMap if it exists
            if layer[1] in currentLayersStr:
                RemoveLayer(df,currentLayersObj[currentLayersStr.index(layer[1])])
            # Raster Layers need to be handled differently than vector layers
            if layer[1] == "LiDAR HEL Summary":
                rasterLayer = MakeRasterLayer(layer[0],layer[1])
                tempLayer = rasterLayer.getOutput(0)
                AddLayer(df,tempLayer,"TOP")
                # define the symbology layer and convert it to a layer object
                updateLayer = ListLayers(mxd,layer[1], df)[0]
                symbologyLyr = path.join(path.dirname(argv[0]),layer[1].lower().replace(" ","") + ".lyr")
                sourceLayer = Layer(symbologyLyr)
                UpdateLayer(df,updateLayer,sourceLayer)
            else:
                # add layer to arcmap
                symbologyLyr = path.join(path.dirname(argv[0]),layer[1].lower().replace(" ","") + ".lyr")
                AddLayer(df, Layer(symbologyLyr.strip("'")), "TOP")

            # This layer should be turned on if no PHEL values were processed. Symbology should also be updated to reflect current values.
            if layer[1] in ("Initial HEL Summary") and bNoPHELvalues:
                for lyr in ListLayers(mxd, layer[1]):
                    lyr.visible = True

            # these 2 layers should be turned off by default if full processing happens
            if layer[1] in ("Initial HEL Summary","LiDAR HEL Summary") and not bNoPHELvalues:
                for lyr in ListLayers(mxd, layer[1]):
                    lyr.visible = False

            AddMsgAndPrint("Added " + layer[1] + " to your ArcMap Session",0)

        # Unselect CLU polygons; Looks goofy after processed layers have been added to ArcMap. Turn it off as well
        for lyr in ListLayers(mxd, "*" + str(Describe(cluLayer).nameString).split("\\")[-1], df):
            SelectLayerByAttribute(lyr, "CLEAR_SELECTION")
            lyr.visible = False

        # Turn off the original HEL layer to put the outputs into focus
        helLyr = ListLayers(mxd, Describe(helLayer).nameString, df)[0]
        helLyr.visible = False
        # set dataframe extent to the extent of the Field Determintation layer buffered by 50 meters.
        fieldDeterminationBuffer = "in_memory" + sep + path.basename(CreateScratchName("fdBuffer",data_type="FeatureClass",workspace=scratchWS))
        Buffer(fieldDetermination, fieldDeterminationBuffer, "50 Meters", "FULL", "", "ALL", "")
        df.extent = Describe(fieldDeterminationBuffer).extent
        Delete(fieldDeterminationBuffer)
        RefreshTOC()
        RefreshActiveView()

    except:
        errorMsg()


def populateForm():
    """ This function will prepare the 1026 form by adding 16 fields to the fieldDetermination feauture class. This function will
        still be invoked even if there was no geoprocessing executed due to no PHEL values to be computed. If bNoPHELvalues is true,
        3 fields will be added that otherwise would've been added after HEL geoprocessing. However, values to these fields will be
        populated from the ogCLUinfoDict and represent original HEL values primarily determined by the 33.33% or 50 acre rule."""
    try:
        AddMsgAndPrint("\nPreparing and Populating NRCS-CPA-026 Form", 0)
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
            stateCode = ([row[0] for row in SearchCursor(fieldDetermination,"STATECD")])
            state = [stAbbrev for (stAbbrev, code) in list(stateCodeDict.items()) if code == stateCode[0]][0]
        except:
            state = getuser().replace('.',' ').replace('\'','')

        # Get the geographic and admin counties
        stFIPS = ''
        coFIPS = ''
        astFIPS = ''
        acoFIPS = ''

        # check that lookup table exists and search it for geographic and admin state and county codes
        if Exists(lu_table):
            with SearchCursor(fieldDetermination, ['statecd','countycd','admnstate','admncounty']) as cursor:
                for row in cursor:
                    stFIPS = row[0]
                    coFIPS = row[1]
                    astFIPS = row[2]
                    acoFIPS = row[3]
                    break

            # Lookup state and county names from recovered fips codes
            GeoCode = str(stFIPS) + str(coFIPS)
            expression1 = ("{} = '" + GeoCode + "'").format(AddFieldDelimiters(lu_table, 'GEOID'))
            with SearchCursor(lu_table, ['GEOID','NAME','STPOSTAL'], where_clause=expression1) as cursor:
                for row in cursor:
                    GeoCounty = row[1]
                    GeoState = row[2]
                    # We should only get one result if using installed lookup table from US Census Tiger table, so break
                    break

            AdminCode = str(astFIPS) + str(acoFIPS)
            expression2 = ("{} = '" + AdminCode + "'").format(AddFieldDelimiters(lu_table, 'GEOID'))
            with SearchCursor(lu_table, ['GEOID','NAME','STPOSTAL'], where_clause=expression2) as cursor:
                for row in cursor:
                    AdminCounty = row[1]
                    AdminState = row[2]
                    # We should only get one result if using installed lookup table from US Census Tiger table, so break
                    break

            GeoLocation = GeoCounty + ", " + GeoState
            AdminLocation = AdminCounty + ", " + AdminState

        else:
            GeoLocation = "Not Found"
            AdminLocation = "Not Found"

        # Add 20 Fields to the fieldDetermination feature class
        remarks_txt = r'This Highly Erodible Land determination was conducted offsite using the soil survey. If PHEL soil map units were present, they may have been evaluated using elevation data.'
        fieldDict = {"Signature":("TEXT",dcSignature,50),"SoilAvailable":("TEXT","Yes",5),"Completion":("TEXT","Office",10),
                        "SodbustField":("TEXT","No",5),"Delivery":("TEXT","Mail",10),"Remarks":("TEXT",remarks_txt,255),
                        "RequestDate":("DATE",""),"LastName":("TEXT","",50),"FirstName":("TEXT",input_cust,50),"Address":("TEXT","",50),
                        "City":("TEXT","",25),"ZipCode":("TEXT","",10),"Request_from":("TEXT","AD-1026",15),"HELFarm":("TEXT","0",5),
                        "Determination_Date":("DATE",today),"state":("TEXT",state,2),"SodbustTract":("TEXT","No",5),"Lidar":("TEXT","Yes",5),
                        "Location_County":("TEXT",GeoLocation,50),"Admin_County":("TEXT",AdminLocation,50)}

        SetProgressor("step", "Preparing and Populating NRCS-CPA-026 Form", 0, len(fieldDict), 1)

        for field,params in fieldDict.items():
            SetProgressorLabel("Adding Field: " + field + r' to "Field Determination" layer')
            try:
                fldLength = params[2]
            except:
                fldLength = 0
                pass
            AddField(fieldDetermination,field,params[0],"#","#",fldLength)
            if len(params[1]) > 0:
                expression = "\"" + params[1] + "\""
                CalculateField(fieldDetermination,field,expression,"VB")

        if bAccess:
            AddMsgAndPrint("\tOpening NRCS-CPA-026 Form",0)
            try:
                Popen([msAccessPath,helDatabase])
            except:
                try:
                    startfile(helDatabase)
                except:
                    AddMsgAndPrint("\tCould not locate the Microsoft Access Software",1)
                    AddMsgAndPrint("\tOpen Microsoft Access manually to access the NRCS-CPA-026 Form",1)
                    SetProgressorLabel("Could not locate the Microsoft Access Software")
        else:
            AddMsgAndPrint("\tOpening NRCS-CPA-026 Form",0)
            try:
                startfile(helDatabase)
            except:
                AddMsgAndPrint("\tCould not locate the Microsoft Access Software",1)
                AddMsgAndPrint("\tOpen Microsoft Access manually to access the NRCS-CPA-026 Form",1)
                SetProgressorLabel("Could not locate the Microsoft Access Software")

        return True
    except:
        return False


def configLayout():
    """ This function will gather and update information for elements of the map layout"""
    try:
        # Confirm Map Doc existence
        mxd = MapDocument("CURRENT")
        # Set starting variables as empty for logical comparison later on, prior to updating layout
        farmNum = ''
        trNum = ''
        county = ''
        state = ''

        # end function if lookup table from tool parameters does not exist
        if not Exists(lu_table):
            return False

        # Get CLU information from first row of the cluLayer.
        # All CLU records should have this info, so break after one record.
        with SearchCursor(fieldDetermination, ['statecd','countycd','farmnbr','tractnbr']) as cursor:
            for row in cursor:
                stCD = row[0]
                coCD = row[1]
                farmNum = row[2]
                trNum = row[3]
                break

        # Lookup state and county name
        stco_code = str(stCD) + str(coCD)
        expression = ("{} = '" + stco_code + "'").format(AddFieldDelimiters(lu_table, 'GEOID'))
        with SearchCursor(lu_table, ['GEOID','NAME','STPOSTAL'], where_clause=expression) as cursor:
            for row in cursor:
                county = row[1]
                state = row[2]
                # We should only get one result if using installed lookup table from US Census Tiger table, so break
                break

        # Find and hook map layout elements to variables
        for elm in ListLayoutElements(mxd):
            if elm.name == "farm_txt":
                farm_ele = elm
            if elm.name == "tract_txt":
                tract_ele = elm
            if elm.name == "customer_txt":
                customer_ele = elm
            if elm.name == "county_txt":
                county_ele = elm
            if elm.name == "state_txt":
                state_ele = elm #Not used anywhere?

        # Configure the text boxes
        # If any of the info is missing, the script still runs and boxes are reset to a manual entry prompt
        if farmNum != '':
            farm_ele.text = "Farm: " + str(farmNum)
        else:
            farm_ele.text = "Farm: <dbl-click to enter>"
        if trNum != '':
            tract_ele.text = "Tract: " + str(trNum)
        else:
            tract_ele.text = "Tract: <dbl-click to enter>"
        if input_cust != '':
            customer_ele.text = "Customer(s): " + str(input_cust)
        else:
            customer_ele.text = "Customer(s): <dbl-click to enter>"
        if county != '':
            county_ele.text = "County: " + str(county) + ", " + str(state)
        else:
            county_ele.text = "County: <dbl-click to enter>"

        return True
    except:
        return False


if __name__ == '__main__':
    try:
        cluLayer = GetParameter(0)
        helLayer = GetParameter(1)
        inputDEM = GetParameter(2)
        zUnits = GetParameterAsText(3)
        dcSignature = GetParameterAsText(4)
        input_cust = GetParameterAsText(5)
        use_runoff_ls = GetParameter(6)

        kFactorFld = "K"
        tFactorFld = "T"
        rFactorFld = "R"
        helFld = "MUHELCL"

        bLog = False # If True, log to text file
        SetProgressorLabel("Checking input values and environments")
        AddMsgAndPrint("\nChecking input values and environments")

        # Check HEL Access Database
        helDatabase = path.dirname(argv[0]) + sep + r'HEL.mdb'
        if not Exists(helDatabase):
            AddMsgAndPrint("\nHEL Access Database does not exist in the same path as HEL Tools",2)
            exit()
        # Also define the lu_table, but it's still ok to continue if it's not present
        lu_table = path.dirname(argv[0]) + sep + r'census_fips_lut.dbf'

        # Close Microsoft Access Database software if it is open.
        # forcibly kill image name msaccess if open. remove access record-locking information
        try:
            killAccess = system("TASKKILL /F /IM msaccess.exe")
            if killAccess == 0:
                AddMsgAndPrint("\tMicrosoft Access was closed in order to continue")
            accessLockFile = path.dirname(argv[0]) + sep + r'HEL.ldb'
            if path.exists(accessLockFile):
                remove(accessLockFile)
            sleep(2)
        except:
            sleep(2)
            pass

        # Establish path to access database layers
        fieldDetermination = path.join(helDatabase, r'Field_Determination')
        helSummary = path.join(helDatabase, r'Initial_HEL_Summary')
        lidarHEL = path.join(helDatabase, r'LiDAR_HEL_Summary')
        finalHELSummary = path.join(helDatabase, r'Final_HEL_Summary')
        accessLayers = [fieldDetermination,helSummary,lidarHEL,finalHELSummary]
        for layer in accessLayers:
            if Exists(layer):
                try:
                    Delete(layer)
                except:
                    AddMsgAndPrint("\tCould not delete the " + path.basename(layer) + " feature class in the HEL access database. Creating an additional layer",2)
                    newName = str(layer)
                    newName = CreateScratchName(path.basename(layer),data_type="FeatureClass",workspace=helDatabase)

        # Determine Microsoft Access path from windows version
        bAccess = True
        winVersion = getwindowsversion()
        # Windows 10
        if winVersion.build == 9200:
            msAccessPath = r'C:\Program Files (x86)\Microsoft Office\root\Office16\MSACCESS.EXE'
        # Windows 7
        elif winVersion.build == 7601:
            msAccessPath = r'C:\Program Files (x86)\Microsoft Office\Office15\MSACCESS.EXE'
        else:
            AddMsgAndPrint("\nCould not determine Windows version, will not populate 026 Form",2)
            bAccess = False
        if bAccess and not path.isfile(msAccessPath):
            bAccess = False

        # Checkout Spatial Analyst Extension and set scratch workspace
        # TODO: Move this extension check to tool validation
        try:
            if CheckExtension("Spatial") == "Available":
                CheckOutExtension("Spatial")
        except Exception:
            AddMsgAndPrint("\n\nSpatial Analyst license is unavailable.  Go to Customize -> Extensions to activate it",2)
            AddMsgAndPrint("\n\nExiting!")
            exit()

        # Set overwrite option
        env.overwriteOutput = True

        # define and set the scratch workspace
        scratchWS = path.dirname(argv[0]) + sep + r'scratch.gdb'
        if not Exists(scratchWS):
            CreateFileGDB(path.dirname(argv[0]), r'scratch.gdb')

        if not scratchWS:
            AddMsgAndPrint("\nCould Not set scratchWorkspace!")
            exit()

        env.scratchWorkspace = scratchWS
        scratchLayers = list()

        # Stamp CLU into field determination fc. Exit if no CLU fields selected
        cluDesc = Describe(cluLayer)
        if cluDesc.FIDset == '':
            AddMsgAndPrint("\nPlease select fields from the CLU Layer. Exiting!",2)
            exit()
        else:
            fieldDetermination = CopyFeatures(cluLayer,fieldDetermination)

        # Make sure TRACTNBR and FARMNBR  are unique; exit otherwise
        uniqueTracts = list(set([row[0] for row in SearchCursor(fieldDetermination,("TRACTNBR"))]))
        uniqueFarm   = list(set([row[0] for row in SearchCursor(fieldDetermination,("FARMNBR"))]))
        uniqueFields = list(set([row[0] for row in SearchCursor(fieldDetermination,("CLUNBR"))]))

        if len(uniqueTracts) != 1:
           AddMsgAndPrint("\n\tThere are " + str(len(uniqueTracts)) + " different Tract Numbers. Exiting!",2)
           for tract in uniqueTracts:
               AddMsgAndPrint("\t\tTract #: " + str(tract),2)
           removeScratchLayers()
           exit()

        if len(uniqueFarm) != 1:
           AddMsgAndPrint("\n\tThere are " + str(len(uniqueFarm)) + " different Farm Numbers. Exiting!",2)
           for farm in uniqueFarm:
               AddMsgAndPrint("\t\tFarm #: " + str(farm),2)
           removeScratchLayers()
           exit()

        # Create Text file to log info to
        textFilePath = createTextFile(uniqueTracts[0],uniqueFarm[0],uniqueFields)
        bLog = True

        # Update the map layout for the current site being run
        configLayout()

        AddMsgAndPrint("\nNumber of CLU fields selected: {}".format(len(cluDesc.FIDset.split(";"))))

        # Add Calcacre field if it doesn't exist. Should be part of the CLU layer.
        calcAcreFld = "CALCACRES"
        if not len(ListFields(fieldDetermination,calcAcreFld)) > 0:
            AddField(fieldDetermination,calcAcreFld,"DOUBLE")

        # Note: FSA MIDAS uses "square meters * 0.0002471" based on NAD 83 for the current UTM Zone and then rounds to two decimal points to set its calc acres.
        # If we changed all calc acres formulas to match FSA's formula, we would have matching FSA acres, but slightly incorrect amounts.
        # The variance is approximately two hundred thouandths of an acre or about 3/10ths of a square inch per acre.
        # If we set all internal acres computations to 2 decimal places from rounding based on the above, all acres would be consistent, except possibly for raster derived acres (need to check).
        CalculateField(fieldDetermination,calcAcreFld,"!shape.area@acres!","PYTHON_9.3")
        totalAcres = float("%.1f" % (sum([row[0] for row in SearchCursor(fieldDetermination, (calcAcreFld))])))
        AddMsgAndPrint("\tTotal Acres: " + str(totalAcres))

        # Z-factor conversion Lookup table
        # lookup dictionary to convert XY units to area. Key = XY unit of DEM; Value = conversion factor to sq.meters
        acreConversionDict = {'Meter':4046.8564224,'Foot':43560,'Foot_US':43560,'Centimeter':40470000,'Inch':6273000}

        # Assign Z-factor based on XY and Z units of DEM
        # the following represents a matrix of possible z-Factors
        # using different combination of xy and z units
        # ----------------------------------------------------
        #                      Z - Units
        #                       Meter    Foot     Centimeter     Inch
        #          Meter         1	    0.3048	    0.01	    0.0254
        #  XY      Foot        3.28084	  1	      0.0328084	    0.083333
        # Units    Centimeter   100	    30.48	     1	         2.54
        #          Inch        39.3701	  12       0.393701	      1
        # ---------------------------------------------------

        unitLookUpDict = {'Meter':0,'Meters':0,'Foot':1,'Foot_US':1,'Feet':1,'Centimeter':2,'Centimeters':2,'Inch':3,'Inches':3}
        zFactorList = [[1,0.3048,0.01,0.0254],
                       [3.28084,1,0.0328084,0.083333],
                       [100,30.48,1,2.54],
                       [39.3701,12,0.393701,1]]

        # Compute Summary of original HEL values
        # Intersect fieldDetermination (CLU & AOI) with soils (helLayer) -> finalHELSummary
        AddMsgAndPrint("\nComputing summary of original HEL Values")
        SetProgressorLabel("Computing summary of original HEL Values")
        cluHELintersect_pre = "in_memory" + sep + path.basename(CreateScratchName("cluHELintersect_pre",data_type="FeatureClass",workspace=scratchWS))

        # Use the catalog path of the hel layer to avoid using a selection
        helLayerPath = Describe(helLayer).catalogPath

        # Intersect fieldDetermination with soils and explode into single part
        Intersect([fieldDetermination,helLayerPath],cluHELintersect_pre,"ALL")
        MultipartToSinglepart(cluHELintersect_pre,finalHELSummary)
        scratchLayers.append(cluHELintersect_pre)

        # Test intersection --- Should we check the percentage of intersection here? what if only 50% overlap
        # No modification needed for these acres. The total is used only for this check.
        totalIntAcres = sum([row[0] for row in SearchCursor(finalHELSummary, ("SHAPE@AREA"))]) / acreConversionDict.get(Describe(finalHELSummary).SpatialReference.LinearUnitName)
        if not totalIntAcres:
            AddMsgAndPrint("\tThere is no overlap between HEL soil layer and CLU Layer. EXITTING!",2)
            removeScratchLayers()
            exit()

        # Dissolve intersection output by the following fields -> helSummary
        cluNumberFld = "CLUNBR"
        dissovleFlds = [cluNumberFld,"TRACTNBR","FARMNBR","COUNTYCD","CALCACRES",helFld]

        # Dissolve the finalHELSummary to report input summary
        Dissolve(finalHELSummary, helSummary, dissovleFlds, "","MULTI_PART", "DISSOLVE_LINES")

        # Add and Update fields in the HEL Summary Layer (Og_HELcode, Og_HEL_Acres, Og_HEL_AcrePct)
        # Add 3 fields to the intersected layer. The intersected 'clueHELintersect' layer will be used for the dissolve process and at the end of the script.
        HELrasterCode = 'Og_HELcode'    # Used for rasterization purposes
        HELacres = 'Og_HEL_Acres'
        HELacrePct = 'Og_HEL_AcrePct'

        if not len(ListFields(helSummary,HELrasterCode)) > 0:
            AddField(helSummary,HELrasterCode,"SHORT")

        if not len(ListFields(helSummary,HELacres)) > 0:
            AddField(helSummary,HELacres,"DOUBLE")

        if not len(ListFields(helSummary,HELacrePct)) > 0:
            AddField(helSummary,HELacrePct,"DOUBLE")

        # Calculate HELValue Field
        helSummaryDict = dict()     ## tallies acres by HEL value i.e. {PHEL:100}
        nullHEL = 0                 ## # of polygons with no HEL values
        wrongHELvalues = list()     ## Stores incorrect HEL Values
        maxAcreLength = list()      ## Stores the number of acre digits for formatting purposes
        bNoPHELvalues = False       ## Boolean flag to indicate PHEL values are missing

        # HEL Field, Og_HELcode, Og_HEL_Acres, Og_HEL_AcrePct, "SHAPE@AREA", "CALCACRES"
        with UpdateCursor(helSummary,[helFld,HELrasterCode,HELacres,HELacrePct,"SHAPE@AREA",calcAcreFld]) as cursor:
            for row in cursor:
                # Update HEL value field; Continue if NULL HEL value
                if row[0] is None or row[0] == '' or len(row[0]) == 0:
                    nullHEL+=1
                    continue
                elif row[0] == "HEL":
                    row[1] = 0
                elif row[0] == "NHEL":
                    row[1] = 1
                elif row[0] == "PHEL":
                    row[1] = 2
                elif row[0] == "NA":
                    row[1] = 1
                else:
                    if not str(row[0]) in wrongHELvalues:
                        wrongHELvalues.append(str(row[0]))

                # Update Acre field
                # Here we calculated acres differently than we did than when we updated the calc acres in the field determination layer. Seems like we could be consistent here.
                # Differences may be inconsequential if our decimal places match ArcMap's and everything is consistent for coordinate systems for the layers.
                #acres = float("%.1f" % (row[3] / acreConversionDict.get(Describe(helSummary).SpatialReference.LinearUnitName)))
                acres = row[4] / acreConversionDict.get(Describe(helSummary).SpatialReference.LinearUnitName)
                row[2] = acres
                maxAcreLength.append(float("%.1f" %(acres)))

                # Update Pct field
                pct = float("%.2f" %((row[2] / row[5]) * 100)) # HEL acre percentage
                if pct > 100.0: pct = 100.0                    # set pct to 100 if its greater; rounding issue
                row[3] = pct

                # Add hel value to dictionary to summarize by total project
                if row[0] not in helSummaryDict:
                    helSummaryDict[row[0]] = acres
                else:
                    helSummaryDict[row[0]] += acres

                cursor.updateRow(row)
                del acres

        # No PHEL values were found; Bypass geoprocessing and populate form
        if 'PHEL' not in helSummaryDict:
            bNoPHELvalues = True

        # Inform user about NULL values; Exit if any NULLs exist.
        if nullHEL > 0:
            AddMsgAndPrint("\n\tERROR: There are " + str(nullHEL) + " polygon(s) with missing HEL values. EXITING!",2)
            removeScratchLayers()
            exit()

        # Inform user about invalid HEL values (not PHEL,HEL, NHEL); Exit if invalid values exist.
        if wrongHELvalues:
            AddMsgAndPrint("\n\tERROR: There is " + str(len(set(wrongHELvalues))) + " invalid HEL values in HEL Layer:",1)
            for wrongVal in set(wrongHELvalues):
                AddMsgAndPrint("\t\t" + wrongVal)
            removeScratchLayers()
            exit()

        del dissovleFlds,nullHEL,wrongHELvalues

        # Report HEl Layer Summary by field
        AddMsgAndPrint("\n\tSummary by CLU:")

        # Create 2 temporary tables to capture summary statistics
        ogHelSummaryStats = "in_memory" + sep + path.basename(CreateScratchName("ogHELSummaryStats",data_type="ArcInfoTable",workspace=scratchWS))
        ogHelSummaryStatsPivot = "in_memory" + sep + path.basename(CreateScratchName("ogHELSummaryStatsPivot",data_type="ArcInfoTable",workspace=scratchWS))

        stats = [[HELacres,"SUM"]]
        caseField = [cluNumberFld,helFld]
        Statistics(helSummary, ogHelSummaryStats, stats, caseField)
        sumHELacreFld = [fld.name for fld in ListFields(ogHelSummaryStats,"*" + HELacres)][0]
        scratchLayers.append(ogHelSummaryStats)

        # Pivot table will have CLUNBR & any HEL values present (HEL,NHEL,PHEL)
        PivotTable(ogHelSummaryStats,cluNumberFld,helFld,sumHELacreFld,ogHelSummaryStatsPivot)
        scratchLayers.append(ogHelSummaryStatsPivot)

        pivotFields = [fld.name for fld in ListFields(ogHelSummaryStatsPivot)][1:]  # ['CLUNBR','HEL','NHEL','PHEL']
        numOfhelValues = len(pivotFields)                                                 # Number of Pivot table fields; Min 2 fields
        maxAcreLength.sort(reverse=True)
        bSkipGeoprocessing = True             # Skip processing until a field is neither HEL >= 33.33% or NHEL > 66.67%

        # This dictionary will only be used if FINAL results are all HEL or all NHEL to reference original
        # acres and not use tabulate area acres.  It will also be used when there are no PHEL Values.
        # {cluNumber:(HEL value, cluAcres, HEL Pct} -- HEL value is determined by the 33.33% or 50 acre rule
        ogCLUinfoDict = dict()

        # Iterate through the pivot table and report HEL values by CLU - ['CLUNBR','HEL','NHEL','PHEL']
        with SearchCursor(ogHelSummaryStatsPivot,pivotFields) as cursor:
            for row in cursor:

                og_cluHELrating = None         # original field HEL Rating
                og_cluHELacresList = list()    # temp list of acres by HEL value
                og_cluHELpctList = list()      # temp list of pct by HEL value
                msgList = list()               # temp list of messages to print
                cluAcres = sum([row[i] for i in range(1,numOfhelValues,1)])

                # strictly to determine if geoprocessing is needed
                bHELgreaterthan33 = False
                bNHELgreaterthan66 = False

                # iterate through the pivot table fields by record
                for i in range(1,numOfhelValues,1):
                    acres =  float("%.1f" % (row[i]))
                    pct = float("%.1f" % ((row[i] / cluAcres) * 100))

                    # set pct to 100 if its greater; rounding issue
                    if pct > 100.0: pct = 100.0

                    # Determine HEL rating of original fields and populate acres
                    # and pc into ogCLUinfoDict.  Primarily for bNoPHELvalues.
                    # Also determine if further geoProcessing is needed.
                    if og_cluHELrating == None:

                        # Set field to HEL
                        if pivotFields[i] == "HEL" and (pct >= 33.33 or acres >= 50):
                            og_cluHELrating = "HEL"
                            if not row[0] in ogCLUinfoDict:
                                ogCLUinfoDict[row[0]] = (og_cluHELrating,cluAcres,pct)
                            bHELgreaterthan33 = True

                        # Set field to NHEL
                        elif pivotFields[i] == "NHEL" and pct > 66.67:
                            bNHELgreaterthan66 = True
                            og_cluHELrating = "NHEL"
                            if not row[0] in ogCLUinfoDict:
                                ogCLUinfoDict[row[0]] = (og_cluHELrating,cluAcres,pct)

                        # This is the last field in the pivot table
                        elif i == (numOfhelValues - 1):
                            og_cluHELrating = pivotFields[i]
                            if not row[0] in ogCLUinfoDict:
                                ogCLUinfoDict[row[0]] = (og_cluHELrating,cluAcres,pct)

                        # First field did not meet HEL criteria; add it to a temp list
                        else:
                            og_cluHELacresList.append(row[i])
                            og_cluHELpctList.append(pct)

                    # Formulate messages but don't print yet
                    firstSpace = " " * (4-len(pivotFields[i]))                                    # PHEL has 4 characters
                    secondSpace = " " * (len(str(maxAcreLength[0])) - len(str(acres)))            # Number of spaces
                    msgList.append(str("\t\t\t" + pivotFields[i] + firstSpace + " -- " + str(acres) + secondSpace + " .ac -- " + str(pct) + " %"))
                    del acres,pct,firstSpace,secondSpace

                # Skip geoprocessing if HEL >=33.33% or NHEL > 66.67%
                if bSkipGeoprocessing:
                    if not bHELgreaterthan33 and not bNHELgreaterthan66:
                        bSkipGeoprocessing = False

                # Report messages to user; og CLU HEL rating will be reported if bNoPHELvalues is true.
                if bNoPHELvalues:
                    AddMsgAndPrint("\n\t\tCLU #: " + str(row[0]) + " - Rating: " + og_cluHELrating)
                else:
                    AddMsgAndPrint("\n\t\tCLU #: " + str(row[0]))
                for msg in msgList:
                    AddMsgAndPrint(msg)

                del og_cluHELrating,og_cluHELacresList,og_cluHELpctList,msgList,cluAcres

        del stats,caseField,sumHELacreFld,pivotFields,numOfhelValues,maxAcreLength

        # No PHEL Values Found
        # If there are no PHEL Values add helSummary and fieldDetermination layers to ArcMap and prepare 1026 form. Skip geoprocessing.
        if bNoPHELvalues or bSkipGeoprocessing:
            if bNoPHELvalues:
               AddMsgAndPrint("\n\tThere are no PHEL values in HEL layer",1)
               AddMsgAndPrint("\tNo Geoprocessing is required.")

            # Only Print this if there are PHEL values but they don't need
            # to be processed; Otherwise it should be captured by above statement.
            if bSkipGeoprocessing and not bNoPHELvalues:
               AddMsgAndPrint("\n\tHEL values are >= 33.33% or more than 50 acres, or NHEL values are > 66.67%",1)
               AddMsgAndPrint("\tNo Geoprocessing is required.\n")

            AddLayersToArcMap()

            if not populateForm():
                AddMsgAndPrint("\nFailed to correclty populate NRCS-CPA-026 form",2)

            # Clean up time
            SetProgressorLabel("")
            AddMsgAndPrint("\n")
            RefreshCatalog(scratchWS)
            exit()

        # Check and create DEM clip from buffered CLU
        # Exit if a DEM is not present; At this point PHEL mapunits are present and requires a DEM to process them.
        try:
            Describe(inputDEM).baseName
        except:
            AddMsgAndPrint("\nDEM is required to process PHEL values. EXITING!")
            exit()

        units,zFactor,dem = extractDEM(inputDEM,zUnits)
        if not zFactor or not dem:
            removeScratchLayers()
            exit()
        scratchLayers.append(dem)

        # Check DEM for NoData overlaps with input CLU fields
        AddMsgAndPrint("\nChecking input DEM for site coverage...")
        vectorNull = "in_memory" + sep + path.basename(CreateScratchName("vectorNull",data_type="FeatureClass",workspace=scratchWS))
        demCheck = "in_memory" + sep + path.basename(CreateScratchName("demCheck",data_type="FeatureClass",workspace=scratchWS))

        # Use Set Null statement to change things with value of 0 to NoData
        whereClause = "VALUE = 0"
        setNull = SetNull(dem, dem, whereClause)
        scratchLayers.append(SetNull)

        # Use IsNull to convert NoData values in the DEM to 1 and all other values to 0
        demNull = IsNull(setNull)
        scratchLayers.append(demNull)

        # Convert the IsNull raster to a vector layer
        try:
            RasterToPolygon(demNull, vectorNull, "SIMPLIFY", "Value", "MULTIPLE_OUTER_PART")
        except:
            RasterToPolygon(demNull, vectorNull, "SIMPLIFY", "Value")
        scratchLayers.append(vectorNull)

        # Clip the IsNull vector layer by the field determination layer
        Clip_a(vectorNull, fieldDetermination, demCheck)
        scratchLayers.append(demCheck)

        # Search for any values of 1 in the demCheck layer and issue a warning to the user if present
        fields = ['gridcode']
        cursor = SearchCursor(demCheck, fields)
        nd_warning = False
        for row in cursor:
            if row[0] == 1:
                nd_warning = True

        # If no data warning is True, show error messages
        if nd_warning == True:
            AddMsgAndPrint("\nInput DEM problem detected! The input DEM may have null data areas covering the input CLU fields!",2)
            AddMsgAndPrint("\nPHEL map unit and slope analysis is likely to be invalid!",2)
            AddMsgAndPrint("\nPlease review input DEM data for actual coverage over the site.",2)
            AddMsgAndPrint("\nIf input DEM does not cover the site, the determination must be made with traditional methods.",2)
        else:
            AddMsgAndPrint("\nDEM values in site extent are not null. Continuing...")

        del fields, cursor, nd_warning

        # Create Slope Layer
        # Perform a minor fill to reduce LiDAR data noise and minor irregularities. Try to use a max fill height of no more than 1 foot, based on input zUnits.
        SetProgressorLabel("Filling small sinks in DEM")
        AddMsgAndPrint("\nFilling small sinks in DEM")
        if zUnits == "Feet":
            zLimit = 1
        elif zUnits == "Meters":
            zLimit = 0.3048
        elif zUnits == "Inches":
            zLimit = 12
        elif zUnits == "Centimeters":
            zLimit = 30.48
        else:
            # Assume worst case z units of Meters
            zLimit = 0.3048

        # 1 Perform the fill using the zLimit as the max fill amount
        filled = Fill(dem, zLimit)
        scratchLayers.append(filled)

        # 2 Run a FocalMean to smooth the DEM of LiDAR data noise. This should be run prior to creating derivative products.
        # This replaces running FocalMean on the slope layer itself.
        SetProgressorLabel("Running Focal Statistics on DEM")
        AddMsgAndPrint("Running Focal Statistics on DEM")
        preslope = FocalStatistics(filled, NbrRectangle(3,3,"CELL"),"MEAN","DATA")

        # 3 Create Slope
        SetProgressorLabel("Creating Slope Derivative")
        AddMsgAndPrint("\nCreating Slope Derivative")
        slope = Slope(preslope,"PERCENT_RISE",zFactor)

        # 4 Create Flow Direction and Flow Length
        SetProgressorLabel("Calculating Flow Direction")
        AddMsgAndPrint("Calculating Flow Direction")
        flowDirection = FlowDirection(preslope, "FORCE")
        scratchLayers.append(flowDirection)

        # 5 Calculate Flow Length
        SetProgressorLabel("Calculating Flow Length")
        AddMsgAndPrint("Calculating Flow Length")
        preflowLength = FlowLength(flowDirection,"UPSTREAM", "")
        scratchLayers.append(preflowLength)

        # 6 Run a focal statistics on flow length
        SetProgressorLabel("Running Focal Statistics on Flow Length")
        AddMsgAndPrint("Running Focal Statistics on Flow Length")
        flowLength = FocalStatistics(preflowLength, NbrRectangle(3,3,"CELL"),"MAXIMUM","DATA")
        scratchLayers.append(flowLength)

        # 7 convert Flow Length distance units to feet if original DEM LINEAR UNITS ARE not in feet.
        if not units in ('Feet','Foot','Foot_US'):
            AddMsgAndPrint("Converting Flow Length Distance units to Feet")
            flowLengthFT = flowLength * 3.280839896
            scratchLayers.append(flowLengthFT)
        else:
            flowLengthFT = flowLength
            scratchLayers.append(flowLengthFT)

        # 8 Convert slope percent to radians for use in various LS equations
        radians = ATan(Times(slope,0.01))

        # Compute LS Factor
        # If Northwest US 'Use Runoff LS Equation' flag was active, use the following equation
        if use_runoff_ls:
            SetProgressorLabel("Calculating LS Factor")
            AddMsgAndPrint("Calculating LS Factor")
            lsFactor = (Power((flowLengthFT/72.6)*Cos(radians),0.5))*(Power(Sin((radians))/(Sin(5.143*((pi)/180))),0.7))

        # Otherwise, use the standard AH537 LS computation
        else:
            # 9 Calculate S Factor
            SetProgressorLabel("Calculating S Factor")
            AddMsgAndPrint("\nCalculating S Factor")
            # Compute S factor using formula in AH537, pg 12
            sFactor = ((Power(Sin(radians),2)*65.41)+(Sin(radians)*4.56)+(0.065))
            scratchLayers.append(sFactor)

            # 10 Calculate L Factor
            SetProgressorLabel("Calculating L Factor")
            AddMsgAndPrint("Calculating L Factor")

            # Original outlFactor lines
            """outlFactor = Con(Raster(slope),Power(Raster(flowLengthFT) / 72.6,0.2),
                               Con(Raster(slope),Power(Raster(flowLengthFT) / 72.6,0.3),
                               Con(Raster(slope),Power(Raster(flowLengthFT) / 72.6,0.4),
                               Power(Raster(flowLengthFT) / 72.6,0.5),"VALUE >= 3 AND VALUE < 5"),"VALUE >= 1 AND VALUE < 3"),"VALUE<1")"""

            # Remove 'Raster' function from above
            lFactor = Con(slope,Power(flowLengthFT / 72.6,0.2),
                            Con(slope,Power(flowLengthFT / 72.6,0.3),
                            Con(slope,Power(flowLengthFT / 72.6,0.4),
                            Power(flowLengthFT / 72.6,0.5),"VALUE >= 3 AND VALUE < 5"),"VALUE >= 1 AND VALUE < 3"),"VALUE<1")

            scratchLayers.append(lFactor)

            # 11 Calculate LS Factor "%l_factor%" * "%s_factor%"
            SetProgressorLabel("Calculating LS Factor")
            AddMsgAndPrint("Calculating LS Factor")
            lsFactor = lFactor * sFactor

        scratchLayers.append(radians)
        scratchLayers.append(lsFactor)

        # Convert K,T & R Factor and HEL Value to Rasters
        AddMsgAndPrint("\nConverting Vector to Raster for Spatial Analysis Purpose")
        cellSize = Describe(dem).MeanCellWidth

        # This works in 10.5 AND works in 10.6.1 and 10.7 but slows processing
        kFactor = CreateScratchName("kFactor",data_type="RasterDataset",workspace=scratchWS)
        tFactor = CreateScratchName("tFactor",data_type="RasterDataset",workspace=scratchWS)
        rFactor = CreateScratchName("rFactor",data_type="RasterDataset",workspace=scratchWS)
        helValue = CreateScratchName("helValue",data_type="RasterDataset",workspace=scratchWS)

        # 12 Convert KFactor to raster
        SetProgressorLabel("Converting K Factor field to a raster")
        AddMsgAndPrint("\tConverting K Factor field to a raster")
        FeatureToRaster(finalHELSummary,kFactorFld,kFactor,cellSize)

        # 13 Convert TFactor to raster
        SetProgressorLabel("Converting T Factor field to a raster")
        AddMsgAndPrint("\tConverting T Factor field to a raster")
        FeatureToRaster(finalHELSummary,tFactorFld,tFactor,cellSize)

        # 14 Convert RFactor to raster
        SetProgressorLabel("Converting R Factor field to a raster")
        AddMsgAndPrint("\tConverting R Factor field to a raster")
        FeatureToRaster(finalHELSummary,rFactorFld,rFactor,cellSize)

        SetProgressorLabel("Converting HEL Value field to a raster")
        AddMsgAndPrint("\tConverting HEL Value field to a raster")
        FeatureToRaster(helSummary,HELrasterCode,helValue,cellSize)

        scratchLayers.append(kFactor)
        scratchLayers.append(tFactor)
        scratchLayers.append(rFactor)
        scratchLayers.append(helValue)

        # Calculate EI Factor
        SetProgressorLabel("Calculating EI Factor")
        AddMsgAndPrint("\nCalculating EI Factor")
        eiFactor = Divide((lsFactor * kFactor * rFactor),tFactor)
        scratchLayers.append(eiFactor)

        # Calculate Final HEL Factor
        # Create Conditional statement to reflect the following:
        # 1) PHEL Value = 0 -- Take EI factor -- Depends     2
        # 2) HEL Value  = 1 -- Assign 9                      0
        # 3) NHEL Value = 2 -- Assign 2 (No action needed)   1
        # Anything above 8 is HEL

        SetProgressorLabel("Calculating HEL Factor")
        AddMsgAndPrint("Calculating HEL Factor")
        helFactor = Con(helValue,eiFactor,Con(helValue,9,helValue,"VALUE=0"),"VALUE=2")
        scratchLayers.append(helFactor)

        # Reclassify values:
        #       < 8 = Value_1 = NHEL
        #       > 8 = Value_2 = HEL
        remapString = "0 8 1;8 100000000 2"
        Reclassify_3d(helFactor, "VALUE", remapString, lidarHEL,'NODATA')

        # Determine if individual PHEL delineations are HEL/NHEL"""
        SetProgressorLabel("Computing summary of LiDAR HEL Values:")
        AddMsgAndPrint("\nComputing summary of LiDAR HEL Values:\n")

        # Summarize new values between HEL soil polygon and lidarHEL raster
        outPolyTabulate = "in_memory" + sep + path.basename(CreateScratchName("HEL_Polygon_Tabulate",data_type="ArcInfoTable",workspace=scratchWS))
        zoneFld = Describe(finalHELSummary).OIDFieldName
        TabulateArea(finalHELSummary,zoneFld,lidarHEL,"VALUE",outPolyTabulate,cellSize)
        tabulateFields = [fld.name for fld in ListFields(outPolyTabulate)][2:]
        scratchLayers.append(outPolyTabulate)

        # Add 4 fields to Final HEL Summary layer
        newFields = ['Polygon_Acres','Final_HEL_Value','Final_HEL_Acres','Final_HEL_Percent']
        for fld in newFields:
            if not len(ListFields(finalHELSummary,fld)) > 0:
               if fld == 'Final_HEL_Value':
                  AddField(finalHELSummary,'Final_HEL_Value',"TEXT","","",5)
               else:
                    AddField(finalHELSummary,fld,"DOUBLE")

        # In some cases, the finalHELSummary layer's OID field name was "OBJECTID_1" which
        # conflicted with the output of the tabulate area table.
        if zoneFld.find("_") > -1:
            outputJoinFld = zoneFld
        else:
            outputJoinFld = zoneFld + "_1"
        JoinField(finalHELSummary,zoneFld,outPolyTabulate,outputJoinFld,tabulateFields)

        # Booleans to indicate if only HEL or only NHEL is present
        bOnlyHEL = False; bOnlyNHEL = False

        # Check if VALUE_1(NHEL) or VALUE_2(HEL) are missing from outPolyTabulate table
        finalHELSummaryFlds = [fld.name for fld in ListFields(finalHELSummary)][2:]
        if len(finalHELSummaryFlds):

            # NHEL is not Present - so All is HEL; All is VALUE2
            if not "VALUE_1" in tabulateFields:
                AddMsgAndPrint("\tWARNING: Entire Area is HEL",1)
                AddField(finalHELSummary,"VALUE_1","DOUBLE")
                CalculateField(finalHELSummary,"VALUE_1",0)
                bOnlyHEL = True

            # HEL is not Present - All is NHEL; All is VALUE1
            if not "VALUE_2" in tabulateFields:
                AddMsgAndPrint("\tWARNING: Entire Area is NHEL",1)
                AddField(finalHELSummary,"VALUE_2","DOUBLE")
                CalculateField(finalHELSummary,"VALUE_2",0)
                bOnlyNHEL = True
        else:
            AddMsgAndPrint("\n\tReclassifying helFactor Failed",2)
            exit()

        newFields.append("VALUE_2")
        newFields.append("SHAPE@AREA")
        newFields.append(cluNumberFld)

        # this will be used for field determination
        fieldDeterminationDict = dict()

        # [polyAcres,finalHELvalue,finalHELacres,finalHELpct,"VALUE_2","SHAPE@AREA","CLUNBR"]
        with UpdateCursor(finalHELSummary,newFields) as cursor:
            for row in cursor:
                # Calculate polygon acres
                row[0] = row[5] / acreConversionDict.get(Describe(finalHELSummary).SpatialReference.LinearUnitName)
                # Convert "VALUE_2" values to acres.  Represent acres from a poly that is HEL.
                # The intersection of CLU and soils may cause slivers below the tabulate cell size
                # which will create NULLs.  Set these slivers to 0 acres.
                try:
                    row[2] = row[4] / acreConversionDict.get(Describe(finalHELSummary).SpatialReference.LinearUnitName)
                except:
                    row[2] = 0

                # Calculate percentage of the polygon that is HEL
                row[3] = (row[2] / row[0]) * 100

                # set pct to 100 if its greater; rounding issue
                if row[3] > 100.0: row[3] = 100.0

                # polygon HEL Pct is greater than 50%; HEL
                if row[3] > 50.0:
                    row[1] = "HEL"
                    # Add the HEL polygon acres to the dict
                    if not row[6] in fieldDeterminationDict:
                        fieldDeterminationDict[row[6]] = row[0]
                    else:
                        fieldDeterminationDict[row[6]] += row[0]

                # polygon HEL Pct is less than 50%; NHEL
                else:
                    row[1] = "NHEL"
                    # Don't Add NHEL polygon acres to dict but place
                    # holder for the clu
                    if not row[6] in fieldDeterminationDict:
                        fieldDeterminationDict[row[6]] = 0

                cursor.updateRow(row)

        # Delete unwanted fields from the finalHELSummary Layer
        newFields.remove("VALUE_2")
        validFlds = [cluNumberFld,"STATECD","TRACTNBR","FARMNBR","COUNTYCD","CALCACRES",helFld,"MUSYM","MUNAME","MUWATHEL","MUWNDHEL"] + newFields

        deleteFlds = list()
        for fld in [f.name for f in ListFields(finalHELSummary)]:
            if fld in (zoneFld,'Shape_Area','Shape_Length','Shape'):continue
            if not fld in validFlds:
                deleteFlds.append(fld)

        DeleteField(finalHELSummary,deleteFlds)
        del zoneFld,finalHELSummaryFlds,tabulateFields,newFields,validFlds

        # Determine if field is HEL/NHEL. Add 3 fields to fieldDetermination layer
        fieldList = ["HEL_YES","HEL_Acres","HEL_Pct"]
        for field in fieldList:
            if not FindField(fieldDetermination,field):
                if field == "HEL_YES":
                    AddField(fieldDetermination,field,"TEXT","","",5)
                else:
                    AddField(fieldDetermination,field,"FLOAT")

        fieldList.append(cluNumberFld)
        fieldList.append(calcAcreFld)
        cluDict = dict()  # Strictly for formatting; ClUNBR: (len of clu, helAcres, helPct, len of Acres, len of pct,is it HEL?)

        # ['HEL_YES','HEL_Acres','HEL_Pct','CLUNBR','CALCACRES']
        with UpdateCursor(fieldDetermination,fieldList) as cursor:
            for row in cursor:
                # if results are completely HEL or NHEL then get total clu acres from ogCLUinfoDict
                if bOnlyHEL or bOnlyNHEL:
                    if bOnlyHEL:
                        helAcres = ogCLUinfoDict.get(row[3])[1]
                        nhelAcres = 0.0
                        helPct = 100.0
                        nhelPct = 0.0
                    else:
                        nhelAcres = ogCLUinfoDict.get(row[3])[1]
                        helAcres = 0.0
                        helPct = 0.0
                        nhelPct = 100.0
                else:
                    helAcres = fieldDeterminationDict[row[3]]    # total HEL acres for field
                    helPct = (helAcres / row[4]) * 100           # helAcres / CALCACRES
                    nhelAcres = row[4] - helAcres
                    nhelPct = 100 - helPct
                    # set pct to 100 if its greater; rounding issue
                    if helPct > 100.0: helPct = 100.0
                    if nhelPct > 100.0: nhelPct = 100.0

                clu = row[3]

                if helPct >= 33.33 or helAcres > 50.0:
                    row[0] = "HEL"
                else:
                    row[0] = "NHEL"

                row[1] = helAcres
                row[2] = helPct

                helAcres = float("%.1f" %(helAcres))   # Strictly for formatting
                helPct = float("%.1f" %(helPct))       # Strictly for formatting
                nhelAcres = float("%.1f" %(nhelAcres)) # Strictly for formatting
                nhelPct = float("%.1f" %(nhelPct))     # Strictly for formatting

                cluDict[clu] = (helAcres,len(str(helAcres)),helPct,nhelAcres,len(str(nhelAcres)),nhelPct,row[0]) #  {8: (25.3, 4, 45.1, 30.8, 4, 54.9, 'HEL')}
                del helAcres,helPct,nhelAcres,nhelPct,clu

                cursor.updateRow(row)
        del cursor

        # Strictly for formatting and printing
        maxHelAcreLength = sorted([cluinfo[1] for clu,cluinfo in cluDict.items()],reverse=True)[0]
        maxNHelAcreLength = sorted([cluinfo[4] for clu,cluinfo in cluDict.items()],reverse=True)[0]

        for clu in sorted(cluDict.keys()):
            firstSpace = " "  * (maxHelAcreLength - cluDict[clu][1])
            secondSpace = " " * (maxNHelAcreLength - cluDict[clu][4])
            helAcres = cluDict[clu][0]
            helPct = cluDict[clu][2]
            nHelAcres = cluDict[clu][3]
            nHelPct = cluDict[clu][5]
            yesOrNo = cluDict[clu][6]
            AddMsgAndPrint("\tCLU #: " + str(clu))
            AddMsgAndPrint("\t\tHEL Acres:  " + str(helAcres) + firstSpace + " .ac -- " + str(helPct) + " %")
            AddMsgAndPrint("\t\tNHEL Acres: " + str(nHelAcres) + secondSpace + " .ac -- " + str(nHelPct) + " %")
            AddMsgAndPrint("\t\tHEL Determination: " + yesOrNo + "\n")
            del firstSpace,secondSpace,helAcres,helPct,nHelAcres,nHelPct,yesOrNo

        del fieldList,cluDict,maxHelAcreLength,maxNHelAcreLength

        # Prepare Symboloby for ArcMap and 1026 form
        AddLayersToArcMap()

        if not populateForm():
            AddMsgAndPrint("\nFailed to correclty populate NRCS-CPA-026 form",2)

        # Clean up time
        removeScratchLayers()
        SetProgressorLabel("")
        AddMsgAndPrint("\n")
        RefreshCatalog(scratchWS)

    except:
        removeScratchLayers()
        errorMsg()
