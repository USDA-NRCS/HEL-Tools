# ==========================================================================================
# Name:   Merge Local DEMs by CLU
#
# Author(s): Chris.Morse
#            IN State GIS Specialist
# e-mail:    chris.morse@usda.gov
# phone:     317.295.5849

#            Adolfo.Diaz
#            GIS Specialist
#            National Soil Survey Center
#            USDA - NRCS
# e-mail:    adolfo.diaz@usda.gov
# phone:     608.662.4422 ext. 216

# Contributor: Christiane Roy
#              MN Area GIS Specialist
# e-mail:      christian.roy@usda.gov
# phone:       507.405.3580

# Maintenance Point of Contact: T.B.D.


## ===================================================================================
def AddMsgAndPrint(msg, severity=0):
    # prints message to screen if run as a python script
    # Adds tool message to the geoprocessor
    #
    #Split the message on \n first, so that if it's multiple lines, a GPMessage will be added for each line
    try:
        #print msg

        #for string in msg.split('\n'):
            #Add a geoprocessing message (in case this is run as a tool)
        if severity == 0:
            arcpy.AddMessage(msg)

        elif severity == 1:
            arcpy.AddWarning(msg)

        elif severity == 2:
            arcpy.AddError(msg)

    except:
        pass

## ===================================================================================
def errorMsg():
# Print traceback exceptions.  If sys.exit was trapped by default exception then
# ignore traceback message.

    try:
        exc_type, exc_value, exc_traceback = sys.exc_info()
        theMsg = "\t" + traceback.format_exception(exc_type, exc_value, exc_traceback)[1] + "\n\t" + traceback.format_exception(exc_type, exc_value, exc_traceback)[-1]

        if theMsg.find("sys.exit") > -1:
            AddMsgAndPrint("\n\n")
            pass
        else:
            AddMsgAndPrint(theMsg,2)

    except:
        AddMsgAndPrint("Unhandled error in errorMsg method", 2)
        pass
## ===================================================================================
def setScratchWorkspace():
    """ This function will set the scratchWorkspace for the interim of the execution
        of this tool.  The scratchWorkspace is used to set the scratchGDB which is
        where all of the temporary files will be written to.  The path of the user-defined
        scratchWorkspace will be compared to existing paths from the user's system
        variables.  If there is any overlap in directories the scratchWorkspace will
        be set to C:\TEMP, assuming C:\ is the system drive.  If all else fails then
        the packageWorkspace Environment will be set as the scratchWorkspace. This
        function returns the scratchGDB environment which is set upon setting the scratchWorkspace"""

    try:
        AddMsgAndPrint("\nSetting Scratch Workspace")
        scratchWK = arcpy.env.scratchWorkspace

        # -----------------------------------------------
        # Scratch Workspace is defined by user or default is set
        if scratchWK is not None:

            # dictionary of system environmental variables
            envVariables = os.environ

            # get the root system drive
            if 'SYSTEMDRIVE' in envVariables:
                sysDrive = envVariables['SYSTEMDRIVE']
            else:
                sysDrive = None

            varsToSearch = ['ESRI_OS_DATADIR_LOCAL_DONOTUSE','ESRI_OS_DIR_DONOTUSE','ESRI_OS_DATADIR_MYDOCUMENTS_DONOTUSE',
                            'ESRI_OS_DATADIR_ROAMING_DONOTUSE','TEMP','LOCALAPPDATA','PROGRAMW6432','COMMONPROGRAMFILES','APPDATA',
                            'USERPROFILE','PUBLIC','SYSTEMROOT','PROGRAMFILES','COMMONPROGRAMFILES(X86)','ALLUSERSPROFILE']

            """ This is a printout of my system environmmental variables - Windows 7
            -----------------------------------------------------------------------------------------
            ESRI_OS_DATADIR_LOCAL_DONOTUSE C:\\Users\adolfo.diaz\AppData\Local\
            ESRI_OS_DIR_DONOTUSE C:\\Users\ADOLFO~1.DIA\AppData\Local\Temp\6\arc3765\
            ESRI_OS_DATADIR_MYDOCUMENTS_DONOTUSE C:\\Users\adolfo.diaz\Documents\
            ESRI_OS_DATADIR_COMMON_DONOTUSE C:\ProgramData\
            ESRI_OS_DATADIR_ROAMING_DONOTUSE C:\\Users\adolfo.diaz\AppData\Roaming\
            TEMP C:\\Users\ADOLFO~1.DIA\AppData\Local\Temp\6\arc3765\
            LOCALAPPDATA C:\\Users\adolfo.diaz\AppData\Local
            PROGRAMW6432 C:\Program Files
            COMMONPROGRAMFILES :  C:\Program Files (x86)\Common Files
            APPDATA C:\\Users\adolfo.diaz\AppData\Roaming
            USERPROFILE C:\\Users\adolfo.diaz
            PUBLIC C:\\Users\Public
            SYSTEMROOT :  C:\Windows
            PROGRAMFILES :  C:\Program Files (x86)
            COMMONPROGRAMFILES(X86) :  C:\Program Files (x86)\Common Files
            ALLUSERSPROFILE :  C:\ProgramData
            ------------------------------------------------------------------------------------------"""

            bSetTempWorkSpace = False

            """ Iterate through each Environmental variable; If the variable is within the 'varsToSearch' list
                list above then check their value against the user-set scratch workspace.  If they have anything
                in common then switch the workspace to something local  """
            for var in envVariables:

                if not var in varsToSearch:
                    continue

                # make a list from the scratch and environmental paths
                varValueList = (envVariables[var].lower()).split(os.sep)          # ['C:', 'Users', 'adolfo.diaz', 'AppData', 'Local']
                scratchWSList = (scratchWK.lower()).split(os.sep)                 # [u'C:', u'Users', u'adolfo.diaz', u'Documents', u'ArcGIS', u'Default.gdb', u'']

                # remove any blanks items from lists
                if '' in varValueList: varValueList.remove('')
                if '' in scratchWSList: scratchWSList.remove('')

                # First element is the drive letter; remove it if they are
                # the same otherwise review the next variable.
                if varValueList[0] == scratchWSList[0]:
                    scratchWSList.remove(scratchWSList[0])
                    varValueList.remove(varValueList[0])

                # obtain a similarity ratio between the 2 lists above
                #sM = SequenceMatcher(None,varValueList,scratchWSList)

                # Compare the values of 2 lists; order is significant
                common = [i for i, j in zip(varValueList, scratchWSList) if i == j]

                if len(common) > 0:
                    bSetTempWorkSpace = True
                    break

            # The current scratch workspace shares 1 or more directory paths with the
            # system env variables.  Create a temp folder at root
            if bSetTempWorkSpace:
                AddMsgAndPrint("\tCurrent Workspace: " + scratchWK,0)

                if sysDrive:
                    tempFolder = sysDrive + os.sep + "TEMP"

                    if not os.path.exists(tempFolder):
                        os.makedirs(tempFolder,mode=777)

                    arcpy.env.scratchWorkspace = tempFolder
                    AddMsgAndPrint("\tTemporarily setting scratch workspace to: " + arcpy.env.scratchGDB,1)

                else:
                    packageWS = [f for f in arcpy.ListEnvironments() if f=='packageWorkspace']
                    if arcpy.env[packageWS[0]]:
                        arcpy.env.scratchWorkspace = arcpy.env[packageWS[0]]
                        AddMsgAndPrint("\tTemporarily setting scratch workspace to: " + arcpy.env.scratchGDB,1)
                    else:
                        AddMsgAndPrint("\tCould not set any scratch workspace",2)
                        return False

            # user-set workspace does not violate system paths; Check for read/write
            # permissions; if write permissions are denied then set workspace to TEMP folder
            else:
                arcpy.env.scratchWorkspace = scratchWK

                if arcpy.env.scratchGDB == None:
                    AddMsgAndPrint("\tCurrent scratch workspace: " + scratchWK + " is READ only!",0)

                    if sysDrive:
                        tempFolder = sysDrive + os.sep + "TEMP"

                        if not os.path.exists(tempFolder):
                            os.makedirs(tempFolder,mode=777)

                        arcpy.env.scratchWorkspace = tempFolder
                        AddMsgAndPrint("\tTemporarily setting scratch workspace to: " + arcpy.env.scratchGDB,1)

                    else:
                        packageWS = [f for f in arcpy.ListEnvironments() if f=='packageWorkspace']
                        if arcpy.env[packageWS[0]]:
                            arcpy.env.scratchWorkspace = arcpy.env[packageWS[0]]
                            AddMsgAndPrint("\tTemporarily setting scratch workspace to: " + arcpy.env.scratchGDB,1)

                        else:
                            AddMsgAndPrint("\tCould not set any scratch workspace",2)
                            return False

                else:
                    AddMsgAndPrint("\tUser-defined scratch workspace is set to: "  + arcpy.env.scratchGDB,0)

        # No workspace set (Very odd that it would go in here unless running directly from python)
        else:
            AddMsgAndPrint("\tNo user-defined scratch workspace ",0)
            sysDrive = os.environ['SYSTEMDRIVE']

            if sysDrive:
                tempFolder = sysDrive + os.sep + "TEMP"

                if not os.path.exists(tempFolder):
                    os.makedirs(tempFolder,mode=777)

                arcpy.env.scratchWorkspace = tempFolder
                AddMsgAndPrint("\tTemporarily setting scratch workspace to: " + arcpy.env.scratchGDB,1)

            else:
                packageWS = [f for f in arcpy.ListEnvironments() if f=='packageWorkspace']
                if arcpy.env[packageWS[0]]:
                    arcpy.env.scratchWorkspace = arcpy.env[packageWS[0]]
                    AddMsgAndPrint("\tTemporarily setting scratch workspace to: " + arcpy.env.scratchGDB,1)

                else:
                    AddMsgAndPrint("\tCould not set scratchWorkspace. Not even to default!",2)
                    return False

        arcpy.Compact_management(arcpy.env.scratchGDB)
        return arcpy.env.scratchGDB

    except:

        # All Failed; set workspace to packageWorkspace environment
        try:
            packageWS = [f for f in arcpy.ListEnvironments() if f=='packageWorkspace']
            if arcpy.env[packageWS[0]]:
                arcpy.env.scratchWorkspace = arcpy.env[packageWS[0]]
                arcpy.Compact_management(arcpy.env.scratchGDB)
                return arcpy.env.scratchGDB
            else:
                AddMsgAndPrint("\tCould not set scratchWorkspace. Not even to default!",2)
                return False
        except:
            errorMsg()
            return False

## =============================================== Main Body ====================================================

import arcpy, sys, os, traceback

try:

    source_clu = arcpy.GetParameter(0)
    source_dems = arcpy.GetParameterAsText(1).split(";")

    # Set environmental variables
    arcpy.env.geographicTransformations = "WGS_1984_(ITRF00)_To_NAD_1983"
    arcpy.env.resamplingMethod = "BILINEAR"
    arcpy.env.pyramid = "PYRAMIDS -1 BILINEAR DEFAULT 75 NO_SKIP"
    arcpy.env.outputCoordinateSystem = arcpy.Describe(source_clu).SpatialReference
    arcpy.env.overwriteOutput = True

    # Check out Spatial Analyst License
    if arcpy.CheckExtension("Spatial") == "Available":
        arcpy.CheckOutExtension("Spatial")
    else:
        arcpy.AddError("Spatial Analyst Extension not enabled. Please enable Spatial analyst from the Tools/Extensions menu... Exiting!\n")
        sys.exit()

    # Make sure at least 2 datasets to be merged were entered
    datasets = len(source_dems)
    if datasets < 2:
        arcpy.AddError("\nOnly one input DEM layer selected. If you need multiple layers, please run again and select multiple DEM files... Exiting!\n")
        sys.exit()

    # define and set the scratch workspace
    scratchWS = os.path.dirname(sys.argv[0]) + os.sep + r'scratch.gdb'
    if not arcpy.Exists(scratchWS):
       scratchWS = setScratchWorkspace()

    if not scratchWS:
        AddMsgAndPrint("\Could Not set scratchWorkspace!")
        sys.exit()

    arcpy.env.scratchWorkspace = scratchWS
    arcpy.env.workspace = scratchWS

    temp_dem = arcpy.CreateScratchName("temp_dem",data_type="RasterDataset",workspace=scratchWS)
    merged_dem = arcpy.CreateScratchName("merged_dem",data_type="RasterDataset",workspace=scratchWS)
    clu_selected = "in_memory" + os.sep + os.path.basename(arcpy.CreateScratchName("clu_selected",data_type="FeatureClass",workspace=scratchWS))
    clu_buffer = "in_memory" + os.sep + os.path.basename(arcpy.CreateScratchName("clu_buffer",data_type="FeatureClass",workspace=scratchWS))

    # Make sure CLU fields are selected
    cluDesc = arcpy.Describe(source_clu)
    if cluDesc.FIDset == '':
        AddMsgAndPrint("\nPlease select fields from the CLU Layer. Exiting!",2)
        sys.exit()
    else:
        source_clu = arcpy.CopyFeatures_management(source_clu,clu_selected)

    AddMsgAndPrint("\nNumber of CLU fields selected: {}".format(len(cluDesc.FIDset.split(";"))))

    # Dictionary for code values returned by the getRasterProperties tool (keys)
    # Values represent the pixel_type inputs for the mosaic to new raster tool
    pixelTypeDict = {0:'1_BIT',1:'2_BIT',2:'4_BIT',3:'8_BIT_UNSIGNED',4:'8_BIT_SIGNED',5:'16_BIT_UNSIGNED',
                    6:'16_BIT_SIGNED',7:'32_BIT_UNSIGNED',8:'32_BIT_SIGNED',9:'32_BIT_FLOAT',10:'64_BIT'}

    x = 0
    numOfBands = ""
    pixelType = ""

    # --------------------------------------------------------------------------------------- Evaluate every input raster to be merged
    AddMsgAndPrint("\nChecking " + str(datasets) + " input raster layers")
    while x < datasets:
        raster = source_dems[x].replace("'", "")
        desc = arcpy.Describe(raster)
        sr = desc.SpatialReference
        units = sr.LinearUnitName
        bandCount = desc.bandCount

        # Check for Projected Coordinate System
        if sr.Type != "Projected":
            arcpy.addError("\nThe " + str(raster) + " input must have a projected coordinate system... Exiting!\n")
            sys.exit()

        # Check for linear Units
        if units == "Meter":
            tolerance = 3
        elif units == "Foot":
            tolerance = 9.84252
        elif units == "Foot_US":
            tolerance = 9.84252
        else:
            AddMsgAndPrint("\nHorizontal units of " + str(desc.baseName) + " must be in feet or meters... Exiting!\n")
            sys.exit()

        # Check for cell size; Reject if greater than 3m
        if desc.MeanCellWidth > tolerance:
            arcpy.AddError("\nThe cell size of the " + str(raster) + " input exceeds 3 meters or 9.84252 feet which cannot be used in the NRCS HEL Determination Tool... Exiting!\n")
            sys.exit()

        # Check for consistent bit depth
        cellValueCode =  int(arcpy.GetRasterProperties_management(raster,"VALUETYPE").getOutput(0))  # This is the 'VALUETYPE' code returned by the getRasterProperty tool
        bitDepth = pixelTypeDict[cellValueCode]  # Convert the code to pixel depth using dictionary

        if pixelType == "":
            pixelType = bitDepth
        else:
            if pixelType != bitDepth:
                AddMsgAndPrint("\nCannot Mosaic different pixel types: " + bitDepth + " & " + pixelType,2)
                AddMsgAndPrint("Pixel Types must be the same for all input rasters",2)
                AddMsgAndPrint("Contact your state GIS Coordinator to resolve this issue",2)
                sys.exit()

        # Check for consistent band count --- highly unlikely more than 1 band is input
        if numOfBands == "":
           numOfBands = bandCount
        else:
            if numOfBands != bandCount:
                AddMsgAndPrint("\nCannot mosaic rasters with multiple raster bands: " + str(numOfBands) + " & " + str(bandCount),2)
                AddMsgAndPrint("Number of bands must be the same for all input rasters",2)
                AddMsgAndPrint("Contact your state GIS Coordinator to resolve this issue",2)
                sys.exit()
        x += 1

    # --------------------------------------------------------------------------------------------------------- Buffer the selected CLUs by 400m
    arcpy.AddMessage("\nBuffering input CLU fields")
    # Use 410 meter radius so that you have a bit of extra area for the HEL Determination tool to clip against
    # This distance is designed to minimize problems of no data crashes if the HEL Determiation tool's resampled 3-meter DEM doesn't perfectly snap with results from this tool.
    arcpy.Buffer_analysis(source_clu, clu_buffer, "410 Meters", "FULL", "", "ALL", "")

    # --------------------------------------------------------------------------------------------------------- Clip out the DEMs that were entered
    arcpy.AddMessage("\nClipping Raster Layers...")
    x = 0
    del_list = [] # Start an empty list that will be used to clean up the temporary clips after merge is done
    mergeRasters = ""


    while x < datasets:
        current_dem = source_dems[x].replace("'", "")
        out_clip = temp_dem + "_" + str(x)

        arcpy.SetProgressorLabel("Clipping " + current_dem + " " + str(x+1) + " of " + str(datasets))

        try:
            AddMsgAndPrint("\tClipping " + current_dem + " " + str(x+1) + " of " + str(datasets))
            extractedDEM = arcpy.sa.ExtractByMask(current_dem, clu_buffer)
            extractedDEM.save(out_clip)
        except:
            arcpy.AddError("\nThe input CLU fields may not cover the input DEM files. Clip & Merge failed...Exiting!\n")
            sys.exit()

        # Create merge statement
        if x == 0:
            # Start list of layers to merge
            mergeRasters = "" + str(out_clip) + ""
        else:
            # Append to list
            mergeRasters = mergeRasters + ";" + str(out_clip)

        # Append name of temporary output to the list of temp soil layers to be deleted
        del_list.append(str(out_clip))
        x += 1

    # --------------------------------------------------------------------------------------------------------- Merge Clipped Datasets
    arcpy.AddMessage("\nMerging inputs...")

    if arcpy.Exists(merged_dem):
        arcpy.Delete_management(merged_dem)

    cellsize = 3
    arcpy.MosaicToNewRaster_management(mergeRasters, scratchWS, os.path.basename(merged_dem), "#", pixelType, cellsize, numOfBands, "MEAN", "#")

    # Clean-up
    #arcpy.AddMessage("\nCleaning up...")
    for lyr in del_list:
        arcpy.Delete_management(lyr)
    arcpy.Delete_management(clu_buffer)

    # Add resulting data to map
    arcpy.AddMessage("\nAdding " + os.path.basename(merged_dem) + " to ArcMap session\n")
    arcpy.SetParameterAsText(2, merged_dem)

except:
    errorMsg()
