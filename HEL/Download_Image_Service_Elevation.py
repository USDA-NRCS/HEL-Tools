# ==========================================================================================
# Name:   Merge Local DEMs by CLU
#
# Author(s): Chris.Morse
#            State GIS Specialist
# e-mail:    chris.morse@usda.gov
# phone:     317.295.5849

#            Adolfo.Diaz
#            GIS Specialist
#            National Soil Survey Center
#            USDA - NRCS
# e-mail:    adolfo.diaz@usda.gov
# phone:     608.662.4422 ext. 216

# Contributor: Kevin Godsey
#              Soil Scientist
# e-mail:      kevin.godsey@usda.gov
# phone:       608.662.4422 ext. 190

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

# Set overwrite
arcpy.env.overwriteOutput = True

# Set other environments
arcpy.env.geographicTransformations = "WGS_1984_(ITRF00)_To_NAD_1983"
arcpy.env.resamplingMethod = "BILINEAR"
arcpy.env.pyramid = "PYRAMIDS -1 BILINEAR DEFAULT 75 NO_SKIP"

try:
    # Check out Spatial Analyst License
    if arcpy.CheckExtension("Spatial") == "Available":
        arcpy.CheckOutExtension("Spatial")
    else:
        arcpy.AddError("Spatial Analyst Extension not enabled. Please enable Spatial analyst from the Tools/Extensions menu... Exiting!\n")
        sys.exit()

    # Define and set the scratch workspace
    #scratchWS = setScratchWorkspace()
    #arcpy.env.scratchWorkspace = scratchWS

    # Inputs
    arcpy.AddMessage("\nExaming inputs...")
    source_clu = arcpy.GetParameterAsText(0)
    source_Service = arcpy.GetParameterAsText(1)

    # Variables
    WGS84_DEM =    os.path.join(os.path.dirname(sys.argv[0]), "scratch.gdb" + os.sep + "WGS84_DEM")
    final_DEM =    os.path.join(os.path.dirname(sys.argv[0]), "scratch.gdb" + os.sep + "Downloaded_DEM")
    clu_buffer =   os.path.join(os.path.dirname(sys.argv[0]), "scratch.gdb" + os.sep + "CLU_Buffer")
    clu_buff_wgs = os.path.join(os.path.dirname(sys.argv[0]), "scratch.gdb" + os.sep + "CLU_Buffer_WGS")

    # Buffer the selected CLUs by 400m
    arcpy.AddMessage("\nBuffering input CLU fields...")
    # Use 410 meter radius so that you have a bit of extra area for the HEL Determination tool to clip against
    # This distance is designed to minimize problems of no data crashes if the HEL Determiation tool's resampled 3-meter DEM doesn't perfectly snap with results from this tool.
    arcpy.Buffer_analysis(source_clu, clu_buffer, "410 Meters", "FULL", "", "ALL", "")
    arcpy.AddMessage("Done!\n")

    # Delete the all posible temp datasets if they already exist
    if arcpy.Exists(WGS84_DEM):
        try:
            arcpy.Delete_management(WGS84_DEM)
        except:
            pass
    if arcpy.Exists(final_DEM):
        try:
            arcpy.Delete_management(final_DEM)
        except:
            pass
    if arcpy.Exists(clu_buff_wgs):
        try:
            arcpy.Delete_management(clu_buff_wgs)
        except:
            pass

    # Re-project the AOI to WGS84 Geographic (EPSG WKID: 4326)
    arcpy.AddMessage("\nConverting CLU Buffer to WGS 1984...")
    wgs_CS = arcpy.SpatialReference(4326)
    arcpy.Project_management(clu_buffer, clu_buff_wgs, wgs_CS)
    arcpy.AddMessage("Done!\n")

    # Use the WGS 1984 AOI to clip/extract the DEM from the service
    arcpy.AddMessage("\nDownloading Data...")
    aoi_ext = arcpy.Describe(clu_buff_wgs).extent
    xMin = aoi_ext.XMin
    yMin = aoi_ext.YMin
    xMax = aoi_ext.XMax
    yMax = aoi_ext.YMax
    clip_ext = str(xMin) + " " + str(yMin) + " " + str(xMax) + " " + str(yMax)
    arcpy.Clip_management(source_Service, clip_ext, WGS84_DEM, "", "", "", "NO_MAINTAIN_EXTENT")
    arcpy.AddMessage("Done!\n")

    # Project the WGS 1984 DEM to the coordinate system of the input CLU layer
    arcpy.AddMessage("\nProjecting data to match input CLU...\n")
    final_CS = arcpy.Describe(source_clu).spatialReference.factoryCode
    cellsize = 3
    arcpy.ProjectRaster_management(WGS84_DEM, final_DEM, final_CS, "BILINEAR", cellsize)
    arcpy.AddMessage("Done!\n")

    # Clean-up
    # Delete temporary data
    arcpy.AddMessage("\nCleaning up...")
    arcpy.Delete_management(WGS84_DEM)
    arcpy.Delete_management(clu_buff_wgs)
    arcpy.Delete_management(clu_buffer)
    arcpy.AddMessage("Done!\n")

    # Add resulting data to map
    arcpy.AddMessage("\nAdding layer to map...")
    arcpy.SetParameterAsText(2, final_DEM)
    arcpy.AddMessage("Done!\n")

except SystemExit:
    pass

