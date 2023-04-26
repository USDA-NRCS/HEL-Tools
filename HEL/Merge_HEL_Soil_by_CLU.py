# ==========================================================================================
# Name:   Merge HEL Soil by CLU
#
# Author(s): Chris.Morse
#            State GIS Specialist
# e-mail:    chris.morse@usda.gov
# phone:     317.295.5849

#            Adolfo.Diaz
#            Region 10 GIS Specialist
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

try:
    # Define and set the scratch workspace
    scratchWS = setScratchWorkspace()
    arcpy.env.scratchWorkspace = scratchWS

    if not scratchWS:
        arcpy.AddError("\n Could not set scratch workspace... Exiting!\n")
        sys.exit()
        
    # Inputs
    arcpy.AddMessage("\nExaming inputs...")
    source_clu = arcpy.GetParameterAsText(0)
    source_soils = arcpy.GetParameterAsText(1).split(";")
    datasets = len(source_soils)
    
    # Make sure at least 2 datasets to be merged were entered
    if datasets < 2:
        arcpy.AddError("\n Only one input soil layer selected. If you need multiple layers, please run again and select multiple soil layers... Exiting!\n")
        sys.exit()

    # List of valid HEL soil layer schema for the tool (in lower case for comparative purposes)
    schema = ["areasymbol", "spatialver", "musym", "muname", "muhelcl", "t", "k", "r"]
    # Check the input soil layers to make sure they contain fields with the same field names
    x = 0
    for layer in source_soils:
        field_names = [f.name.lower() for f in arcpy.ListFields(source_soils[x])]
        for s in schema:
            if s not in field_names:
                arcpy.AddMessage("\n The layer " + str(source_soils[x]) + " is missing field " + str(s) + "... Exiting!\n")
                sys.exit()
        x += 1
    del x

    arcpy.AddMessage("Done!\n")

    # Variables
    temp_soil = os.path.join(os.path.dirname(sys.argv[0]), "HEL.mdb" + os.sep + "temp_soil")
    merged_soil = os.path.join(os.path.dirname(sys.argv[0]), "HEL.mdb" + os.sep + "Merged_HEL_Soil")

    # Clip out the soils that were entered
    arcpy.AddMessage("\nClipping inputs...")
    x = 0
    del_list = []
    # Start an empty list that will be used to clean up the temporary clips after merge is done
    while x < datasets:
        current_soil = source_soils[x].replace("'", "")
        out_clip = temp_soil + "_" + str(x)
        try:
            arcpy.Clip_analysis(current_soil, source_clu, out_clip)
        except:
            arcpy.AddError("\n The input fields may not cover the input soil layers. Clip & Merge failed...Exiting!\n")
            sys.exit()
        if x == 0:
            # Start list of layers to merge
            merge_list = "" + str(out_clip) + ""
        else:
            # Append to list
            merge_list = merge_list + ";" + str(out_clip)
        # Append name of temporary output to the list of temp soil layers to be deleted
        del_list.append(str(out_clip))
        x += 1
    del x
    arcpy.AddMessage("Done!\n")

    # Merge Clipped Datasets
    arcpy.AddMessage("\nMerging inputs...")
    if arcpy.Exists(merged_soil):
        arcpy.Delete_management(merged_soil)
    arcpy.Merge_management(merge_list, merged_soil, "")
    arcpy.AddMessage("Done!\n")

    # Clean-up
    # Delete temporary soils
    arcpy.AddMessage("\nCleaning up...")
    for lyr in del_list:
        arcpy.Delete_management(lyr)
    arcpy.AddMessage("Done!\n")

    # Add resulting data to map
    arcpy.AddMessage("\nAdding layer to map...")
    arcpy.SetParameterAsText(2, merged_soil)
    arcpy.AddMessage("Done!\n")
    
except:
    pass
