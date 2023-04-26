# ==========================================================================================
# Name:   HEL Determination by AOI
#
# Author: Adolfo.Diaz
#         GIS Specialist
#         National Soil Survey Center
#         USDA - NRCS
# e-mail: adolfo.diaz@usda.gov
# phone: 608.662.4422 ext. 216

# Contributor: Kevin Godsey
#              Soil Scientist
# e-mail: kevin.godsey@usda.gov
# phone: 608.662.4422 ext. 190

# Contributor: Christiane Roy
#              MN Area GIS Specialist
# e-mail: christian.roy@usda.gov
# phone: 507.405.3580

# Contributor: Chris Morse
#              IN State GIS Specialist
# e-mail: chris.morse@usda.gov
# phone: 317-295-5849

# Contributor: Edwin Muniz
#               NJ Assistant State Soil Scientist
# e-mail: edwin.muniz@usda.gov
# phone: 732-537-6062

# ==========================================================================================
# Modified 10/24/2016
# Line 705; outflowLengthFT was added to the scratchLayers to be deleted.  It should've
# been flowLengthFT
#
# The acreConversion variable was updated to reference a dictionary instead.  The original
# acreConversion variable was being determined from the DEM.  If input DEM was in FT then
# the wrong acre conversion was applied to layers that were in Meters.

# ==========================================================================================
# Modified 11/19/2018
# Tim Prescott was having acre disrepancies.  He had a wide variety of coordinate systems
# in his arcmap project.  I couldn't isolate the problem but I wound up changine
# the way the z-factor is assigned.  I was assuming a meter and/or feet combination but
# there could be other combinations like centimeters. Created a matrix of XY and Z
# unit combinations to be used as a look up table.

# ==========================================================================================
# Modified 11/15/2018
# Christiane reported duplicate labeling in the HEL Summary feature class.  The duplicate
# labels go away when the corresponding .lyr file is added instead of the feature class.
# Modified the code to add the .lyr to Arcmap only for the HEL Summary layer.

# ==========================================================================================
# Modified 3/1/2019
# - Error Encountered:
#      Related to optional parameter of zUnits
#	   File "C:\python_scripts\GitHub\HEL\HEL_Determination_by_AOI.py", line 517, in <module>
#      zFactor = zFactorList[unitLookUpDict.get(zUnits)][unitLookUpDict.get(linearUnits)]
#	   TypeError: list indices must be integers, not NoneType
#
# - Add Focal Statistics (MAXIMUM) to Flow Length output to smooth it out
# - switched method to exit python interpreter from exit() to sys.exit() and SystemExit Exception
#   to catch errors.
# - Switched ALL output layers to in-memory layers to optimize performance; Performance increased
#   by at least 25%
# - Updated the Determination Date to include the date the script was executed: mm/dd/yyyy
# - In the 'Compute Summary of NEW HEL values' section the total CLU Acres were re-computed from the
#   outTabulate table vs. looking them up from the 'fieldDetermination' fc b/c they wouldn't add up to 100%
#   The outTabulate represents sums from raster cells vs. polygons.
# - Wrapped main body into 'main' function and executed from if __name__ == '__main__':
# - sys.exit() was throwing an exception and forcing a traceback to be printed.  Ignore traceback
#   message if thrown by sys.exit()

# ==========================================================================================
# Modified 3/14/2019
# - The tool was updated so that it no longer exits when there is no PHEL present or final results
#   are all HEL or NHEL.  The 'outTabulate' table had to be updated to add either 'VALUE_1' or
#   'VALUE_2' field and calc'ed to 0 so that the script could finish.
#
# - Created 'ogCLUinfoDict' to capture cluNumber and cluAcres during the original summary process.
#   This dictionary is only used if final results are ALL HEL or ALL NHEL since acres don't need
#   to be recalculated, especially from the tabulate areas results where the acreages may slightly
#   differ b/c of pixel tabulation.  This way the acreages will always match the original figures.
#
# - 2 booleans (bOnlyHEL & bOnlyNHEL) were added to determine if final results are ALL HEL or NHEL

# ==========================================================================================
# Modified 2/4/2019
# - Seperated the scratch layers by individual task vs adding several layers at a time.
#   This will ensure all scratch layers are deleted when a failure occurs.
# - Run Focal Statistics (MEAN) to Slope output to smooth it out
# - Changed Focal Statistics on Flow Length from MAXIMUM to MEAN

# ==========================================================================================
# Modified 4/12/2019
# - Report summary of CLU HEL values regardless of no PHEL Values found.  If no PHEL values
#   are found then geoprocessing will be entirely skipped however the 1026 form will still
#   be populated and brought up.  The populateForm function was added to be able to skip
#   geoprocessing and still open the 1026 form.  If no PHEL values are found then HEL will
#   be determined by the 33.33% or 50 acre rule.  If a CLU's acreage is greater or equal to 50
#   or a CLU's acre percentage is greater or equal to 33.3% then it is "HEL" otherwise the
#   rating is determined by its dominant acre HEL value.
# - Set the 33.3% or 50 acre HEL rule when computing new HEL values after geoprocessing.
# - Autopopulated 11th parameter of tool, DC Signature, to computer user name.  Assumption
#   is made that the user is the DC. User could be the technician.
# - Autopopulate 10th Parameter, State, by isolating state namne from user computer name.

# ==========================================================================================
# Modified 4/30/2019
# - Arcmap was throwing drawing error when ading the 'HEL YES NO' layer only when no
#   PHEL values were present.  This happened b/c geoprocessing was entirely skipped and
#   some necessary fields were missing.  Added the following fields: HEL_YES,HEL_Acres,HEL_Pct
#   to the 'HEL YES NO' layer.  These fields are being populated from the ogCLUinfoDict dict.
# - Due to rounding in acres and percentages, some percentages were coming out at slightly
#   above 100%.  If percentages are over 100%, then set to 100%.
# - Added code to kill the MSACCESS.EXE process if it is open.  Having access open will
#   cause fieldDetermination, helSummary, finalHELmap from being deleted properly.
#   User will be prompt if Access was open and closed.
# - Absolute path for HELLayer was used instead of the parameter object b/c there could
#   be selected polygons within the HELLayer which could cause an incorrect intersection
#   with the CLUs.
# - Removed validation code to autopopulate the DC Signature.  It was decided that more than
#   likely the user would not be the DC and therefore autopopulated incorrect.

# ==========================================================================================
# Updated 5/15/2019
# A meeting was held in St.Paul, MN to discuss updates and additions to the functionality
# of this tool.  The following updates were integrated between 5/15 and 6/4:
#
#  1) Final field HEL Determination is determined by 2 factors:
#     a) Determine HEL value for each delineation.  A delineation is HEL if 50% of its
#        acres are HEL
#     b) Determine HEL value for each field.  A field is HEL if it's HEL acres are greater
#        or equal to 33.33 percent OR the HEL acres are greater than 50.
#  2) Rename output layers to the access database:
#      HELYESNO        --> Field Determination
#      cluHELIntersect --> Final HEL Summary
#      HELsummary      --> Initial HEL Summary
#      FinalHELMap     --> LiDAR HEL Summary
#  3) HEL output values are switched from YES/NO to HEL/NHEL
#  4) Polygons are exploded after the intersection between CLUs and Soils b/c each
#     delination must be assessed individually
#  5) Buffer CLU by 400 meters instead of 300.  This is approximately a 1/4 of a mile.
#  6) Updated sq.meter to acre conversion factor to 4046.8564224
#  7) Removed the 'This preliminary determination was conducted off-site with LiDAR data
#     only if PHEL mapunits are present' remark
#  8) Redid the AddLayersToArcMap to utilize .lyr files for full symbology.  This allows
#     field offices with the flexibility of updating symbology without altering the code.
#  9) Renamed Toolbox and tool to 'NRCS HEL Determination'
# 10) Removed Parameter #1 (manual Digitizing of AOI).  The tool will not exit if a
#     selected set of CLU polygons is not used.  Instead, the tool will exit if there
#     are multiple tracts found. This also leaves flexibility for definition queries
#     to be used without a selected set.
# 11) Removed HEL Summary by AOI.  This was a summary of all of the fields together.
#     It was determined that this summary was not valuable b/c fields are what are
#     important.
# 12) Added functionality to utilize a DEM image service.  Added 2 new functions
#     to handle this capability: extractDEM and extractDEMfromImageService.  Image
#     service function could use a little more work to determine cell resolution of
#     a service in WGS84.
# 13) Bypass geoprocessing if field is HEL and (pct >= 33.3 or acres >= 50) OR
#     NHEL pct > 66.67.  The boolean bSkipGeoprocessing will is set to True by default
#     until 1 field does not meet the HEL or NHEL criteria above, in which case the
#     field will have to be analyzed using LiDAR.  If 1 field requires geoprocessing
#     ALL fields will be processed.
# 14) After adding 2-4 layers to arcmap, depending on the outcome, the original selected
#     polygons from the CLU are cleared and the dataframe extent is set to the extent
#     of the 'Field Determination' layer.
# 15) Final HEL results are summarized in order similar to the initial HEL Summary.
#     Originally, the final results were printed in random CLU Field order b/c the
#     results were being read from a dictionary.  A dict is inherently not ordered.
# 16) All messages printed to the Arcmap console will be logged to a text file that
#     will be saved in the 'HEL_Text_Files' folder.  If the folder does not exist it
#     will automatically be created in the same folder that the scripts are located.
#     The text file will have the prefix "NRCS_HEL_Determination" and the Tract, Farm
#     and field numbers appended to the end.
# 17) Added code to verify that a subset of CLU fields were selected otherwise the
#     tool will exit.  This will prevent an entire CLU layer from being executed.
# 18) Remove all progress positions. very minor!
# 19) Exit if there are any NULLs or invalid HEL labels.  Previously this was only
#     set as a WARNING.  Now it is an ERROR.
# 20) Removed the state parameter from the tool and instead mined the state from
#     the field determination layer.  A dictionary was created to be used as a
#     conversion from state fips code to state abbreviation.  If obtaining the state
#     abbreviation fails using the dictionary the state name will be extracted from
#     the user computer name.
# 21) All .lyr files were set to relative path and description populated.
# 22) Validation code was updated to autopopulate layers and fields based on the
#     newly adopted naming convention of layers and field schema.

# ==========================================================================================
# Updated 6/26/2019
# - When an elevation service was used the slope derivative was being derived incorrectly.
#   It turns out that the environmental output coordinate system variable set to
#   4326 (WGS84) in the extractDEMfromImageService function was overriding the output
#   coordinate system in the project raster tool.  I didn't think this variable could
#   override a specfied tool parameter.  I reintroduced this variable before the project
#   raster tool and set it to the same as the output coordinate system used in the project
#   raster tool.
# - Code was added to set the dataframe extent to the field determination layer but buffered
#   to 50 meters.  The buffer was introduced so that the dataframe extent wasn't too tight
#   around the selected fields.  The dataframe extent was set to the selected fields so
#   that the cartography looks better.
# - Altered the populateForm function so that all tables are populated regardless of whether
#   Microsoft Access is found.  Microsoft Access is installed differently in the VDI
#   environment so the tables never populated.  This way the user can manually open Access
#   and continue with the workflow.

# ==========================================================================================
# Updated 7/1/2019
# - To minimize confusion, 'Final HEL Values' was switched to LiDAR HEL values; The confusion
#   when there are NHEL pixels that are summarized from the geoprocessing, the user expects
#   to see NHEL polygons.  However, the polygons are assessed differently.
# - "RuntimeError: Too few parameters. Expected 1" error message was encountered when
#   determining if individual PHEL delineations are HEL or NHEL.  Specifically, the error
#   was during an updateCursor to the finalHELSummary layer when there was no HEL values
#   present.  "Value_2" was added to the outPolyTabulate table but it wasn't being recognized.
#   It turns out it is an access datbase issue related to adding field names:
#   https://support.esri.com/en/technical-article/000006080
#   To circumvent this issue the finalHELSummary layer and the outPolyTabulate table
#   are first joined and then if HEL values (VALUE_2) is not present then its added to
#   the finalHELSumary layer.
# - There was an error when an image service was used AND z-units left blank.  Normally
#   the z-units, if left blank, are set to the same as linear units but most image
#   services' linear units are set to degrees, which isn't a valid z-unit.  A check was put
#   in place.  the Z-units are set to the projected dem linear units.

# Updated 7/17/2019
# - Received an error from Julie Pitts: ERROR 000728: Field OBJECTID_1_1 does not exist within
#   table.  When joining 2 tables using OBJECTID as a join field, the output table will
#   create its own OBJECTID forcing the join field to be renamed to OBJECTID_1.  However, if
#   one of the incoming tables has an existing OBJECTID_1 then it creates a problem.  A fix
#   was introduced to look for an "_" in the OID fieldname and append "_1" to that field if it
#   is present.  This should take care of any errors with join fields.
# - It turns out that the field level determination calculation was wrong.  The tool was summarizing
#   HEL values, inlcuding pixels from geoprocessing to determing if field was HEL.  The tool
#   has been updated to reflect sums of acres from polygons that are HEL and then making the field
#   evaluation.
# - Added "NA" as a valid HEL attribute.  It will be processed as NHEL.
# - Updated validation code to find HEL Layer from the Arcmap TOC based on the new schema of
#   'HEL_Frozen_a'
# - Added a check to make sure a DEM is present if PHEL attributes are present.

# Updated 8/12/2019, C.M.
# - Re-runs would crash on ArcMap 10.6.1 and up. Changed scratch workspace for K, T, R, and
#   HELValue layers to use the scratch.gdb instead of memory to avoid this crash.
# - Updated L-factor formula to use 72.6 as the constant (per AH 537 pg. 12)
# - Updated S-factor formula to use radians from slope % (per AH 537 pg. 12)
# - Updated zfactor conversion lookups to correctly handle a variety of input DEMs
# - Updated conditions for conversion of flow length layer to feet to use new linear units
#   from the dem returned from the extractDEM function. Previous condition assessing zUnits
#   led to incorrect results when using DEMs with vertical feet.
# - Set Snap Raster environment equal to the final dem returned from extract DEM. This
#   setting properly aligns the rasters generated from the soils for K, T, and R to fit
#   to the DEM pixel space. We can't use the soils layer to generate snap rasters from the
#   K, T, or R layers that are generated because the origin for those layers will vary clip
#   to clip and site to site, compared to a stock DEM that already exists in pixel space. Also,
#   shifting the DEM would reduce confidence in slope percentage and length computations.

# Updated 8/15/2019, C.M.
# - A Fill function is run on the DEM prior to creating derivative layers. The fill surface is
#   limited to filling depths no greater than 1 foot. A conversion factor for setting the fill
#   limit was created to express the limit in amounts that correspond to the input DEM's Z units.
#   If no Z Units are specified, the smallest fill amount is assumed (0.3048, as if units are
#   meters).
# - Focal Statistics Mean was moved from running on slope to running on the DEM, per typically
#   accepted lidar-based DEM modeling workflows for creating slope percentage.

# Updated 8/20/2019, C.M.
# - Created a new map layout for use in the release version where map elements are tagged
#   with element names as hooks to program against.
# - Added configLayout function to autopopulate elements of the map layout which include:
#    - Farm Number
#    - Tract Number
#    - Customer Name (from parameter inputs)
#    - County Name (Actual tract location; not administrative county)
#    - State (Postal Code) (Actual tract location; not administrative state)
# - Updated populateForm function:
#    - Request From field's default entry changed to "FSA", per Jason Outlaw.
#    - Default remarks statement changed, per Jason Outlaw, and lengthened to 255 characters
#    - Customer Name input parameter is output to the First Name parameter for the time being.
#    - Increased First Name field length to 50.
# - Disabled Snap Raster for release version 1. Need to fix in future versions. See notes by
#   env.snapRaster.

# Updated 9/16/2019, C.M.
# -Added alternate equation for regions of the northwestern US that need a special LS equation
#  that is used for HEL determinations involving runoff on partially frozen soil.
# -Updated HEL Determination Tool validations to better handle automatic DEM and units assignment.
# -Fixed a crash in the extractDEMfromImageService function related to variables and returns.
#  The direct use of NRCS Web Service DEMs is working again without need for the download utility.

# Updated 8/10/2020, C.M.
# -Updated county lookups to pass county names to the table for use in Access and on a new 026 form.
# -Added workaround to handle DEMs that generate slightly larger than 3m but slightly increasing
#  tolerance values.
# -Updated populateForm to set the HEL on Farm value to not always populate, in case results are
#  all NHEL.
# -Incorporated a code snippet from Kentucky to handle input CLU layers if the layer is a member
#  of a group layer in ArcMap.

#-------------------------------------------------------------------------------

## ===================================================================================
def AddMsgAndPrint(msg, severity=0):
    # prints message to screen if run as a python script
    # Adds tool message to the geoprocessor
    #
    #Split the message on \n first, so that if it's multiple lines, a GPMessage will be added for each line
    try:

        if bLog:
            f = open(textFilePath,'a+')
            f.write(msg + " \n")
            f.close
            del f

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
            AddMsgAndPrint("\n\tNRCS HEL Tool Error: -------------------------",2)
            AddMsgAndPrint(theMsg,2)

    except:
        AddMsgAndPrint("Unhandled error in errorMsg method", 2)
        pass

## ==================================================================================
def tic():
    """ Returns the current time """

    return time.time()

## ==================================================================================
def toc(_start_time):
    """ Returns the total time by subtracting the start time - finish time"""

    try:

        t_sec = round(time.time() - _start_time)
        (t_min, t_sec) = divmod(t_sec,60)
        (t_hour,t_min) = divmod(t_min,60)

        if t_hour:
            return ('{} hour(s): {} minute(s): {} second(s)'.format(int(t_hour),int(t_min),int(t_sec)))
        elif t_min:
            return ('{} minute(s): {} second(s)'.format(int(t_min),int(t_sec)))
        else:
            return ('{} second(s)'.format(int(t_sec)))

    except:
        errorMsg()

## ===================================================================================
def createTextFile(Tract,Farm,cluList):
    """ This function sets up the text file to begin recording all messages
        reported to the console.  The text file will be created in a folder
        called 'HEL_Text_Files' in argv[0].  The text file will have the prefix
        "NRCS_HEL_Determination" and the Tract, Farm and field numbers appended
        to the end.  Basic information will be collected and logged to the
        text file.  The function will return the full path to the text file."""

    try:
        # Set log file
        helTextNotesDir = os.path.dirname(sys.argv[0]) + os.sep + 'HEL_Text_Files'
        if not os.path.isdir(helTextNotesDir):
           os.makedirs(helTextNotesDir)

##        # name was too long when using numerous fields in the name
##        textFileName = "NRCS_HEL_Determination_TRACT(" + str(Tract) + ")_FARM(" + str(Farm) \
##                        + ")_CLU(" + str(cluList).replace('[','').replace(']','').replace(' ','') \
##                        + ").txt"

        textFileName = "NRCS_HEL_Determination_TRACT(" + str(Tract) + ")_FARM(" + str(Farm) \
                        + ").txt"

        textPath = helTextNotesDir + os.sep + textFileName

        # Version check
        version = str(arcpy.GetInstallInfo()['Version'])
        if version.find("10.") > -1:
            ArcGIS10 = True
        else:
            ArcGIS10 = False

        # Convert version string to a float value
        versionFlt = float(version[0:4])
        if versionFlt < 10.3:
            AddMsgAndPrint.AddError("\nThis tool has only been tested on ArcGIS version 10.4 or greater",1)

        f = open(textPath,'a+')
        f.write('#' * 80 + "\n")
        f.write("NRCS HEL Determination Tool\n")
        f.write("User Name: " + getpass.getuser() + "\n")
        f.write("Date Executed: " + time.ctime() + "\n")
        f.write("ArcGIS Version: " + str(version) + "\n")
        f.close

        return textPath

    except:
        errorMsg()

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
                above then check their value against the user-set scratch workspace.  If they have anything
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

        #arcpy.Compact_management(arcpy.env.scratchGDB)
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

## ===============================================================================================================
def splitThousands(someNumber):
# will determine where to put a thousands seperator if one is needed.
# Input is an integer.  Integer with or without thousands seperator is returned.

    try:
        return re.sub(r'(\d{3})(?=\d)', r'\1,', str(someNumber)[::-1])[::-1]

    except:
        errorMsg()
        return someNumber

## ================================================================================================================
def FindField(layer,chkField):
    # Check table or featureclass to see if specified field exists
    # If fully qualified name is found, return that name; otherwise return ""
    # Set workspace before calling FindField

    try:

        if arcpy.Exists(layer):

            theDesc = arcpy.Describe(layer)
            theFields = theDesc.fields
            theField = theFields[0]

            for theField in theFields:

                # Parses a fully qualified field name into its components (database, owner name, table name, and field name)
                parseList = arcpy.ParseFieldName(theField.name) # (null), (null), (null), MUKEY

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

## ================================================================================================================
def getDatum(layer):
    # This function will isolate the datum of the input raster and return the spatial
    # reference datum if it is WGS84 or NAD83 otherwise return "Other".  The input
    # must be the catalog path.  The datum name or code cannot be mined using the
    # describe object on projected coordinate systems, only Geographic Coordinate
    # systems which is why the Create Spatial Reference tool is used.

    try:
        # Create Spatial Reference of the input fc. It must first be converted in to string in ArcGIS10
        # otherwise .find will not work.
        spatialRef = str(arcpy.CreateSpatialReference_management("#",layer,"#","#","#","#"))
        datum_start = spatialRef.find("DATUM") + 7
        datum_stop = spatialRef.find(",", datum_start) - 1
        datum = spatialRef[datum_start:datum_stop]

        # Create the GCS WGS84 spatial reference and datum name using the factory code
        WGS84_sr = arcpy.SpatialReference(4326)
        WGS84_datum = WGS84_sr.datumName

        NAD83_datum = "D_North_American_1983"

        # Input datum is either WGS84 or NAD83; return true
        if datum == WGS84_datum:
            return WGS84_datum

        elif datum == NAD83_datum:
            return NAD83_datum

        # Input Datum is some other Datum; return false
        else:
            return "Other"

    except arcpy.ExecuteError:
        AddMsgAndPrint(arcpy.GetMessages(2),2)
        return False

    except:
        errorMsg()
        return False

## ================================================================================================================
def extractDEMfromImageService(demSource,zUnits):
    # This function will extract a DEM from a Web Image Service that is in WGS.  The
    # CLU will be buffered to 410 meters and set to WGS84 GCS in order to clip the DEM.
    # The clipped DEM will then be projected to the same coordinate system as the CLU.
    # -- Eventually code will be added to determine the approximate cell size  of the
    #    image service using y-distances from the center of the cells.  Cell size from
    #    a WGS84 service is difficult to calculate.
    # -- Clip is the fastest however it doesn't honor cellsize so a project is required.
    # -- Original Z-factor on WGS84 service cannot be calculated b/c linear units are
    #    unknown.  Assume linear units and z-units are the same.
    # Returns a clipped DEM and new Z-Factor

    try:
        #startTime = tic()
        desc = arcpy.Describe(demSource)
        sr = desc.SpatialReference
        cellSize = desc.MeanCellWidth
        outputCellsize = 3

        AddMsgAndPrint("\nInput DEM Image Service: " + desc.baseName)
        AddMsgAndPrint("\tGeographic Coordinate System: " + sr.Name)
        AddMsgAndPrint("\tUnits (XY): " + sr.AngularUnitName)

        # Set output env variables to WGS84 to project clu
        arcpy.env.geographicTransformations = "WGS_1984_(ITRF00)_To_NAD_1983"
        arcpy.env.outputCoordinateSystem = arcpy.SpatialReference(4326)

        # Buffer CLU by 410 Meters. Output buffer will be in GCS
        cluBuffer = "in_memory" + os.sep + os.path.basename(arcpy.CreateScratchName("cluBuffer_GCS",data_type="FeatureClass",workspace=scratchWS))
        arcpy.Buffer_analysis(fieldDetermination, cluBuffer, "410 Meters", "FULL", "", "ALL", "")

        # Use the WGS 1984 AOI to clip/extract the DEM from the service
        cluExtent = arcpy.Describe(cluBuffer).extent
        clipExtent = str(cluExtent.XMin) + " " + str(cluExtent.YMin) + " " + str(cluExtent.XMax) + " " + str(cluExtent.YMax)

        arcpy.SetProgressorLabel("Downloading DEM from " + desc.baseName + " Image Service")
        AddMsgAndPrint("\n\tDownloading DEM from " + desc.baseName + " Image Service")

        demClip = "in_memory" + os.sep + os.path.basename(arcpy.CreateScratchName("demClipIS",data_type="RasterDataset",workspace=scratchWS))
        arcpy.Clip_management(demSource, clipExtent, demClip, "", "", "", "NO_MAINTAIN_EXTENT")

        # Project DEM subset from WGS84 to CLU coord system
        outputCS = arcpy.Describe(cluLayer).SpatialReference
        arcpy.env.outputCoordinateSystem = outputCS
        demProject = "in_memory" + os.sep + os.path.basename(arcpy.CreateScratchName("demProjectIS",data_type="RasterDataset",workspace=scratchWS))
        arcpy.ProjectRaster_management(demClip, demProject, outputCS, "BILINEAR", outputCellsize)

##        testing_raster = r'C:\Temp\Temp2.gdb\dem_testing'
##        arcpy.CopyRaster_management(demClip, testing_raster)
        arcpy.Delete_management(demClip)

        # ------------------------------------------------------------------------------------ Report new DEM properties
        desc = arcpy.Describe(demProject)
        newSR = desc.SpatialReference
        newLinearUnits = newSR.LinearUnitName
        newCellSize = desc.MeanCellWidth

        # if zUnits not populated assume it is the same as linearUnits
        if not zUnits: zUnits = newLinearUnits

        #newZfactor = zFactorList[unitLookUpDict.get(zUnits)][unitLookUpDict.get(newLinearUnits)]
        newZfactor = zFactorList[unitLookUpDict.get(newLinearUnits)][unitLookUpDict.get(zUnits)]

        AddMsgAndPrint("\t\tNew Projection Name: " + newSR.Name,0)
        AddMsgAndPrint("\t\tLinear Units (XY): " + newLinearUnits)
        AddMsgAndPrint("\t\tElevation Units (Z): " + zUnits)
        AddMsgAndPrint("\t\tCell Size: " + str(newCellSize) + " " + newLinearUnits )
        AddMsgAndPrint("\t\tZ-Factor: " + str(newZfactor))

        #AddMsgAndPrint(toc(startTime))
        return newLinearUnits,newZfactor,demProject

    except:
        errorMsg()

## ================================================================================================================
def extractDEM(inputDEM,zUnits):
    # This function will return a DEM that has the same extent as the CLU selected fields
    # buffered to 410 Meters.  The DEM can be a local raster layer or a web image server. Datum
    # must be in WGS84 or NAD83 and linear units must be in Meters or Feet otherwise it
    # will exit.  If the cell size is finer than 3M then the DEM will be resampled.
    # The resampling will happen using the Project Raster tool regardless of an actual
    # coordinate system change.  If the cell size is 3M then the DEM will be clipped using
    # the buffered CLU.  Environmental settings are used to control the output coordinate system.
    # -- Function does not check the SR of the CLU.  Assumes the CLU is in a project coord system
    #    and in meters.  Should probably verify before reprojecting to a 3M cell.
    # returns a clipped DEM and new Z-Factor

    try:

        # Set environmental variables
        arcpy.env.geographicTransformations = "WGS_1984_(ITRF00)_To_NAD_1983"
        arcpy.env.resamplingMethod = "BILINEAR"
        arcpy.env.outputCoordinateSystem = arcpy.Describe(cluLayer).SpatialReference

        bImageService = False
        bResample = False
        outputCellSize = 3

        desc = arcpy.Describe(inputDEM)
        inputDEMPath = desc.catalogPath
        sr = desc.SpatialReference
        linearUnits = sr.LinearUnitName
        cellSize = desc.MeanCellWidth

        if desc.format == 'Image Service':
            if sr.Type == "Geographic":
                newLinearUnits,newZfactor,demExtract = extractDEMfromImageService(inputDEM,zUnits)
                return newLinearUnits,newZfactor,demExtract
            bImageService = True

        # ------------------------------------------------------------------------------------ Check DEM properties
        if bImageService:
            AddMsgAndPrint("\nInput DEM Image Service: " + desc.baseName)
        else:
            AddMsgAndPrint("\nInput DEM Information: " + desc.baseName)

        # DEM must be in a project coordinate system to continue
        # unless DEM is an image service.
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
            #tolerance = 3
            tolerance = 3.1
        elif linearUnits == "Foot":
            #tolerance = 9.84252
            tolerance = 10.1706
        elif linearUnits == "Foot_US":
            #tolerance = 9.84252
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
        #zFactor = zFactorList[unitLookUpDict.get(zUnits)][unitLookUpDict.get(linearUnits)]
        zFactor = zFactorList[unitLookUpDict.get(linearUnits)][unitLookUpDict.get(zUnits)]

        # ------------------------------------------------------------------------------------ Print Input DEM properties
        AddMsgAndPrint("\tProjection Name: " + sr.Name)
        AddMsgAndPrint("\tLinear Units (XY): " + linearUnits)
        AddMsgAndPrint("\tElevation Units (Z): " + zUnits)
        if not bZunits: AddMsgAndPrint("\t\tZ-units were auto set to: " + linearUnits,0)
        AddMsgAndPrint("\tCell Size: " + str(desc.MeanCellWidth) + " " + linearUnits)
        AddMsgAndPrint("\tZ-Factor: " + str(zFactor))

        # ------------------------------------------------------------------------------------ Extract DEM
        arcpy.SetProgressorLabel("Buffering AOI by 410 Meters")
        cluBuffer = "in_memory" + os.sep + os.path.basename(arcpy.CreateScratchName("cluBuffer",data_type="FeatureClass",workspace=scratchWS))
        #cluBuffer = arcpy.CreateScratchName("cluBuffer",data_type="FeatureClass",workspace=scratchWS)
        arcpy.Buffer_analysis(fieldDetermination,cluBuffer,"410 Meters","FULL","ROUND")
        arcpy.env.extent = cluBuffer

        # CLU clip extents
        cluExtent = arcpy.Describe(cluBuffer).extent
        clipExtent = str(cluExtent.XMin) + " " + str(cluExtent.YMin) + " " + str(cluExtent.XMax) + " " + str(cluExtent.YMax)

        # Cell Resolution needs to change; Clip and Project
        if bResample:
            arcpy.SetProgressorLabel("Changing resolution from " + str(cellSize) + " " + linearUnits + " to 3 Meters")
            AddMsgAndPrint("\n\tChanging resolution from " + str(cellSize) + " " + linearUnits + " to 3 Meters")

            demClip = "in_memory" + os.sep + os.path.basename(arcpy.CreateScratchName("demClip_resample",data_type="RasterDataset",workspace=scratchWS))
            arcpy.Clip_management(inputDEM, clipExtent, demClip, "", "", "", "NO_MAINTAIN_EXTENT")

            #demExtract = arcpy.CreateScratchName("demClip_project",data_type="RasterDataset",workspace=scratchWS)
            demExtract = "in_memory" + os.sep + os.path.basename(arcpy.CreateScratchName("demClip_project",data_type="RasterDataset",workspace=scratchWS))
            arcpy.ProjectRaster_management(demClip,demExtract,arcpy.env.outputCoordinateSystem,arcpy.env.resamplingMethod,
                                           outputCellSize,"#","#",sr)

            arcpy.Delete_management(demClip)

        # Resolution is correct; Clip the raster
        else:
            arcpy.SetProgressorLabel("Clipping DEM using buffered CLU")
            AddMsgAndPrint("\n\tClipping DEM using buffered CLU")

            #demExtract = arcpy.CreateScratchName("demClip",data_type="RasterDataset",workspace=scratchWS)
            demExtract = "in_memory" + os.sep + os.path.basename(arcpy.CreateScratchName("demClip",data_type="RasterDataset",workspace=scratchWS))
            arcpy.Clip_management(inputDEM, clipExtent, demExtract, "", "", "", "NO_MAINTAIN_EXTENT")

        # ------------------------------------------------------------------------------------ Report any new DEM properties
        desc = arcpy.Describe(demExtract)
        newSR = desc.SpatialReference
        newLinearUnits = newSR.LinearUnitName
        newCellSize = desc.MeanCellWidth
        #newZfactor = zFactorList[unitLookUpDict.get(zUnits)][unitLookUpDict.get(newLinearUnits)]
        newZfactor = zFactorList[unitLookUpDict.get(newLinearUnits)][unitLookUpDict.get(zUnits)]

        if newSR.name != sr.Name:
            AddMsgAndPrint("\t\tNew Projection Name: " + newSR.Name,0)
        if newCellSize != cellSize:
            AddMsgAndPrint("\t\tNew Cell Size: " + str(newCellSize) + " " + newLinearUnits )
        if newZfactor != zFactor:
            AddMsgAndPrint("\t\tNew Z-Factor: " + str(newZfactor))

        arcpy.Delete_management(cluBuffer)
        return newLinearUnits,newZfactor,demExtract

    except:
        errorMsg()
        return False,False,False

## ================================================================================================================
def removeScratchLayers():
    # This function is the last task that is executed or gets invoked in
    # an except clause.  Simply removes all temporary scratch layers.
    #import itertools

    try:
        #for lyr in list(itertools.chain(*scratchLayers)):
        #arcpy.AddMessage("The scratch layers are: " + str(scratchLayers))
        #for lyr in scratchLayers:
            #arcpy.AddMessage("Layer name: " + str(lyr) + " is in the scratchLayers list.")
        for lyr in scratchLayers:
            #if arcpy.Exists(lyr):
            try:
                #arcpy.AddMessage("Deleting Layer: " + str(lyr))
                arcpy.Delete_management(lyr)
            except:
                arcpy.AddMessage("Deleting Layer: " + str(lyr) + " failed.")
                continue
    except:
        pass

## ================================================================================================================
def AddLayersToArcMap():
    # This Function will add necessary layers to ArcMap.  Nothing is returned
    # If no PHEL Values were present only 2 layers will be added: Initial HEL Summary
    # and Field Determination.  If PHEL values were processed than all 4 layers
    # will be added.  This function does not utilize the setParameterastext function
    # to add layers to arcmap through the toolbox.

    try:
        # ------------------------------------------------------------------------ Post Process layers
        # Add 3 fields to the field determination layer and populate them
        # from ogCLUinfoDict and 4 fields to the Final HEL Summary layer that
        # otherwise would've been added after geoprocessing was successful.
        if bNoPHELvalues or bSkipGeoprocessing:

            # Add 3 fields to fieldDetermination layer
            fieldList = ["HEL_YES","HEL_Acres","HEL_Pct"]
            for field in fieldList:
                if not FindField(fieldDetermination,field):
                    if field == "HEL_YES":
                        arcpy.AddField_management(fieldDetermination,field,"TEXT","","",5)
                    else:
                        arcpy.AddField_management(fieldDetermination,field,"FLOAT")
            fieldList.append(cluNumberFld)

            # Update new fields using ogCLUinfoDict
            with arcpy.da.UpdateCursor(fieldDetermination,fieldList) as cursor:
                 for row in cursor:
                     row[0] = ogCLUinfoDict.get(row[3])[0]   # "HEL_YES" value
                     row[1] = ogCLUinfoDict.get(row[3])[1]   # "HEL_Acres" value
                     row[2] = ogCLUinfoDict.get(row[3])[2]   # "HEL_Pct" value
                     cursor.updateRow(row)

            # Add 4 fields to Final HEL Summary layer
            newFields = ['Polygon_Acres','Final_HEL_Value','Final_HEL_Acres','Final_HEL_Percent']
            for fld in newFields:
                if not len(arcpy.ListFields(finalHELSummary,fld)) > 0:
                   if fld == 'Final_HEL_Value':
                      arcpy.AddField_management(finalHELSummary,'Final_HEL_Value',"TEXT","","",5)
                   else:
                        arcpy.AddField_management(finalHELSummary,fld,"DOUBLE")

            newFields.append(helFld)
            newFields.append(cluNumberFld)
            newFields.append("SHAPE@AREA")

            # [polyAcres,finalHELvalue,finalHELacres,finalHELpct,MUHELCL,'CLUNBR',"SHAPE@AREA"]
            with arcpy.da.UpdateCursor(finalHELSummary,newFields) as cursor:
                for row in cursor:

                    # Calculate polygon acres;
                    row[0] = row[6] / acreConversionDict.get(arcpy.Describe(finalHELSummary).SpatialReference.LinearUnitName)

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
        mxd = arcpy.mapping.MapDocument("CURRENT")
        df = arcpy.mapping.ListDataFrames(mxd)[0]

        # redundant workaround.  ListLayers returns a list of layer objects
        # had to create a list of layer name Strings in order to see if a
        # specific layer currently exists in Arcmap.
        currentLayersObj = arcpy.mapping.ListLayers(mxd)
        currentLayersStr = [str(x) for x in arcpy.mapping.ListLayers(mxd)]

        # List of layers to add to Arcmap (layer path, arcmap layer name)
        if bNoPHELvalues or bSkipGeoprocessing:
            addToArcMap = [(helSummary,"Initial HEL Summary"),(fieldDetermination,"Field Determination")]

            # Remove these layers from arcmap if they are present since they were not produced
            if 'LiDAR HEL Summary' in currentLayersStr:
                arcpy.mapping.RemoveLayer(df,currentLayersObj[currentLayersStr.index('LiDAR HEL Summary')])

            if 'Final HEL Summary' in currentLayersStr:
                arcpy.mapping.RemoveLayer(df,currentLayersObj[currentLayersStr.index('Final HEL Summary')])

        else:
            addToArcMap = [(lidarHEL,"LiDAR HEL Summary"),(helSummary,"Initial HEL Summary"),(finalHELSummary,"Final HEL Summary"),(fieldDetermination,"Field Determination")]

        for layer in addToArcMap:

            # remove layer from ArcMap if it exists
            if layer[1] in currentLayersStr:
                arcpy.mapping.RemoveLayer(df,currentLayersObj[currentLayersStr.index(layer[1])])

            # Raster Layers need to be handled differently than vector layers
            if layer[1] == "LiDAR HEL Summary":
                rasterLayer = arcpy.MakeRasterLayer_management(layer[0],layer[1])
                tempLayer = rasterLayer.getOutput(0)
                arcpy.mapping.AddLayer(df,tempLayer,"TOP")

                # define the symbology layer and convert it to a layer object
                updateLayer = arcpy.mapping.ListLayers(mxd,layer[1], df)[0]
                symbologyLyr = os.path.join(os.path.dirname(sys.argv[0]),layer[1].lower().replace(" ","") + ".lyr")
                sourceLayer = arcpy.mapping.Layer(symbologyLyr)
                arcpy.mapping.UpdateLayer(df,updateLayer,sourceLayer)

            else:
                # add layer to arcmap
                symbologyLyr = os.path.join(os.path.dirname(sys.argv[0]),layer[1].lower().replace(" ","") + ".lyr")
                arcpy.mapping.AddLayer(df, arcpy.mapping.Layer(symbologyLyr.strip("'")), "TOP")

            # This layer should be turned on if no PHEL values were processed.
            # Symbology should also be updated to reflect current values.
            if layer[1] in ("Initial HEL Summary") and bNoPHELvalues:
                for lyr in arcpy.mapping.ListLayers(mxd, layer[1]):
                    lyr.visible = True

            # these 2 layers should be turned off by default if full processing happens
            if layer[1] in ("Initial HEL Summary","LiDAR HEL Summary") and not bNoPHELvalues:
                for lyr in arcpy.mapping.ListLayers(mxd, layer[1]):
                    lyr.visible = False

            AddMsgAndPrint("Added " + layer[1] + " to your ArcMap Session",0)

        # Unselect CLU polygons; Looks goofy after processed layers have been added to ArcMap
        # Turn it off as well
        #for lyr in arcpy.mapping.ListLayers(mxd, arcpy.Describe(cluLayer).nameString, df):
        for lyr in arcpy.mapping.ListLayers(mxd, "*" + str(arcpy.Describe(cluLayer).nameString).split("\\")[-1], df):
            arcpy.SelectLayerByAttribute_management(lyr, "CLEAR_SELECTION")
            lyr.visible = False

        # Turn off the original HEL layer to put the outputs into focus
        helLyr = arcpy.mapping.ListLayers(mxd, arcpy.Describe(helLayer).nameString, df)[0]
        helLyr.visible = False

        # set dataframe extent to the extent of the Field Determintation layer buffered by 50 meters.
        fieldDeterminationBuffer = "in_memory" + os.sep + os.path.basename(arcpy.CreateScratchName("fdBuffer",data_type="FeatureClass",workspace=scratchWS))
        arcpy.Buffer_analysis(fieldDetermination, fieldDeterminationBuffer, "50 Meters", "FULL", "", "ALL", "")
        # fdLayer = arcpy.mapping.ListLayers(mxd, "Field Determination", df)[0]
        df.extent = arcpy.Describe(fieldDeterminationBuffer).extent
        arcpy.Delete_management(fieldDeterminationBuffer)

        arcpy.RefreshTOC()
        arcpy.RefreshActiveView()

    except:
        errorMsg()

## ================================================================================================================
def populateForm():
    # This function will prepare the 1026 form by adding 16 fields to the
    # fieldDetermination feauture class.  This function will still be invoked even if
    # there was no geoprocessing executed due to no PHEL values to be computed.
    # if bNoPHELvalues is true, 3 fields will be added that otherwise would've
    # been added after HEL geoprocessing.  However, values to these fields will
    # be populated from the ogCLUinfoDict and represent original HEL values
    # primarily determined by the 33.33% or 50 acre rule

    try:

        AddMsgAndPrint("\nPreparing and Populating NRCS-CPA-026 Form", 0)

        # ---------------------------------------------------------------------  Collect Time information
        today = datetime.date.today()
        today = today.strftime('%b %d, %Y')

        # ---------------------------------------------------------------------- Get State Abbreviation
        stateCodeDict = {'WA': '53', 'DE': '10', 'DC': '11', 'WI': '55', 'WV': '54', 'HI': '15',
                        'FL': '12', 'WY': '56', 'PR': '72', 'NJ': '34', 'NM': '35', 'TX': '48',
                        'LA': '22', 'NC': '37', 'ND': '38', 'NE': '31', 'TN': '47', 'NY': '36',
                        'PA': '42', 'AK': '02', 'NV': '32', 'NH': '33', 'VA': '51', 'CO': '08',
                        'CA': '06', 'AL': '01', 'AR': '05', 'VT': '50', 'IL': '17', 'GA': '13',
                        'IN': '18', 'IA': '19', 'MA': '25', 'AZ': '04', 'ID': '16', 'CT': '09',
                        'ME': '23', 'MD': '24', 'OK': '40', 'OH': '39', 'UT': '49', 'MO': '29',
                        'MN': '27', 'MI': '26', 'RI': '44', 'KS': '20', 'MT': '30', 'MS': '28',
                        'SC': '45', 'KY': '21', 'OR': '41', 'SD': '46'}

        # Try to get the state using the field determination layer
        try:
            stateCode = ([row[0] for row in arcpy.da.SearchCursor(fieldDetermination,"STATECD")])
            state = [stAbbrev for (stAbbrev, code) in list(stateCodeDict.items()) if code == stateCode[0]][0]
        # Otherwise get the state from the computer user name
        except:
            state = getpass.getuser().replace('.',' ').replace('\'','')

        # Get the geographic and admin counties
        stFIPS = ''
        coFIPS = ''
        astFIPS = ''
        acoFIPS = ''

        # check that lookup table exists and search it for geographic and admin state and county codes
        if arcpy.Exists(lu_table):
            with arcpy.da.SearchCursor(fieldDetermination, ['statecd','countycd','admnstate','admncounty']) as cursor:
                for row in cursor:
                    stFIPS = row[0]
                    coFIPS = row[1]
                    astFIPS = row[2]
                    acoFIPS = row[3]
                    break

            # Lookup state and county names from recovered fips codes
            GeoCode = str(stFIPS) + str(coFIPS)
            expression1 = ("{} = '" + GeoCode + "'").format(arcpy.AddFieldDelimiters(lu_table, 'GEOID'))
            with arcpy.da.SearchCursor(lu_table, ['GEOID','NAME','STPOSTAL'], where_clause=expression1) as cursor:
                for row in cursor:
                    GeoCounty = row[1]
                    GeoState = row[2]
                    # We should only get one result if using installed lookup table from US Census Tiger table, so break
                    break

            AdminCode = str(astFIPS) + str(acoFIPS)
            expression2 = ("{} = '" + AdminCode + "'").format(arcpy.AddFieldDelimiters(lu_table, 'GEOID'))
            with arcpy.da.SearchCursor(lu_table, ['GEOID','NAME','STPOSTAL'], where_clause=expression2) as cursor:
                for row in cursor:
                    AdminCounty = row[1]
                    AdminState = row[2]
                    # We should only get one result if using installed lookup table from US Census Tiger table, so break
                    break

            del expression1, expression2

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

        arcpy.SetProgressor("step", "Preparing and Populating NRCS-CPA-026 Form", 0, len(fieldDict), 1)

        for field,params in fieldDict.items():
            arcpy.SetProgressorLabel("Adding Field: " + field + r' to "Field Determination" layer')
            try:
                fldLength = params[2]
            except:
                fldLength = 0
                pass

            arcpy.AddField_management(fieldDetermination,field,params[0],"#","#",fldLength)

            if len(params[1]) > 0:
                expression = "\"" + params[1] + "\""
                arcpy.CalculateField_management(fieldDetermination,field,expression,"VB")

        if bAccess:
            AddMsgAndPrint("\tOpening NRCS-CPA-026 Form",0)
            try:
                subprocess.Popen([msAccessPath,helDatabase])
            except:
                try:
                    os.startfile(helDatabase)
                except:
                    AddMsgAndPrint("\tCould not locate the Microsoft Access Software",1)
                    AddMsgAndPrint("\tOpen Microsoft Access manually to access the NRCS-CPA-026 Form",1)
                    arcpy.SetProgressorLabel("Could not locate the Microsoft Access Software")
        else:
            AddMsgAndPrint("\tOpening NRCS-CPA-026 Form",0)
            try:
                os.startfile(helDatabase)
            except:
                AddMsgAndPrint("\tCould not locate the Microsoft Access Software",1)
                AddMsgAndPrint("\tOpen Microsoft Access manually to access the NRCS-CPA-026 Form",1)
                arcpy.SetProgressorLabel("Could not locate the Microsoft Access Software")

        return True

    except:
        return False
        errorMsg()

## ================================================================================================================
def configLayout():
    # This function will gather and update information for elements of the map layout

    try:
        # Confirm Map Doc existence
        mxd = arcpy.mapping.MapDocument("CURRENT")

        # Set starting variables as empty for logical comparison later on, prior to updating layout
        farmNum = ''
        trNum = ''
        #input_cust comes from input parameters to tool in which this function was built
        county = ''
        state = ''

        # end function if lookup table from tool parameters does not exist
        if not arcpy.Exists(lu_table):
            return False

        # Get CLU information from first row of the cluLayer.
        # All CLU records should have this info, so break after one record.
        with arcpy.da.SearchCursor(fieldDetermination, ['statecd','countycd','farmnbr','tractnbr']) as cursor:
            for row in cursor:
                stCD = row[0]
                coCD = row[1]
                farmNum = row[2]
                trNum = row[3]
                break

        # Lookup state and county name
        stco_code = str(stCD) + str(coCD)
        #expression = (u'{} = ' + stco_code).format(arcpy.AddFieldDelimiters(lu_table, 'GEOID'))
        expression = ("{} = '" + stco_code + "'").format(arcpy.AddFieldDelimiters(lu_table, 'GEOID'))
        with arcpy.da.SearchCursor(lu_table, ['GEOID','NAME','STPOSTAL'], where_clause=expression) as cursor:
            for row in cursor:
                county = row[1]
                state = row[2]
                # We should only get one result if using installed lookup table from US Census Tiger table, so break
                break

        # Find and hook map layout elements to variables
        for elm in arcpy.mapping.ListLayoutElements(mxd):
            if elm.name == "farm_txt":
                farm_ele = elm
            if elm.name == "tract_txt":
                tract_ele = elm
            if elm.name == "customer_txt":
                customer_ele = elm
            if elm.name == "county_txt":
                county_ele = elm
            if elm.name == "state_txt":
                state_ele = elm

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

## =========================================================== Main Body ========================================================
import sys, string, os, traceback, re
import arcpy, subprocess, getpass, time
from arcpy import env
from arcpy.sa import *
import importlib
importlib.reload(sys)
sys.setdefaultencoding('utf-8')


if __name__ == '__main__':

    try:
        cluLayer = arcpy.GetParameter(0)
        helLayer = arcpy.GetParameter(1)
        inputDEM = arcpy.GetParameter(2)
        zUnits = arcpy.GetParameterAsText(3)
        dcSignature = arcpy.GetParameterAsText(4)
        input_cust = arcpy.GetParameterAsText(5)
        use_runoff_ls = arcpy.GetParameter(6)
        #use_runoff_ls = False      ## Used for testing when not set as a parameter

        ## 8/20/2019 - Observation: stateThreshold is not used anywhere in the code. Was intended as a PHEL predominance % paramter.
        ##Commenting it out for now.
        #stateThreshold = 50

        kFactorFld = "K"
        tFactorFld = "T"
        rFactorFld = "R"
        helFld = "MUHELCL"

        bLog = False # boolean to begin logging to text file.
        arcpy.SetProgressorLabel("Checking input values and environments")
        AddMsgAndPrint("\nChecking input values and environments")

        ## ----------------------------------------------------------------------------------------------------------------- Check HEL Access Database"""
        # --------------------------------------------------- Make sure the HEL access database is present, Exit otherwise
        helDatabase = os.path.dirname(sys.argv[0]) + os.sep + r'HEL.mdb'
        if not arcpy.Exists(helDatabase):
            AddMsgAndPrint("\nHEL Access Database does not exist in the same path as HEL Tools",2)
            sys.exit()
        # Also define the lu_table, but it's still ok to continue if it's not present
        lu_table = os.path.dirname(sys.argv[0]) + os.sep + r'census_fips_lut.dbf'

        #-------------------------------------------------------- Close Microsoft Access Database software if it is open.
##        tasks = os.popen('tasklist /v').read().strip().split('\n')
##        for i in range(len(tasks)):
##            task = tasks[i]
##            if 'MSACCESS.EXE' in task:
##                try:
##                    os.system("taskkill /f /im MSACCESS.EXE"); time.sleep(2)
##                    AddMsgAndPrint("\nMicrosoft Access was closed in order to continue")
##                except:
##                    AddMsgAndPrint("\nMicrosoft Access could not be closed")
##                break
        # forcibly kill image name msaccess if open.
        # remove access record-locking information
        try:
            killAccess = os.system("TASKKILL /F /IM msaccess.exe")
            if killAccess == 0:
                AddMsgAndPrint("\tMicrosoft Access was closed in order to continue")

            accessLockFile = os.path.dirname(sys.argv[0]) + os.sep + r'HEL.ldb'
            if os.path.exists(accessLockFile):
               os.remove(accessLockFile)
            time.sleep(2)
        except:
            time.sleep(2)
            pass

        # ---------------------------------------------------------------------- establish path to access database layers
        fieldDetermination = os.path.join(helDatabase, r'Field_Determination')
        helSummary = os.path.join(helDatabase, r'Initial_HEL_Summary')
        lidarHEL = os.path.join(helDatabase, r'LiDAR_HEL_Summary')
        #lidarHEL = os.path.dirname(sys.argv[0]) + os.sep + r'LiDAR_HEL_Summary.img'
        finalHELSummary = os.path.join(helDatabase, r'Final_HEL_Summary')
        accessLayers = [fieldDetermination,helSummary,lidarHEL,finalHELSummary]

        for layer in accessLayers:
            if arcpy.Exists(layer):
                try:
                    arcpy.Delete_management(layer)
                except:
                    AddMsgAndPrint("\tCould not delete the " + os.path.basename(layer) + " feature class in the HEL access database. Creating an additional layer",2)
                    newName = str(layer)
                    newName = arcpy.CreateScratchName(os.path.basename(layer),data_type="FeatureClass",workspace=helDatabase)

        # ---------------------------------------------------------- determine Microsoft Access path from windows version
        bAccess = True
        winVersion = sys.getwindowsversion()

        # Windows 10
        if winVersion.build == 9200:
            msAccessPath = r'C:\Program Files (x86)\Microsoft Office\root\Office16\MSACCESS.EXE'
        # Windows 7
        elif winVersion.build == 7601:
            msAccessPath = r'C:\Program Files (x86)\Microsoft Office\Office15\MSACCESS.EXE'
        else:
            AddMsgAndPrint("\nCould not determine Windows version, will not populate 026 Form",2)
            bAccess = False

        if bAccess and not os.path.isfile(msAccessPath):
            bAccess = False

        ## ------------------------------------------------------------------------------------ Checkout Spatial Analyst Extension and set scratch workspace """
        # Check Availability of Spatial Analyst Extension
        try:
            if arcpy.CheckExtension("Spatial") == "Available":
                arcpy.CheckOutExtension("Spatial")
            else:
                raise LicenseError

        except LicenseError:
            AddMsgAndPrint("\n\nSpatial Analyst license is unavailable.  Go to Customize -> Extensions to activate it",2)
            AddMsgAndPrint("\n\nExiting!")
            sys.exit()
        except arcpy.ExecuteError:
            AddMsgAndPrint(arcpy.GetMessages(2),2)
            sys.exit()

        # Set overwrite option
        arcpy.env.overwriteOutput = True

        # define and set the scratch workspace
        scratchWS = os.path.dirname(sys.argv[0]) + os.sep + r'scratch.gdb'
        if not arcpy.Exists(scratchWS):
            scratchWS = setScratchWorkspace()
        #else:
            #AddMsgAndPrint("\nScratch workspace set to " + scratchWS)

        if not scratchWS:
            AddMsgAndPrint("\nCould Not set scratchWorkspace!")
            sys.exit()

        arcpy.env.scratchWorkspace = scratchWS
        scratchLayers = list()

        ## -------------------------------------------------------------------------------------------------------- Stamp CLU into field determination fc.
        # ------------------------------------------------------------------ Exit if no CLU fields selected
        cluDesc = arcpy.Describe(cluLayer)
        if cluDesc.FIDset == '':
            AddMsgAndPrint("\nPlease select fields from the CLU Layer. Exiting!",2)
            sys.exit()
        else:
            fieldDetermination = arcpy.CopyFeatures_management(cluLayer,fieldDetermination)

        # -------------------------------------------- Make sure TRACTNBR and FARMNBR  are unique; exit otherwise
        uniqueTracts = list(set([row[0] for row in arcpy.da.SearchCursor(fieldDetermination,("TRACTNBR"))]))
        uniqueFarm   = list(set([row[0] for row in arcpy.da.SearchCursor(fieldDetermination,("FARMNBR"))]))
        uniqueFields = list(set([row[0] for row in arcpy.da.SearchCursor(fieldDetermination,("CLUNBR"))]))

        if len(uniqueTracts) != 1:
           AddMsgAndPrint("\n\tThere are " + str(len(uniqueTracts)) + " different Tract Numbers. Exiting!",2)
           for tract in uniqueTracts:
               AddMsgAndPrint("\t\tTract #: " + str(tract),2)
           removeScratchLayers
           sys.exit()

        if len(uniqueFarm) != 1:
           AddMsgAndPrint("\n\tThere are " + str(len(uniqueFarm)) + " different Farm Numbers. Exiting!",2)
           for farm in uniqueFarm:
               AddMsgAndPrint("\t\tFarm #: " + str(farm),2)
           removeScratchLayers
           sys.exit()

        # Create Text file to log info to
        textFilePath = createTextFile(uniqueTracts[0],uniqueFarm[0],uniqueFields)
        bLog = True

        # Update the map layout for the current site being run
        configLayout()

        AddMsgAndPrint("\nNumber of CLU fields selected: {}".format(len(cluDesc.FIDset.split(";"))))

        # Add Calcacre field if it doesn't exist. Should be part of the CLU layer.
        calcAcreFld = "CALCACRES"
        if not len(arcpy.ListFields(fieldDetermination,calcAcreFld)) > 0:
            arcpy.AddField_management(fieldDetermination,calcAcreFld,"DOUBLE")

        # Note: FSA MIDAS uses "square meters * 0.0002471" based on NAD 83 for the current UTM Zone and then rounds to two decimal points to set its calc acres.
        # If we changed all calc acres formulas to match FSA's formula, we would have matching FSA acres, but slightly incorrect amounts.
        # The variance is approximately two hundred thouandths of an acre or about 3/10ths of a square inch per acre.
        # If we set all internal acres computations to 2 decimal places from rounding based on the above, all acres would be consistent, except possibly for raster derived acres (need to check).
        arcpy.CalculateField_management(fieldDetermination,calcAcreFld,"!shape.area@acres!","PYTHON_9.3")
        totalAcres = float("%.1f" % (sum([row[0] for row in arcpy.da.SearchCursor(fieldDetermination, (calcAcreFld))])))
        AddMsgAndPrint("\tTotal Acres: " + splitThousands(totalAcres))

        ## ------------------------------------------------------------------------------------------------------------ Z-factor conversion Lookup table
        # lookup dictionary to convert XY units to area.  Key = XY unit of DEM; Value = conversion factor to sq.meters
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

        ### ------------------------------------------------------------------------------------------------------------- Compute Summary of original HEL values"""
        # -------------------------------------------------------------------------- Intersect fieldDetermination (CLU & AOI) with soils (helLayer) -> finalHELSummary
        AddMsgAndPrint("\nComputing summary of original HEL Values")
        arcpy.SetProgressorLabel("Computing summary of original HEL Values")
        cluHELintersect_pre = "in_memory" + os.sep + os.path.basename(arcpy.CreateScratchName("cluHELintersect_pre",data_type="FeatureClass",workspace=scratchWS))

        # Use the catalog path of the hel layer to avoid using a selection
        helLayerPath = arcpy.Describe(helLayer).catalogPath

        # Intersect fieldDetermination with soils and explode into single part
        arcpy.Intersect_analysis([fieldDetermination,helLayerPath],cluHELintersect_pre,"ALL")
        arcpy.MultipartToSinglepart_management(cluHELintersect_pre,finalHELSummary)
        scratchLayers.append(cluHELintersect_pre)

        # Test intersection --- Should we check the percentage of intersection here? what if only 50% overlap
        # No modification needed for these acres. The total is used only for this check.
        totalIntAcres = sum([row[0] for row in arcpy.da.SearchCursor(finalHELSummary, ("SHAPE@AREA"))]) / acreConversionDict.get(arcpy.Describe(finalHELSummary).SpatialReference.LinearUnitName)
        if not totalIntAcres:
            AddMsgAndPrint("\tThere is no overlap between HEL soil layer and CLU Layer. EXITTING!",2)
            removeScratchLayers()
            sys.exit()

        # -------------------------------------------------------------------------- Dissolve intersection output by the following fields -> helSummary
        cluNumberFld = "CLUNBR"
        dissovleFlds = [cluNumberFld,"TRACTNBR","FARMNBR","COUNTYCD","CALCACRES",helFld]

        # Dissolve the finalHELSummary to report input summary
        arcpy.Dissolve_management(finalHELSummary, helSummary, dissovleFlds, "","MULTI_PART", "DISSOLVE_LINES")

        # --------------------------------------------------------------------------- Add and Update fields in the HEL Summary Layer (Og_HELcode, Og_HEL_Acres, Og_HEL_AcrePct)
        # Add 3 fields to the intersected layer.  The intersected 'clueHELintersect' layer will
        # be used for the dissolve process and at the end of the script.
        HELrasterCode = 'Og_HELcode'    # Used for rasterization purposes
        HELacres = 'Og_HEL_Acres'
        HELacrePct = 'Og_HEL_AcrePct'

        if not len(arcpy.ListFields(helSummary,HELrasterCode)) > 0:
            arcpy.AddField_management(helSummary,HELrasterCode,"SHORT")

        if not len(arcpy.ListFields(helSummary,HELacres)) > 0:
            arcpy.AddField_management(helSummary,HELacres,"DOUBLE")

        if not len(arcpy.ListFields(helSummary,HELacrePct)) > 0:
            arcpy.AddField_management(helSummary,HELacrePct,"DOUBLE")

        # Calculate HELValue Field
        helSummaryDict = dict()     ## tallies acres by HEL value i.e. {PHEL:100}
        nullHEL = 0                 ## # of polygons with no HEL values
        wrongHELvalues = list()     ## Stores incorrect HEL Values
        maxAcreLength = list()      ## Stores the number of acre digits for formatting purposes
        bNoPHELvalues = False       ## Boolean flag to indicate PHEL values are missing

        # HEL Field, Og_HELcode, Og_HEL_Acres, Og_HEL_AcrePct, "SHAPE@AREA", "CALCACRES"
        with arcpy.da.UpdateCursor(helSummary,[helFld,HELrasterCode,HELacres,HELacrePct,"SHAPE@AREA",calcAcreFld]) as cursor:
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
                #acres = float("%.1f" % (row[3] / acreConversionDict.get(arcpy.Describe(helSummary).SpatialReference.LinearUnitName)))
                acres = row[4] / acreConversionDict.get(arcpy.Describe(helSummary).SpatialReference.LinearUnitName)
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
            #AddMsgAndPrint("\n\tWARNING: There are no PHEL values in HEL layer",1)
            bNoPHELvalues = True

        # Inform user about NULL values; Exit if any NULLs exist.
        if nullHEL > 0:
            AddMsgAndPrint("\n\tERROR: There are " + str(nullHEL) + " polygon(s) with missing HEL values. EXITING!",2)
            removeScratchLayers()
            sys.exit()

        # Inform user about invalid HEL values (not PHEL,HEL, NHEL); Exit if invalid values exist.
        if wrongHELvalues:
            AddMsgAndPrint("\n\tERROR: There is " + str(len(set(wrongHELvalues))) + " invalid HEL values in HEL Layer:",1)
            for wrongVal in set(wrongHELvalues):
                AddMsgAndPrint("\t\t" + wrongVal)
            removeScratchLayers()
            sys.exit()

        del dissovleFlds,nullHEL,wrongHELvalues

        ### --------------------------------------------------------------------------------------------------------- Report HEl Layer Summary by field
        AddMsgAndPrint("\n\tSummary by CLU:")

        # Create 2 temporary tables to capture summary statistics
        ogHelSummaryStats = "in_memory" + os.sep + os.path.basename(arcpy.CreateScratchName("ogHELSummaryStats",data_type="ArcInfoTable",workspace=scratchWS))
        ogHelSummaryStatsPivot = "in_memory" + os.sep + os.path.basename(arcpy.CreateScratchName("ogHELSummaryStatsPivot",data_type="ArcInfoTable",workspace=scratchWS))
        #ogHelSummaryStats = arcpy.CreateScratchName("ogHELSummaryStats",data_type="ArcInfoTable",workspace=scratchWS)
        #ogHelSummaryStatsPivot = arcpy.CreateScratchName("ogHELSummaryStatsPivot",data_type="ArcInfoTable",workspace=scratchWS)

        stats = [[HELacres,"SUM"]]
        caseField = [cluNumberFld,helFld]
        arcpy.Statistics_analysis(helSummary, ogHelSummaryStats, stats, caseField)
        sumHELacreFld = [fld.name for fld in arcpy.ListFields(ogHelSummaryStats,"*" + HELacres)][0]
        scratchLayers.append(ogHelSummaryStats)

        # Pivot table will have CLUNBR & any HEL values present (HEL,NHEL,PHEL)
        arcpy.PivotTable_management(ogHelSummaryStats,cluNumberFld,helFld,sumHELacreFld,ogHelSummaryStatsPivot)
        scratchLayers.append(ogHelSummaryStatsPivot)

        pivotFields = [fld.name for fld in arcpy.ListFields(ogHelSummaryStatsPivot)][1:]  # ['CLUNBR','HEL','NHEL','PHEL']
        numOfhelValues = len(pivotFields)                                                 # Number of Pivot table fields; Min 2 fields
        maxAcreLength.sort(reverse=True)
        bSkipGeoprocessing = True             # Skip processing until a field is neither HEL >= 33.33% or NHEL > 66.67%

        # This dictionary will only be used if FINAL results are all HEL or all NHEL to reference original
        # acres and not use tabulate area acres.  It will also be used when there are no PHEL Values.
        # {cluNumber:(HEL value, cluAcres, HEL Pct} -- HEL value is determined by the 33.33% or 50 acre rule
        ogCLUinfoDict = dict()

        # Iterate through the pivot table and report HEL values by CLU - ['CLUNBR','HEL','NHEL','PHEL']
        with arcpy.da.SearchCursor(ogHelSummaryStatsPivot,pivotFields) as cursor:
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
                    #AddMsgAndPrint("\t\t\t" + pivotFields[i] + firstSpace + " -- " + str(acres) + secondSpace + " .ac -- " + str(pct) + " %")
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

        # ------------------------------------------------------------------------- No PHEL Values Found
        # If there are no PHEL Values add helSummary and fieldDetermination layers to ArcMap
        # and prepare 1026 form.  Skip geoProcessing.

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
            arcpy.SetProgressorLabel("")
            AddMsgAndPrint("\n")
            arcpy.RefreshCatalog(scratchWS)
            sys.exit()

        ### ---------------------------------------------------------------------------------------------- Check and create DEM clip from buffered CLU
        # Exit if a DEM is not present; At this point PHEL mapunits are present
        # and requires a DEM to process them.
        try:
            arcpy.Describe(inputDEM).baseName
        except:
            AddMsgAndPrint("\nDEM is required to process PHEL values. EXITING!")
            sys.exit()

        units,zFactor,dem = extractDEM(inputDEM,zUnits)
        if not zFactor or not dem:
           removeScratchLayers()
           sys.exit()
        scratchLayers.append(dem)

        ### ----------------------------------------------------------------------------------------------------------------------------- Set Snap Raster
        # Disabled for now due to shifting soils results.
        # Investigating change so that soil derived rasters exist first and are the snap rasters (9/16/2019)
        #arcpy.env.snapRaster = dem

        ### ------------------------------------------------------------------------------------ Check DEM for NoData overlaps with input CLU fields
        AddMsgAndPrint("\nChecking input DEM for site coverage...")
        vectorNull = "in_memory" + os.sep + os.path.basename(arcpy.CreateScratchName("vectorNull",data_type="FeatureClass",workspace=scratchWS))
        demCheck = "in_memory" + os.sep + os.path.basename(arcpy.CreateScratchName("demCheck",data_type="FeatureClass",workspace=scratchWS))

        # Use Set Null statement to change things with value of 0 to NoData
        whereClause = "VALUE = 0"
        setNull = SetNull(dem, dem, whereClause)
        scratchLayers.append(SetNull)

        # Use IsNull to convert NoData values in the DEM to 1 and all other values to 0
        demNull = IsNull(setNull)
        scratchLayers.append(demNull)

        # Convert the IsNull raster to a vector layer
        try:
            arcpy.RasterToPolygon_conversion(demNull, vectorNull, "SIMPLIFY", "Value", "MULTIPLE_OUTER_PART")
        except:
            arcpy.RasterToPolygon_conversion(demNull, vectorNull, "SIMPLIFY", "Value")
        scratchLayers.append(vectorNull)

        # Clip the IsNull vector layer by the field determination layer
        arcpy.Clip_analysis(vectorNull, fieldDetermination, demCheck)
        scratchLayers.append(demCheck)

        # Search for any values of 1 in the demCheck layer and issue a warning to the user if present
        fields = ['gridcode']
        cursor = arcpy.da.SearchCursor(demCheck, fields)
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

        ### -------------------------------------------------------------------------------------------------------------------------- Create Slope Layer
        # Perform a minor fill to reduce LiDAR data noise and minor irregularities.
        # Try to use a max fill height of no more than 1 foot, based on input zUnits.
        arcpy.SetProgressorLabel("Filling small sinks in DEM")
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

        #1 Perform the fill using the zLimit as the max fill amount
        filled = Fill(dem, zLimit)
        scratchLayers.append(filled)

        #2 Run a FocalMean to smooth the DEM of LiDAR data noise. This should be run prior to creating derivative products.
        # This replaces running FocalMean on the slope layer itself.
        arcpy.SetProgressorLabel("Running Focal Statistics on DEM")
        AddMsgAndPrint("Running Focal Statistics on DEM")
        #slope = arcpy.CreateScratchName("focStatsMean_DemFilled",data_type="RasterDataset",workspace=scratchWS)
        preslope = FocalStatistics(filled, NbrRectangle(3,3,"CELL"),"MEAN","DATA")
        #outFocalStatistics.save(slope)

        #3 Create Slope
        arcpy.SetProgressorLabel("Creating Slope Derivative")
        AddMsgAndPrint("\nCreating Slope Derivative")
        #preslope = arcpy.CreateScratchName("preslope",data_type="RasterDataset",workspace=scratchWS)
        slope = Slope(preslope,"PERCENT_RISE",zFactor)
        #outSlope.save(preslope)

###### REMOVED IN FAVOR OF FOCAL MEAN ON DEM PRIOR TO RUNNING SLOPE ######
##        # Run a FocalMean statistics on slope output
##        arcpy.SetProgressorLabel("Running Focal Statistics on Slope Percent")
##        AddMsgAndPrint("Running Focal Statistics on Slope Percent")
##        #slope = arcpy.CreateScratchName("focStatsMean_Slope",data_type="RasterDataset",workspace=scratchWS)
##        slope = FocalStatistics(preslope, NbrRectangle(3,3,"CELL"),"MEAN","DATA")
##        #outFocalStatistics.save(slope)
###### REMOVED IN FAVOR OF FOCAL MEAN ON DEM PRIOR TO RUNNING SLOPE ######

        ### ------------------------------------------------------------------------------------------------------------ Create Flow Direction and Flow Length
        #4 Calculate Flow direction
        arcpy.SetProgressorLabel("Calculating Flow Direction")
        AddMsgAndPrint("Calculating Flow Direction")
        #flowDirection = arcpy.CreateScratchName("flowDirection",data_type="RasterDataset",workspace=scratchWS)
        #flowDirection = FlowDirection(dem, "FORCE")
        flowDirection = FlowDirection(preslope, "FORCE")
        #outFlowDirection.save(flowDirection)
        scratchLayers.append(flowDirection)

        #5 Calculate Flow Length
        arcpy.SetProgressorLabel("Calculating Flow Length")
        AddMsgAndPrint("Calculating Flow Length")
        #preflowLength = arcpy.CreateScratchName("flowLength",data_type="RasterDataset",workspace=scratchWS)
        preflowLength = FlowLength(flowDirection,"UPSTREAM", "")
        scratchLayers.append(preflowLength)
        #outpreFlowLength.save(preflowLength)

        #6 Run a focal statistics on flow length
        arcpy.SetProgressorLabel("Running Focal Statistics on Flow Length")
        AddMsgAndPrint("Running Focal Statistics on Flow Length")
        #flowLength = arcpy.CreateScratchName("focStatsMax_FlowLength",data_type="RasterDataset",workspace=scratchWS)
        flowLength = FocalStatistics(preflowLength, NbrRectangle(3,3,"CELL"),"MAXIMUM","DATA")
        #outFocalStatistics.save(flowLength)
        scratchLayers.append(flowLength)

        #7 convert Flow Length distance units to feet if original DEM LINEAR UNITS ARE not in feet.
        # Change this zUnits reference!
        #if not zUnits in ('Feet','Foot','Foot_US'):
        if not units in ('Feet','Foot','Foot_US'):
            AddMsgAndPrint("Converting Flow Length Distance units to Feet")
            #flowLengthFT = arcpy.CreateScratchName("flowLength_FT",data_type="RasterDataset",workspace=scratchWS)
            #outflowLengthFT = Raster(flowLength) * 3.280839896
            flowLengthFT = flowLength * 3.280839896
            #outflowLengthFT.save(flowLengthFT)
            scratchLayers.append(flowLengthFT)

        else:
            flowLengthFT = flowLength
            scratchLayers.append(flowLengthFT)

        ### --------------------------------------------------------------------------------------------------------------- Calculate LS Factor

        #8 Convert slope percent to radians for use in various LS equations
        radians = ATan(Times(slope,0.01))

        # ----------------------------------------------------------------------------- Compute LS Factor
        # If Northwest US 'Use Runoff LS Equation' flag was active, use the following equation
        if use_runoff_ls:
            arcpy.SetProgressorLabel("Calculating LS Factor")
            AddMsgAndPrint("Calculating LS Factor")
            lsFactor = (Power((flowLengthFT/72.6)*Cos(radians),0.5))*(Power(Sin((radians))/(Sin(5.143*((math.pi)/180))),0.7))

        # Otherwise, use the standard AH537 LS computation
        else:
            #9 ------------------------------------------------------------------------------- Calculate S Factor
            arcpy.SetProgressorLabel("Calculating S Factor")
            AddMsgAndPrint("\nCalculating S Factor")
            # Compute S factor using formula in AH537, pg 12
            sFactor = ((Power(Sin(radians),2)*65.41)+(Sin(radians)*4.56)+(0.065))
            scratchLayers.append(sFactor)

            #10 ------------------------------------------------------------------------------ Calculate L Factor
            arcpy.SetProgressorLabel("Calculating L Factor")
            AddMsgAndPrint("Calculating L Factor")
            #lFactor = arcpy.CreateScratchName("lFactor",data_type="RasterDataset",workspace=scratchWS)

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

            #outlFactor.save(lFactor)
            scratchLayers.append(lFactor)

            # ----------------------------------------------------------------------------- Calculate LS Factor
            #11 "%l_factor%" * "%s_factor%"
            arcpy.SetProgressorLabel("Calculating LS Factor")
            AddMsgAndPrint("Calculating LS Factor")
            #lsFactor = arcpy.CreateScratchName("lsFactor",data_type="RasterDataset",workspace=scratchWS)
            #outlsFactor = Raster(lFactor) * Raster(sFactor)  ## Original Line
            lsFactor = lFactor * sFactor
            #outlsFactor.save(lsFactor)

        scratchLayers.append(radians)
        scratchLayers.append(lsFactor)

        ### ------------------------------------------------------------------------------------------------------------- Convert K,T & R Factor and HEL Value to Rasters
        AddMsgAndPrint("\nConverting Vector to Raster for Spatial Analysis Purpose")
        cellSize = arcpy.Describe(dem).MeanCellWidth

        # (Works in 10.5 and under, but not 10.6 and up) All raster datasets will be created in memory
##        kFactor =  "in_memory" + os.sep + os.path.basename(arcpy.CreateScratchName("kFactor",data_type="RasterDataset",workspace=scratchWS))
##        tFactor =  "in_memory" + os.sep + os.path.basename(arcpy.CreateScratchName("tFactor",data_type="RasterDataset",workspace=scratchWS))
##        rFactor =  "in_memory" + os.sep + os.path.basename(arcpy.CreateScratchName("rFactor",data_type="RasterDataset",workspace=scratchWS))
##        helValue = "in_memory" + os.sep + os.path.basename(arcpy.CreateScratchName("helValue",data_type="RasterDataset",workspace=scratchWS))

        # This works in 10.5 AND works in 10.6.1 and 10.7 but slows processing
        kFactor = arcpy.CreateScratchName("kFactor",data_type="RasterDataset",workspace=scratchWS)
        tFactor = arcpy.CreateScratchName("tFactor",data_type="RasterDataset",workspace=scratchWS)
        rFactor = arcpy.CreateScratchName("rFactor",data_type="RasterDataset",workspace=scratchWS)
        helValue = arcpy.CreateScratchName("helValue",data_type="RasterDataset",workspace=scratchWS)

        #12 Convert KFactor to raster
        arcpy.SetProgressorLabel("Converting K Factor field to a raster")
        AddMsgAndPrint("\tConverting K Factor field to a raster")
        arcpy.FeatureToRaster_conversion(finalHELSummary,kFactorFld,kFactor,cellSize)

        #13 Convert TFactor to raster
        arcpy.SetProgressorLabel("Converting T Factor field to a raster")
        AddMsgAndPrint("\tConverting T Factor field to a raster")
        arcpy.FeatureToRaster_conversion(finalHELSummary,tFactorFld,tFactor,cellSize)

        #14 Convert RFactor to raster
        arcpy.SetProgressorLabel("Converting R Factor field to a raster")
        AddMsgAndPrint("\tConverting R Factor field to a raster")
        arcpy.FeatureToRaster_conversion(finalHELSummary,rFactorFld,rFactor,cellSize)

        arcpy.SetProgressorLabel("Converting HEL Value field to a raster")
        AddMsgAndPrint("\tConverting HEL Value field to a raster")
        arcpy.FeatureToRaster_conversion(helSummary,HELrasterCode,helValue,cellSize)

        #scratchLayers.append((kFactor,tFactor,rFactor,helValue))
        scratchLayers.append(kFactor)
        scratchLayers.append(tFactor)
        scratchLayers.append(rFactor)
        scratchLayers.append(helValue)

        ### ------------------------------------------------------------------------------------------------------------- Calculate EI Factor
        arcpy.SetProgressorLabel("Calculating EI Factor")
        AddMsgAndPrint("\nCalculating EI Factor")
        #eiFactor = arcpy.CreateScratchName("eiFactor", data_type="RasterDataset", workspace=scratchWS)
        #outEIfactor = Divide((Raster(lsFactor) * Raster(kFactor) * Raster(rFactor)),Raster(tFactor))  # Original Lines
        eiFactor = Divide((lsFactor * kFactor * rFactor),tFactor)
        #outEIfactor.save(eiFactor)
        scratchLayers.append(eiFactor)

        ### ------------------------------------------------------------------------------------------------------------- Calculate Final HEL Factor
        # Create Conditional statement to reflect the following:
        # 1) PHEL Value = 0 -- Take EI factor -- Depends     2
        # 2) HEL Value  = 1 -- Assign 9                      0
        # 3) NHEL Value = 2 -- Assign 2 (No action needed)   1
        # Anything above 8 is HEL

        arcpy.SetProgressorLabel("Calculating HEL Factor")
        AddMsgAndPrint("Calculating HEL Factor")
        #helFactor = arcpy.CreateScratchName("helFactor",data_type="RasterDataset",workspace=scratchWS)
        #outHELfactor = Con(Raster(helValue),Raster(eiFactor),Con(Raster(helValue),9,Raster(helValue),"VALUE=0"),"VALUE=2")  ## Original Line
        helFactor = Con(helValue,eiFactor,Con(helValue,9,helValue,"VALUE=0"),"VALUE=2")
        scratchLayers.append(helFactor)
        #outHELfactor.save(helFactor)

        #lidarHEL = arcpy.CreateScratchName("lidarHEL",data_type="RasterDataset",workspace=scratchWS)
        # Reclassify values:
        #       < 8 = Value_1 = NHEL
        #       > 8 = Value_2 = HEL
        remapString = "0 8 1;8 100000000 2"
        arcpy.Reclassify_3d(helFactor, "VALUE", remapString, lidarHEL,'NODATA')

        ### ------------------------------------------------------------------------------------- Determine if individual PHEL delineations are HEL/NHEL"""
        arcpy.SetProgressorLabel("Computing summary of LiDAR HEL Values:")
        AddMsgAndPrint("\nComputing summary of LiDAR HEL Values:\n")

        # Summarize new values between HEL soil polygon and lidarHEL raster
        outPolyTabulate = "in_memory" + os.sep + os.path.basename(arcpy.CreateScratchName("HEL_Polygon_Tabulate",data_type="ArcInfoTable",workspace=scratchWS))
        #outPolyTabulate = arcpy.CreateScratchName("HEL_Polygon_Tabulate",data_type="ArcInfoTable",workspace=scratchWS)

        zoneFld = arcpy.Describe(finalHELSummary).OIDFieldName
        TabulateArea(finalHELSummary,zoneFld,lidarHEL,"VALUE",outPolyTabulate,cellSize)
        tabulateFields = [fld.name for fld in arcpy.ListFields(outPolyTabulate)][2:]
        scratchLayers.append(outPolyTabulate)

        # Add 4 fields to Final HEL Summary layer
        newFields = ['Polygon_Acres','Final_HEL_Value','Final_HEL_Acres','Final_HEL_Percent']

        for fld in newFields:
            if not len(arcpy.ListFields(finalHELSummary,fld)) > 0:
               if fld == 'Final_HEL_Value':
                  arcpy.AddField_management(finalHELSummary,'Final_HEL_Value',"TEXT","","",5)
               else:
                    arcpy.AddField_management(finalHELSummary,fld,"DOUBLE")

        # In some cases, the finalHELSummary layer's OID field name was "OBJECTID_1" which
        # conflicted with the output of the tabulate area table.
        if zoneFld.find("_") > -1:
            outputJoinFld = zoneFld
        else:
            outputJoinFld = zoneFld + "_1"
        arcpy.JoinField_management(finalHELSummary,zoneFld,outPolyTabulate,outputJoinFld,tabulateFields)

        # Booleans to indicate if only HEL or only NHEL is present
        bOnlyHEL = False; bOnlyNHEL = False

        # Check if VALUE_1(NHEL) or VALUE_2(HEL) are missing from outPolyTabulate table
        finalHELSummaryFlds = [fld.name for fld in arcpy.ListFields(finalHELSummary)][2:]
        if len(finalHELSummaryFlds):

            # NHEL is not Present - so All is HEL; All is VALUE2
            if not "VALUE_1" in tabulateFields:
                AddMsgAndPrint("\tWARNING: Entire Area is HEL",1)
                arcpy.AddField_management(finalHELSummary,"VALUE_1","DOUBLE")
                arcpy.CalculateField_management(finalHELSummary,"VALUE_1",0)
                bOnlyHEL = True

            # HEL is not Present - All is NHEL; All is VALUE1
            if not "VALUE_2" in tabulateFields:
                AddMsgAndPrint("\tWARNING: Entire Area is NHEL",1)
                arcpy.AddField_management(finalHELSummary,"VALUE_2","DOUBLE")
                arcpy.CalculateField_management(finalHELSummary,"VALUE_2",0)
                bOnlyNHEL = True
        else:
            AddMsgAndPrint("\n\tReclassifying helFactor Failed",2)
            sys.exit()

        newFields.append("VALUE_2")
        newFields.append("SHAPE@AREA")
        newFields.append(cluNumberFld)

        # this will be used for field determination
        fieldDeterminationDict = dict()

        # [polyAcres,finalHELvalue,finalHELacres,finalHELpct,"VALUE_2","SHAPE@AREA","CLUNBR"]
        with arcpy.da.UpdateCursor(finalHELSummary,newFields) as cursor:
            for row in cursor:

                # Calculate polygon acres
                row[0] = row[5] / acreConversionDict.get(arcpy.Describe(finalHELSummary).SpatialReference.LinearUnitName)

                # Convert "VALUE_2" values to acres.  Represent acres from a poly that is HEL.
                # The intersection of CLU and soils may cause slivers below the tabulate cell size
                # which will create NULLs.  Set these slivers to 0 acres.
                try:
                    row[2] = row[4] / acreConversionDict.get(arcpy.Describe(finalHELSummary).SpatialReference.LinearUnitName)
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
        for fld in [f.name for f in arcpy.ListFields(finalHELSummary)]:
            if fld in (zoneFld,'Shape_Area','Shape_Length','Shape'):continue
            if not fld in validFlds:
                deleteFlds.append(fld)

        arcpy.DeleteField_management(finalHELSummary,deleteFlds)
        del zoneFld,finalHELSummaryFlds,tabulateFields,newFields,validFlds

        ### ---------------------------------------------------------------------------------------------------- Determine if field is HEL/NHEL"""
        # Add 3 fields to fieldDetermination layer
        fieldList = ["HEL_YES","HEL_Acres","HEL_Pct"]
        for field in fieldList:
            if not FindField(fieldDetermination,field):
                if field == "HEL_YES":
                    arcpy.AddField_management(fieldDetermination,field,"TEXT","","",5)
                else:
                    arcpy.AddField_management(fieldDetermination,field,"FLOAT")

        fieldList.append(cluNumberFld)
        fieldList.append(calcAcreFld)
        cluDict = dict()  # Strictly for formatting; ClUNBR: (len of clu, helAcres, helPct, len of Acres, len of pct,is it HEL?)

        # ['HEL_YES','HEL_Acres','HEL_Pct','CLUNBR','CALCACRES']
        with arcpy.da.UpdateCursor(fieldDetermination,fieldList) as cursor:
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
        #maxPercentLength = sorted([cluinfo[4] for clu,cluinfo in cluDict.iteritems()],reverse=True)[0]

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

        """----------------------------------------------------------------------------------------------------- Prepare Symboloby for ArcMap and 1026 form"""
        AddLayersToArcMap()

        if not populateForm():
            AddMsgAndPrint("\nFailed to correclty populate NRCS-CPA-026 form",2)

        # Clean up time
        removeScratchLayers()
        arcpy.SetProgressorLabel("")
        AddMsgAndPrint("\n")
        arcpy.RefreshCatalog(scratchWS)

    except:
        removeScratchLayers()
        errorMsg()
