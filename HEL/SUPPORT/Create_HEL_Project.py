from datetime import datetime
from getpass import getuser
from os import mkdir, path
from sys import argv, exit
from time import ctime
from uuid import uuid4

from arcpy import AddFieldDelimiters, Describe, env, Exists, GetParameter, GetParameterAsText, \
    ListFields, SetProgressorLabel, SpatialReference

from arcpy.conversion import FeatureClassToFeatureClass
from arcpy.da import SearchCursor, UpdateCursor

from arcpy.management import AddField, AlterDomain, Append, CalculateField, Compact, CreateFeatureclass, \
    CreateFeatureDataset, CreateFileGDB, Delete, Dissolve, TableToDomain

from arcpy.mp import ArcGISProject, LayerFile

from extract_CLU_by_Tract import getPortalTokenInfo, start
from hel_utils import addLyrxByConnectionProperties, AddMsgAndPrint, errorMsg


textFilePath = ''
def logBasicSettings(textFilePath, projectType, sourceState, sourceCounty, tractNumber, owFlag):
    with open(textFilePath, 'a+') as f:
        f.write('\n######################################################################\n')
        f.write('Executing Tool: Create HEL Project\n')
        f.write(f"User Name: {getuser()}\n")
        f.write(f"Date Executed: {ctime()}\n")
        f.write('User Parameters:\n')
        f.write(f"\tProject Type: {projectType}\n")
        f.write(f"\tAdmin State: {sourceState}\n")
        f.write(f"\tAdmin County: {sourceCounty}\n")
        f.write(f"\tTract: {str(tractNumber)}\n")
        f.write(f"\tOverwrite CLU: {str(owFlag)}\n")


### Initial Tool Validation ###
try:
    aprx = ArcGISProject('CURRENT')
    map = aprx.listMaps('HEL Determination')[0]
except:
    AddMsgAndPrint('This tool must be run from an ArcGIS Pro project that was developed from the template distributed with this toolbox. Exiting...', 2)
    exit()

nrcsPortal = 'https://gis.sc.egov.usda.gov/portal/'
portalToken = getPortalTokenInfo(nrcsPortal)
if not portalToken:
    AddMsgAndPrint('Could not generate Portal token. Please login to GeoPortal. Exiting...', 2)
    exit()

lut = path.join(path.dirname(argv[0]), 'SUPPORT.gdb', 'lut_census_fips')
if not Exists(lut):
    AddMsgAndPrint('Could not find state and county lookup table. Exiting...', 2)
    exit()


### Validate Spatial Reference ###
mapSR = map.spatialReference
if mapSR.type != 'Projected':
    AddMsgAndPrint('\nThe Determinations map is not set to a Projected coordinate system.', 2)
    AddMsgAndPrint('\nPlease assign a WGS 1984 UTM coordinate system to the Determinations map that is appropriate for your site.', 2)
    AddMsgAndPrint('\nThese systems are found in the Determinations Map Properties under: Coordinate Systems -> Projected Coordinate System -> UTM -> WGS 1984.', 2)
    AddMsgAndPrint('\nAfter applying a coordinate system, save your template and try this tool again.', 2)
    AddMsgAndPrint('\nExiting...', 2)
    exit()

if 'WGS' not in mapSR.name or '1984' not in mapSR.name or 'UTM' not in mapSR.name:
    AddMsgAndPrint('\nThe Determinations map is not using a UTM coordinate system tied to WGS 1984.', 2)
    AddMsgAndPrint('\nPlease assign a WGS 1984 UTM coordinate system to the Determinations map that is appropriate for your site.', 2)
    AddMsgAndPrint('\nThese systems are found in the Determinations Map Properties under: Coordinate Systems -> Projected Coordinate System -> UTM -> WGS 1984.', 2)
    AddMsgAndPrint('\nAfter applying a coordinate system, save your template and try this tool again.', 2)
    AddMsgAndPrint('\nExiting...', 2)
    exit()


### ESRI Environment Settings ###
mapSR = SpatialReference(mapSR.factoryCode)
env.outputCoordinateSystem = mapSR
env.overwriteOutput = True


### Input Parameters ###
projectType = GetParameterAsText(0)
existingFolder = GetParameterAsText(1)
sourceState = GetParameterAsText(2)
sourceCounty = GetParameterAsText(3)
tractNumber = GetParameterAsText(4)
owFlag = GetParameter(5)


try:
    workspacePath = 'C:\Determinations'
    # Check Inputs for existence and create FIPS code variables
    # Search for FIPS codes to give to the Extract CLU Tool/Function. Break after the first row (should only find one row in any case).
    # Temporarily adjust source county to handle apostrophes in relation to searching
    sourceCounty = sourceCounty.replace("'", "''")
    stfip, cofip = '', ''
    fields = ['STATEFP','COUNTYFP','NAME','STATE','STPOSTAL']
    field1 = 'STATE'
    field2 = 'NAME'
    expression = "{} = '{}'".format(AddFieldDelimiters(lut,field1), sourceState) + " AND " + "{} = '{}'".format(AddFieldDelimiters(lut,field2), sourceCounty)
    with SearchCursor(lut, fields, where_clause = expression) as cursor:
        for row in cursor:
            stfip = row[0]
            cofip = row[1]
            adStatePostal = row[4]
            break

    if len(stfip) != 2 and len(cofip) != 3:
        AddMsgAndPrint('State and County FIPS codes could not be retrieved! Exiting...', 2)
        exit()

    if adStatePostal == '':
        AddMsgAndPrint('State postal code could not be retrieved! Exiting...', 2)
        exit()

    # Change sourceCounty back to handle apostrophes
    sourceCounty = sourceCounty.replace("''", "'")
        
    # Transfer found values to variables to use for CLU download and project creation.
    adminState = stfip
    adminCounty = cofip
    postal = adStatePostal.lower()

    # Get the current year and month for use in project naming
    current = datetime.now()
    curyear = current.year
    curmonth = current.month
    theyear = str(curyear)
    themonth = str(curmonth)

    # Refine the tract number and month number to be padded with zeros to create uniform length of project name
    sourceTract = str(tractNumber)
    tractlength = len(sourceTract)
    if tractlength < 7:
        addzeros = 7 - tractlength
        tractname = '0'*addzeros + str(sourceTract)
    else:
        tractname = sourceTract

    monthlength = len(themonth)
    if monthlength < 2:
        finalmonth = '0' + themonth
    else:
        finalmonth = themonth

    # Build project folder path
    if projectType == 'New':
        projectFolder = path.join(workspacePath, f"{postal}{adminCounty}_t{tractname}_{theyear}_{finalmonth}")
    else:
        # Get project folder path from user input. Validation was done during script validations on the input
        if existingFolder != '':
            projectFolder = existingFolder
        else:
            AddMsgAndPrint('Project type was specified as Existing, but no existing project folder was selected. Exiting...', 2)
            exit()


    ### Set Additional Local Variables ###
    folderName = path.basename(projectFolder)
    projectName = folderName
    textFilePath = path.join(projectFolder, f"{folderName}_log.txt")
    basedataGDB_name = path.basename(projectFolder).replace(' ','_') + '_BaseData.gdb'
    basedataGDB_path = path.join(projectFolder, basedataGDB_name)
    userWorkspace = path.dirname(basedataGDB_path)
    basedataFD = path.join(basedataGDB_path, 'Layers')
    outputWS = basedataGDB_path
    templateCLU = path.join(path.dirname(argv[0]), 'SUPPORT.gdb', 'Site_CLU_template')
    cluTempName = 'CLU_Temp_' + projectName
    projectCLUTemp = path.join(basedataFD, cluTempName)
    cluName = 'Site_CLU'
    projectCLU = path.join(basedataFD, cluName)
    projectTract = path.join(basedataFD, 'Site_Tract')
    helFolder = path.join(projectFolder, 'HEL')
    helcGDB_name = f"{folderName}_HELC.gdb"
    helcGDB_path = path.join(helFolder, helcGDB_name)
    helcFD = path.join(helcGDB_path, 'HELC_Data')
    sitePrepareCLU_name = 'Site_Prepare_HELC'
    sitePrepareCLU = path.join(helcFD, sitePrepareCLU_name)
    scratchGDB = path.join(path.dirname(argv[0]), 'SCRATCH.gdb')
    jobid = uuid4()
    site_prepare_lyrx = LayerFile(path.join(path.join(path.dirname(argv[0]), 'layer_files'), 'Site_Prepare_HELC.lyrx')).listLayers()[0]
    site_clu_lyrx = LayerFile(path.join(path.join(path.dirname(argv[0]), 'layer_files'), 'Site_CLU.lyrx')).listLayers()[0]


    ### Create Project Folders and Contents ###
    AddMsgAndPrint('\nChecking project directories...')
    SetProgressorLabel('Checking project directories...')
    if not path.exists(workspacePath):
        try:
            SetProgressorLabel('Creating Determinations folder...')
            mkdir(workspacePath)
            AddMsgAndPrint('\nThe Determinations folder did not exist on the C: drive and has been created.')
        except:
            AddMsgAndPrint('\nThe Determinations folder cannot be created. Please check your permissions to the C: drive. Exiting...\n', 2)
            exit()

    if not path.exists(projectFolder):
        try:
            SetProgressorLabel('Creating project folder...')
            mkdir(projectFolder)
            AddMsgAndPrint('\nThe project folder has been created within C:\Determinations.')
        except:
            AddMsgAndPrint('\nThe project folder cannot be created. Please check your permissions to C:\Determinations. Exiting...\n', 2)
            exit()

    # Start logging to text file after project folder exists
    logBasicSettings(textFilePath, projectType, sourceState, sourceCounty, tractNumber, owFlag)

    SetProgressorLabel('Creating project contents...')
    if not path.exists(helFolder):
        try:
            SetProgressorLabel('Creating HEL folder...')
            mkdir(helFolder)
            AddMsgAndPrint(f"\nThe HEL folder has been created within {projectFolder}", textFilePath=textFilePath)
        except:
            AddMsgAndPrint('\nCould not access C:\Determinations. Check your permissions for C:\Determinations. Exiting...\n', 2, textFilePath)
            exit()

    if not Exists(basedataGDB_path):
        AddMsgAndPrint('\nCreating Base Data geodatabase...', textFilePath=textFilePath)
        SetProgressorLabel('Creating Base Data geodatabase...')
        CreateFileGDB(projectFolder, basedataGDB_name)

    if not Exists(basedataFD):
        AddMsgAndPrint('\nCreating Base Data feature dataset...', textFilePath=textFilePath)
        SetProgressorLabel('Creating Base Date feature dataset...')
        CreateFeatureDataset(basedataGDB_path, 'Layers', mapSR)

    if not Exists(helcGDB_path):
        AddMsgAndPrint('\nCreating HEL geodatabase...', textFilePath=textFilePath)
        SetProgressorLabel('Creating HEL geodatabase...')
        CreateFileGDB(helFolder, helcGDB_name)

    if not Exists(helcFD):
        AddMsgAndPrint('\nCreating HEL feature dataset...', textFilePath=textFilePath)
        SetProgressorLabel('Creating HEL feature dataset...')
        CreateFeatureDataset(helcGDB_path, 'HELC_Data', mapSR)


    ### Remove Existing CLU Layers From Map ###
    AddMsgAndPrint('\nRemoving CLU layer from project maps, if present...', textFilePath=textFilePath)
    SetProgressorLabel('Removing CLU layer from project maps, if present...')
    mapLayersToRemove = [cluName, sitePrepareCLU_name]
    try:
        for maps in aprx.listMaps():
            for lyr in maps.listLayers():
                if lyr.longName in mapLayersToRemove:
                    maps.removeLayer(lyr)
    except:
        pass


    ### If Overwrite, Delete Site_CLU, Site_Tract, Site_Prepare_CLU ###
    if owFlag == True:
        AddMsgAndPrint('\nOverwrite selected. Deleting existing CLU data...', textFilePath=textFilePath)
        SetProgressorLabel('Overwrite selected. Deleting existing CLU data...')
        delete_layers = [projectCLU, projectTract, sitePrepareCLU]
        for lyr in delete_layers:
            if Exists(lyr):
                Delete(lyr)
                AddMsgAndPrint(f"\nDeleted {lyr}...", textFilePath=textFilePath)


    ### Download the CLU ###
    if not Exists(projectCLU):
        AddMsgAndPrint('\nDownloading latest CLU data...', textFilePath=textFilePath)
        SetProgressorLabel('Downloading latest CLU data...')
        cluTempPath = start(adminState, adminCounty, tractNumber, mapSR, basedataGDB_path)

        # Convert feature class to the projectTempCLU layer in the project's feature dataset
        # This should work because the input CLU feature class coming from the download should have the same spatial reference as the target feature dataset
        FeatureClassToFeatureClass(cluTempPath, basedataFD, cluTempName)

        # Delete the temporary CLU download
        Delete(cluTempPath)

        # Add state name and county name fields to the projectTempCLU feature class
        AddField(projectCLUTemp, 'job_id', 'TEXT', '128')
        AddField(projectCLUTemp, 'admin_state_name', 'TEXT', '64')
        AddField(projectCLUTemp, 'admin_county_name', 'TEXT', '64')
        AddField(projectCLUTemp, 'state_name', 'TEXT', '64')
        AddField(projectCLUTemp, 'county_name', 'TEXT', '64')

        # Search the downloaded CLU for geographic state and county codes
        stateCo, countyCo = '', ''
        if sourceState == 'Alaska':
            field_names = ['state_ansi_code','county_ansi_code']
        else:
            field_names = ['state_code','county_code']
        with SearchCursor(projectCLUTemp, field_names) as cursor:
            for row in cursor:
                stateCo = row[0]
                countyCo = row[1]
                break
                                        
        # Search for names using FIPS codes.
        stName, coName = '', ''
        fields = ['STATEFP','COUNTYFP','NAME','STATE']
        expression = "{} = '{}'".format(AddFieldDelimiters(lut,'STATEFP'), stateCo) + " AND " + "{} = '{}'".format(AddFieldDelimiters(lut,'COUNTYFP'), countyCo)
        with SearchCursor(lut, fields, where_clause = expression) as cursor:
            for row in cursor:
                coName = row[2]
                stName = row[3]
                break

        if stName == '' or coName == '':
            AddMsgAndPrint('State and County Names for the site could not be retrieved! Exiting...\n', 2, textFilePath)
            exit()

        # Use Update Cursor to populate all rows of the downloaded CLU the same way for the new fields
        field_names = ['job_id','admin_state_name','admin_county_name','state_name','county_name']
        with UpdateCursor(projectCLUTemp, field_names) as cursor:
            for row in cursor:
                row[0] = jobid
                row[1] = sourceState
                row[2] = sourceCounty
                row[3] = stName
                row[4] = coName
                cursor.updateRow(row)

        # If the state is Alaska, update the admin_county FIPS and county_code FIPS from the county_ansi_code field
        if sourceState == 'Alaska':
            field_names = ['admin_county','county_code','county_ansi_code']
            with UpdateCursor(projectCLUTemp, field_names) as cursor:
                for row in cursor:
                    row[0] = row[2]
                    row[1] = row[2]
                    cursor.updateRow(row)
        
        # Create the projectCLU feature class and append projectCLUTemp to it. This is done as a cheat to using field mappings to re-order fields.
        AddMsgAndPrint('\nWriting Site CLU layer...', textFilePath=textFilePath)
        CreateFeatureclass(basedataFD, cluName, 'POLYGON', templateCLU)
        Append(projectCLUTemp, projectCLU, 'NO_TEST')
        Delete(projectCLUTemp)


    ### Create Yes/No Domain in Project's _HELC Geodatabase ###
    if not 'Yes No' in Describe(helcGDB_path).domains:
        yesnoTable = path.join(path.dirname(argv[0]), 'SUPPORT.gdb', 'domain_yesno_sodbust')
        TableToDomain(yesnoTable, 'Code', 'Description', helcGDB_path, 'Yes No Sodbust', 'Yes or no options', 'REPLACE')
        AlterDomain(helcGDB_path, 'Yes No Sodbust', '', '', 'DUPLICATE')


    ### Copy CLU to Project's _HELC Geodatabase ###
    if not Exists(sitePrepareCLU):
        FeatureClassToFeatureClass(projectCLU, helcFD, sitePrepareCLU_name)


    ### Add Sodbust Field to Site_Prepare_CLU and Assign Domain ###
    field_names = [f.name for f in ListFields(sitePrepareCLU)]
    if 'sodbust' not in field_names:
        AddField(sitePrepareCLU, 'sodbust', 'TEXT', field_length='3', field_alias='Sodbust', field_domain='Yes No Sodbust')


    ### Calculate Sodbust Vaue to 'No' ###
    CalculateField(sitePrepareCLU, 'sodbust', '"No"', 'PYTHON3')


    ### Create Tract Layer by Dissolving CLU Layer ###
    if not Exists(projectTract):
        AddMsgAndPrint('\nCreating Tract data...', textFilePath=textFilePath)
        SetProgressorLabel('Creating Tract data...')
        dis_fields = ['job_id','admin_state','admin_state_name','admin_county','admin_county_name','state_code','state_name','county_code','county_name','farm_number','tract_number']
        Dissolve(projectCLU, projectTract, dis_fields, '', 'MULTI_PART', '')


    ### Add CLU Layers to Map ###
    AddMsgAndPrint('\nAdding CLU layers to map...', textFilePath=textFilePath)
    SetProgressorLabel('Adding CLU layers to map...')
    lyr_name_list = [lyr.longName for lyr in map.listLayers()]
    addLyrxByConnectionProperties(map, lyr_name_list, site_clu_lyrx, basedataGDB_path, visible=False)
    addLyrxByConnectionProperties(map, lyr_name_list, site_prepare_lyrx, helcGDB_path)

    lyr_list = map.listLayers()
    for lyr in lyr_list:
        if lyr.longName == 'Common Land Unit Map Service':
            lyr.visible = False


    ### Zoom Map View to CLU ###
    clu_extent = Describe(projectCLU).extent
    clu_extent.XMin = clu_extent.XMin - 100
    clu_extent.XMax = clu_extent.XMax + 100
    clu_extent.YMin = clu_extent.YMin - 100
    clu_extent.YMax = clu_extent.YMax + 100
    map_view = aprx.activeView
    map_view.camera.setExtent(clu_extent)


    ### Compact Geodatabases ###
    try:
        AddMsgAndPrint('\nCompacting File Geodatabases...', textFilePath=textFilePath)
        SetProgressorLabel('Compacting File Geodatabases...')
        Compact(basedataGDB_path)
        Compact(helcGDB_path)
    except:
        pass

    AddMsgAndPrint('\nScript completed successfully', textFilePath=textFilePath)

except SystemExit:
    pass

except:
    try:
        AddMsgAndPrint(errorMsg('Create HEL Project'), 2, textFilePath)
    except FileNotFoundError:
        AddMsgAndPrint(errorMsg('Create HEL Project'), 2)
