from datetime import datetime
from os import mkdir, path
from sys import exit
from uuid import uuid4

import sys
sys.dont_write_bytecode=True
sys.path.append(path.dirname(sys.argv[0]))

from arcpy import AddError, AddFieldDelimiters, AddMessage, Describe, env, Exists, GetParameter, GetParameterAsText, \
    ListFeatureClasses, ListFields, SetParameterAsText, SetProgressorLabel, SpatialReference

from arcpy.conversion import FeatureClassToFeatureClass
from arcpy.da import SearchCursor, UpdateCursor

from arcpy.management import AddField, AlterDomain, Append, CalculateField, Compact, CreateFeatureclass, CreateFeatureDataset, \
    CreateFileGDB, Delete, Dissolve, TableToDomain

from arcpy.mp import ArcGISProject, LayerFile

from extract_CLU_by_Tract import getPortalTokenInfo, start
from hel_utils import AddMsgAndPrint, errorMsg


### Initial Tool Validation ###
try:
    aprx = ArcGISProject('CURRENT')
    map = aprx.listMaps('HEL Determination')[0]
except Exception:
    AddMsgAndPrint('This tool must be run from an ArcGIS Pro project that was developed from the template distributed with this toolbox. Exiting...', 2)
    exit()

nrcsPortal = 'https://gis.sc.egov.usda.gov/portal/'
portalToken = getPortalTokenInfo(nrcsPortal)
if not portalToken:
    AddError('Could not generate Portal token. Please login to GeoPortal. Exiting...')
    exit()

lut = path.join(path.dirname(sys.argv[0]), 'SUPPORT.gdb', 'lut_census_fips')
if not Exists(lut):
    AddError('Could not find state and county lookup table. Exiting...')
    exit()


### Validate Spatial Reference ###
mapSR = map.spatialReference
if mapSR.type != 'Projected':
    AddError('\nThe Determinations map is not set to a Projected coordinate system.')
    AddError('\nPlease assign a WGS 1984 UTM coordinate system to the Determinations map that is appropriate for your site.')
    AddError('\nThese systems are found in the Determinations Map Properties under: Coordinate Systems -> Projected Coordinate System -> UTM -> WGS 1984.')
    AddError('\nAfter applying a coordinate system, save your template and try this tool again.')
    AddError('\nExiting...')
    exit()

if 'WGS' not in mapSR.name or '1984' not in mapSR.name or 'UTM' not in mapSR.name:
    AddError('\nThe Determinations map is not using a UTM coordinate system tied to WGS 1984.')
    AddError('\nPlease assign a WGS 1984 UTM coordinate system to the Determinations map that is appropriate for your site.')
    AddError('\nThese systems are found in the Determinations Map Properties under: Coordinate Systems -> Projected Coordinate System -> UTM -> WGS 1984.')
    AddError('\nAfter applying a coordinate system, save your template and try this tool again.')
    AddError('\nExiting...')
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
        AddError('State and County FIPS codes could not be retrieved! Exiting...')
        exit()

    if adStatePostal == '':
        AddError('State postal code could not be retrieved! Exiting...')
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
            AddError('Project type was specified as Existing, but no existing project folder was selected. Exiting...')
            exit()


    ### Set Additional Local Variables ###
    folderName = path.basename(projectFolder)
    projectName = folderName
    basedataGDB_name = path.basename(projectFolder).replace(' ','_') + '_BaseData.gdb'
    basedataGDB_path = path.join(projectFolder, basedataGDB_name)
    userWorkspace = path.dirname(basedataGDB_path)
    basedataFD = path.join(basedataGDB_path, 'Layers')
    outputWS = basedataGDB_path
    templateCLU = path.join(path.dirname(sys.argv[0]), 'SUPPORT.gdb', 'Site_CLU_template')
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
    scratchGDB = path.join(path.dirname(sys.argv[0]), 'SCRATCH.gdb')
    jobid = uuid4()
    site_prepare_lyrx = LayerFile(path.join(path.join(path.dirname(sys.argv[0]), 'layer_files'), 'Site_Prepare_HELC.lyrx')).listLayers()[0]


    ### Create Project Folders and Contents ###
    AddMessage('\nChecking project directories...')
    SetProgressorLabel('Checking project directories...')
    if not path.exists(workspacePath):
        try:
            SetProgressorLabel('Creating Determinations folder...')
            mkdir(workspacePath)
            AddMessage('\nThe Determinations folder did not exist on the C: drive and has been created.')
        except:
            AddError('\nThe Determinations folder cannot be created. Please check your permissions to the C: drive. Exiting...\n')
            exit()

    if not path.exists(projectFolder):
        try:
            SetProgressorLabel('Creating project folder...')
            mkdir(projectFolder)
            AddMessage('\nThe project folder has been created within C:\Determinations.')
        except:
            AddError('\nThe project folder cannot be created. Please check your permissions to C:\Determinations. Exiting...\n')
            exit()

    SetProgressorLabel('Creating project contents...')
    if not path.exists(helFolder):
        try:
            SetProgressorLabel('Creating HEL folder...')
            mkdir(helFolder)
            AddMsgAndPrint(f"\nThe HEL folder has been created within {projectFolder}")
        except:
            AddMsgAndPrint('\nCould not access C:\Determinations. Check your permissions for C:\Determinations. Exiting...\n', 2)
            exit()

    if not Exists(basedataGDB_path):
        AddMsgAndPrint('\nCreating Base Data geodatabase...')
        SetProgressorLabel('Creating Base Data geodatabase...')
        CreateFileGDB(projectFolder, basedataGDB_name)

    if not Exists(basedataFD):
        AddMsgAndPrint('\nCreating Base Data feature dataset...')
        SetProgressorLabel('Creating Base Date feature dataset...')
        CreateFeatureDataset(basedataGDB_path, 'Layers', mapSR)

    if not Exists(helcGDB_path):
        AddMsgAndPrint('\nCreating HEL geodatabase...')
        SetProgressorLabel('Creating HEL geodatabase...')
        CreateFileGDB(helFolder, helcGDB_name)

    if not Exists(helcFD):
        AddMsgAndPrint('\nCreating HEL feature dataset...')
        SetProgressorLabel('Creating HEL feature dataset...')
        CreateFeatureDataset(helcGDB_path, 'HELC_Data', mapSR)


    ### Remove Existing CLU Layers From Map ###
    AddMsgAndPrint('\nRemoving CLU layer from project maps, if present...\n')
    SetProgressorLabel('Removing CLU layer from project maps, if present...')
    mapLayersToRemove = [cluName, sitePrepareCLU_name]
    try:
        for maps in aprx.listMaps():
            for lyr in maps.listLayers():
                if lyr.longName in mapLayersToRemove:
                    maps.removeLayer(lyr)
    except:
        pass


    ### If Overwrite, Delete Everything and Start Over ###
    if owFlag == True:
        AddMsgAndPrint("\nOverwrite selected. Deleting existing project data...",0)
        SetProgressorLabel("Overwrite selected. Deleting existing project data...")
        if Exists(basedataFD):
            ws = env.workspace
            env.workspace = basedataGDB_path
            fcs = ListFeatureClasses(feature_dataset='Layers')
            for fc in fcs:
                try:
                    fc_path = path.join(basedataFD, fc)
                    Delete(fc_path)
                except:
                    pass
            Delete(basedataFD)
            CreateFeatureDataset(basedataGDB_path, 'Layers', mapSR)


    ### Download the CLU ###
    if not Exists(projectCLU):
        AddMsgAndPrint('\nDownloading latest CLU data...')
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
            AddError('State and County Names for the site could not be retrieved! Exiting...\n')
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
        AddMsgAndPrint('\nWriting Site CLU layer...')
        CreateFeatureclass(basedataFD, cluName, 'POLYGON', templateCLU)
        Append(projectCLUTemp, projectCLU, 'NO_TEST')
        Delete(projectCLUTemp)


    ### Create Yes/No Domain in Project's _HELC Geodatabase ###
    if not 'Yes No' in Describe(helcGDB_path).domains:
        yesnoTable = path.join(path.dirname(sys.argv[0]), 'SUPPORT.gdb', 'domain_yesno')
        TableToDomain(yesnoTable, 'Code', 'Description', helcGDB_path, 'Yes No', 'Yes or no options', 'REPLACE')
        AlterDomain(helcGDB_path, 'Yes No', '', '', 'DUPLICATE')
    
    
    ### Copy CLU to Project's _HELC Geodatabase ###
    if not Exists(sitePrepareCLU):
        FeatureClassToFeatureClass(projectCLU, helcFD, sitePrepareCLU_name)


    ### Add Sodbust Field to Site_Prepare_CLU and Assign Domain ###
    field_names = [f.name for f in ListFields(sitePrepareCLU)]
    if 'sodbust' not in field_names:
        AddField(sitePrepareCLU, 'sodbust', 'TEXT', field_length='3', field_alias='Sodbust', field_domain='Yes No')


    ### Calculate Sodbust Vaue to 'No' ###
    CalculateField(sitePrepareCLU, 'sodbust', '"No"', 'PYTHON3')

    
    ### Create Tract Layer by Dissolving CLU Layer ###
    if not Exists(projectTract):
        AddMsgAndPrint('\nCreating Tract data...')
        SetProgressorLabel('Creating Tract data...')
        dis_fields = ['job_id','admin_state','admin_state_name','admin_county','admin_county_name','state_code','state_name','county_code','county_name','farm_number','tract_number']
        Dissolve(projectCLU, projectTract, dis_fields, '', 'MULTI_PART', '')


    ### Add CLU Layers to Map ###
    AddMsgAndPrint('\nAdding CLU layers to map...')
    SetProgressorLabel('Adding CLU layers to map...')
    SetParameterAsText(6, projectCLU)

    lyr_list = map.listLayers()
    lyr_name_list = []
    for lyr in lyr_list:
        lyr_name_list.append(lyr.longName)

    if sitePrepareCLU_name not in lyr_name_list:
        site_prepare_lyrx_cp = site_prepare_lyrx.connectionProperties
        site_prepare_lyrx_cp['connection_info']['database'] = helcGDB_path
        site_prepare_lyrx_cp['dataset'] = sitePrepareCLU_name
        site_prepare_lyrx.updateConnectionProperties(site_prepare_lyrx.connectionProperties, site_prepare_lyrx_cp)
        map.addLayer(site_prepare_lyrx)


    ### Zoom Map View to CLU ###
    clu_extent = Describe(projectCLU).extent
    map_view = aprx.activeView
    map_view.camera.setExtent(clu_extent)


    ### Compact Geodatabases ###
    try:
        AddMsgAndPrint('\nCompacting File Geodatabases...')
        SetProgressorLabel('Compacting File Geodatabases...')
        Compact(basedataGDB_path)
        Compact(helcGDB_path)
        AddMsgAndPrint('\tSuccessful')
    except:
        pass

except SystemExit:
    pass

except KeyboardInterrupt:
    AddMsgAndPrint('Interruption requested. Exiting...')

except:
    errorMsg()
