from getpass import getuser
from os import path
from sys import argv, exit
from time import ctime

from arcpy import Describe, env, Exists, GetParameterAsText, SetProgressorLabel
from arcpy.conversion import TableToTable
from arcpy.da import InsertCursor, SearchCursor
from arcpy.management import Compact, CreateFeatureDataset, CreateFileGDB, CreateTable, Delete, DeleteRows, GetCount
from arcpy.mp import ArcGISProject

from hel_utils import AddMsgAndPrint, errorMsg


textFilePath = ''
def logBasicSettings(textFilePath, sourceCLU, client, delineator, digitizer, requestType, requestDate):
    with open(textFilePath, 'a+') as f:
        f.write('\n######################################################################\n')
        f.write('Executing Tool: Enter Project Info\n')
        f.write(f"User Name: {getuser()}\n")
        f.write(f"Date Executed: {ctime()}\n")
        f.write('User Parameters:\n')
        f.write(f"\tSelected CLU Layer: {sourceCLU}\n")
        f.write(f"\tClient Name: {client}\n")
        f.write(f"\tDelineator Name: {delineator}\n")
        f.write(f"\tDigitizer Name: {digitizer}\n")
        f.write(f"\tRequest Type: {requestType}\n")
        f.write(f"\tRequest Date: {requestDate}\n")


# Set arcpy Environment Settings
env.overwriteOutput = True

### Initial Tool Validation ###
try:
    aprx = ArcGISProject('CURRENT')
    m = aprx.listMaps('HEL Determination')[0]
except:
    AddMsgAndPrint('\nThis tool must be run from an active ArcGIS Pro project that was developed from the template distributed with this toolbox. Exiting...\n', 2)
    exit()

# Main procedures
try:
    SetProgressorLabel('Reading inputs...')
    sourceCLU = GetParameterAsText(0)         # User selected CLU file from the project
    client = GetParameterAsText(1)            # Client Name
    delineator = GetParameterAsText(2)        # The person who conducts the technical determination
    digitizer = GetParameterAsText(3)         # The person who digitizes the determination (may or may not match the delineator)
    requestType = GetParameterAsText(4)       # Request Type (AD-1026, FSA-569, or NRCS-CPA-38)
    requestDate = GetParameterAsText(5)       # Determination request date per request form signature date
    clientStreet = GetParameterAsText(6)
    clientStreet2 = GetParameterAsText(7)
    clientCity = GetParameterAsText(8)
    clientState = GetParameterAsText(9)
    clientZip = GetParameterAsText(10)

    # Get the basedataGDB_path from the input CLU layer. If else retained in case of other project path oddities.
    sourceCLU_path = Describe(sourceCLU).CatalogPath
    if sourceCLU_path.find('.gdb') > 0 and sourceCLU_path.find('Determinations') > 0 and sourceCLU_path.find('Site_CLU') > 0:
        basedataGDB_path = sourceCLU_path[:sourceCLU_path.find('.gdb')+4]
    else:
        AddMsgAndPrint('\nSelected CLU layer is not from a Determinations project folder. Exiting...', 2)
        exit()

    # Define Variables
    SetProgressorLabel('Setting variables...')
    basedataGDB_name = path.basename(basedataGDB_path)
    userWorkspace = path.dirname(basedataGDB_path)
    projectName = path.basename(userWorkspace).replace(' ', '_')
    cluName = 'Site_CLU'
    projectCLU = path.join(basedataGDB_path, 'Layers', cluName)
    helDir = path.join(userWorkspace, 'HEL')
    
    helGDB_name = f"{path.basename(userWorkspace).replace(' ', '_')}_HELC.gdb"
    helGDB_path = path.join(helDir, helGDB_name)
    helFD = path.join(helGDB_path, 'HELC_Data')
    
    templateTable = path.join(path.dirname(argv[0]), path.join('SUPPORT.gdb', 'table_admin'))
    tableName = f"Table_{projectName}"

    # Permanent Datasets
    projectTable = path.join(basedataGDB_path, tableName)
    helDetTable = path.join(helGDB_path, 'Admin_Table')

    # Start logging to text file
    textFilePath = path.join(userWorkspace, f"{projectName}_log.txt")
    logBasicSettings(textFilePath, sourceCLU, client, delineator, digitizer, requestType, requestDate)

    # Get Job ID from input CLU
    SetProgressorLabel('Recording project Job ID...')
    fields = ['job_id']
    with SearchCursor(projectCLU, fields) as cursor:
        for row in cursor:
            jobid = row[0]
            break

    # If the job's admin table exists, clear all rows, else create the table
    if Exists(projectTable):
        SetProgressorLabel('Located project admin table table...')
        recordsCount = int(GetCount(projectTable)[0])
        if recordsCount > 0:
            DeleteRows(projectTable)
            AddMsgAndPrint('\nCleared existing row from project admin table...', textFilePath=textFilePath)
    else:
        SetProgressorLabel('Creating administrative table...')
        CreateTable(basedataGDB_path, tableName, templateTable)
        AddMsgAndPrint('\nCreated administrative table...', textFilePath=textFilePath)

    # Use a search cursor to get the tract location info from the CLU layer
    AddMsgAndPrint('\nImporting tract data from the CLU...', textFilePath=textFilePath)
    SetProgressorLabel('Importing tract data from the CLU...')
    field_names = ['admin_state','admin_state_name','admin_county','admin_county_name',
                   'state_code','state_name','county_code','county_name','farm_number','tract_number']
    with SearchCursor(sourceCLU, field_names) as cursor:
        for row in cursor:
            adminState = row[0]
            adminStateName = row[1]
            adminCounty = row[2]
            adminCountyName = row[3]
            stateCode = row[4]
            stateName = row[5]
            countyCode = row[6]
            countyName = row[7]
            farmNumber = row[8]
            tractNumber = row[9]
            break

    # Use an insert cursor to add record to the admin table
    AddMsgAndPrint('\nUpdating the administrative table...', textFilePath=textFilePath)
    SetProgressorLabel('Updating the administrative table...')
    field_names = ['admin_state','admin_state_name','admin_county','admin_county_name','state_code','state_name',
                   'county_code','county_name','farm_number','tract_number','client','deter_staff',
                   'dig_staff','request_date','request_type','street','street_2','city','state','zip','job_id']
    row = (adminState, adminStateName, adminCounty, adminCountyName, stateCode, stateName, countyCode,
           countyName, farmNumber, tractNumber, client, delineator, digitizer, requestDate, requestType,
           clientStreet if clientStreet else None, clientStreet2 if clientStreet2 else None, clientCity if clientCity else None,
           clientState if clientState else None, clientZip if clientZip else None, jobid)
    with InsertCursor(projectTable, field_names) as cursor:
        cursor.insertRow(row)

    # Create a text file output version of the admin table for consumption by external data collection forms
    # Set a file name and export to the user workspace folder for the project
    AddMsgAndPrint('\nExporting administrative text file...', textFilePath=textFilePath)
    SetProgressorLabel('Exporting administrative text file...')
    textTable = f"Admin_Info_{projectName}.txt"
    if Exists(textTable):
        Delete(textTable)
    TableToTable(projectTable, userWorkspace, textTable)

    # If project HEL geodatabase and feature dataset do not exist, create them.
    # Get the spatial reference from the Define AOI feature class and use it, if needed
    AddMsgAndPrint('\nChecking project integrity...', textFilePath=textFilePath)
    SetProgressorLabel('Checking project integrity...')
    desc = Describe(sourceCLU)
    sr = desc.SpatialReference
    
    if not Exists(helGDB_path):
        AddMsgAndPrint('\nCreating HEL geodatabase...', textFilePath=textFilePath)
        SetProgressorLabel('Creating HEL geodatabase...')
        CreateFileGDB(helDir, helGDB_name)

    if not Exists(helFD):
        AddMsgAndPrint('\nCreating HEL feature dataset...', textFilePath=textFilePath)
        SetProgressorLabel('Creating HEL feature dataset...')
        CreateFeatureDataset(helGDB_path, 'HELC_Data', sr)

    # Copy the administrative table into the wetlands database for use with the attribute rules during digitizing
    AddMsgAndPrint('\nUpdating administrative table in GDB...', textFilePath=textFilePath)
    if Exists(helDetTable):
        Delete(helDetTable)
    TableToTable(projectTable, helGDB_path, 'Admin_Table')

    # Adjust layer visibility in maps, turn off CLU layer
    AddMsgAndPrint('\nUpdating layer visibility to off...', textFilePath=textFilePath)
    off_names = [cluName]
    for maps in aprx.listMaps():
        for lyr in maps.listLayers():
            for name in off_names:
                if name in lyr.longName:
                    lyr.visible = False

    # Compact file geodatabase
    try:
        AddMsgAndPrint('\nCompacting File Geodatabase...', textFilePath=textFilePath)
        SetProgressorLabel('Compacting File Geodatabase...')
        Compact(basedataGDB_path)
    except:
        pass

    AddMsgAndPrint('\nScript completed successfully', textFilePath=textFilePath)

except SystemExit:
    pass

except:
    if textFilePath:
        AddMsgAndPrint(errorMsg('Enter Project Info'), 2, textFilePath)
    else:
        AddMsgAndPrint(errorMsg('Enter Project Info'), 2)
