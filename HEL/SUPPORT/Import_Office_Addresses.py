from os import path
from sys import argv

from arcpy import AddError, AddMessage, env, Exists, ListFields, SetProgressorLabel
from arcpy.conversion import TableToTable
from arcpy.da import UpdateCursor
from arcpy.management import Delete


AddMessage('Setting variables...')
SetProgressorLabel('Setting variables...')

supportGDB = path.join(path.dirname(argv[0]), 'SUPPORT.gdb')
templates_folder = path.join(path.dirname(argv[0]), 'Templates')
nrcs_address_csv = path.join(templates_folder, 'NRCS_Address.csv')
nrcs_temp_path = path.join(supportGDB, 'nrcs_temp')
fsa_address_csv = path.join(templates_folder, 'FSA_Address.csv')
fsa_temp_path = path.join(supportGDB, 'fsa_temp')
nad_address_csv = path.join(templates_folder, 'NAD_Address.csv')
nad_temp_path = path.join(supportGDB, 'nad_temp')

# Set overwrite flag
env.workspace = supportGDB
env.overwriteOutput = True

temp_tables = [nrcs_temp_path, fsa_temp_path, nad_temp_path]
for item in temp_tables:
    if Exists(item):
        Delete(item)

if Exists(nrcs_address_csv) and Exists(fsa_address_csv) and Exists(nad_address_csv):
    try:
        AddMessage('Importing NRCS Office Table...')
        SetProgressorLabel('Importing NRCS Office Table...')
        TableToTable(nrcs_address_csv, supportGDB, 'nrcs_temp')
        regen = False
        nrcs_fields = ListFields(nrcs_temp_path)
        for field in nrcs_fields:
            if field.type == 'Integer':
                regen = True
        if regen == True:
            # Use Field Mappings to control column format in output (all output fields need to be text/string in output)
            # CSV defaults anyting that looks like a number to a number, which might be the NRCSZIP field as an Integer from the NRCS_Address.csv table if no entered zips have dashes.
            TableToTable(nrcs_address_csv, supportGDB, 'nrcs_addresses', '',
                        r'NRCSOffice "NRCSOffice" true true false 8000 Text 0 0,First,#,' + nrcs_address_csv + ',NRCSOffice,0,8000;' +
                        r'NRCSAddress "NRCSAddress" true true false 8000 Text 0 0,First,#,' + nrcs_address_csv + ',NRCSAddress,0,8000;' +
                        r'NRCSCITY "NRCSCITY" true true false 8000 Text 0 0,First,#,' + nrcs_address_csv + ',NRCSCITY,0,8000;' +
                        r'NRCSSTATE "NRCSSTATE" true true false 8000 Text 0 0,First,#,' + nrcs_address_csv + ',NRCSSTATE,0,8000;' +
                        r'NRCSZIP "NRCSZIP" true true false 8000 Text 0 0,First,#,' + nrcs_address_csv + ',NRCSZIP,-1,-1;' +
                        r'NRCSPHONE "NRCSPHONE" true true false 8000 Text 0 0,First,#,' + nrcs_address_csv + ',NRCSPHONE,0,8000;' +
                        r'NRCSFAX "NRCSFAX" true true false 8000 Text 0 0,First,#,' + nrcs_address_csv + ',NRCSFAX,0,8000;' +
                        r'NRCSCounty "NRCSCounty" true true false 8000 Text 0 0,First,#,' + nrcs_address_csv + ',NRCSCounty,0,8000',
                        '')
        else:
            TableToTable(nrcs_address_csv, supportGDB, 'nrcs_addresses')
        Delete(nrcs_temp_path)
        
        AddMessage('Importing FSA Office Table...')
        SetProgressorLabel('Importing FSA Office Table...')
        TableToTable(fsa_address_csv, supportGDB, 'fsa_temp')
        regen = False
        fsa_fields = ListFields(fsa_temp_path)
        for field in fsa_fields:
            if field.type == 'Integer':
                regen = True
        if regen == True:
            # Use Field Mappings to control column format in output (all output fields need to be text/string in output)
            # CSV defaults anyting that looks like a number to a number, which might be the FSAZIP field as an Integer from the FSA_Address.csv table if no entered zips have dashes.
            TableToTable(fsa_address_csv, supportGDB, 'fsa_addresses', '',
                        r'FSAOffice "FSAOffice" true true false 8000 Text 0 0,First,#,' + fsa_address_csv + ',FSAOffice,0,8000;' +
                        r'FSAAddress "FSAAddress" true true false 8000 Text 0 0,First,#,' + fsa_address_csv + ',FSAAddress,0,8000;' +
                        r'FSACITY "FSACITY" true true false 8000 Text 0 0,First,#,' + fsa_address_csv + ',FSACITY,0,8000;' +
                        r'FSASTATE "FSASTATE" true true false 8000 Text 0 0,First,#,' + fsa_address_csv + ',FSASTATE,0,8000;' +
                        r'FSAZIP "FSAZIP" true true false 8000 Text 0 0,First,#,' + fsa_address_csv + ',FSAZIP,-1,-1;' +
                        r'FSAPHONE "FSAPHONE" true true false 8000 Text 0 0,First,#,' + fsa_address_csv + ',FSAPHONE,0,8000;' +
                        r'FSAFAX "FSAFAX" true true false 8000 Text 0 0,First,#,' + fsa_address_csv + ',FSAFAX,0,8000;' +
                        r'FSACounty "FSACounty" true true false 8000 Text 0 0,First,#,' + fsa_address_csv + ',FSACounty,0,8000',
                        '')
        else:
            TableToTable(fsa_address_csv, supportGDB, 'fsa_addresses')
        Delete(fsa_temp_path)

        AddMessage('Importing NAD Office Table...')
        SetProgressorLabel('Importing NAD Office Table...')
        TableToTable(nad_address_csv, supportGDB, 'nad_temp')
        regen = False
        nad_fields = ListFields(nad_temp_path)
        for field in nad_fields:
            if field.type == 'Integer':
                regen = True
        if regen == True:
            # Use Field Mappings to control column format in output (all output fields need to be text/string in output)
            # CSV defaults anyting that looks like a number to a number, which will be the STATECD field as an Integer from the NAD_Address.csv table.
            TableToTable(nad_address_csv, supportGDB, 'nad_addresses', '',
                        r'STATECD "STATECD" true true false 8000 Text 0 0,First,#,' + nad_address_csv + ',STATECD,-1,-1;' +
                        r'STATE "STATE" true true false 8000 Text 0 0,First,#,' + nad_address_csv + ',STATE,0,8000;' +
                        r'NADADDRESS "NADADDRESS" true true false 8000 Text 0 0,First,#,' + nad_address_csv + ',NADADDRESS,0,8000;' +
                        r'NADCITY "NADCITY" true true false 8000 Text 0 0,First,#,' + nad_address_csv + ',NADCITY,0,8000;' +
                        r'NADSTATE "NADSTATE" true true false 8000 Text 0 0,First,#,' + nad_address_csv + ',NADSTATE,0,8000;' +
                        r'NADZIP "NADZIP" true true false 8000 Text 0 0,First,#,' + nad_address_csv + ',NADZIP,-1,-1;' +
                        r'TOLLFREE "TOLLFREE" true true false 8000 Text 0 0,First,#,' + nad_address_csv + ',TOLLFREE,0,8000;' +
                        r'PHONE "PHONE" true true false 8000 Text 0 0,First,#,' + nad_address_csv + ',PHONE,0,8000;' +
                        r'TTY "TTY" true true false 8000 Text 0 0,First,#,' + nad_address_csv + ',TTY,0,8000;' +
                        r'FAX "FAX" true true false 8000 Text 0 0,First,#,' + nad_address_csv + ',FAX,0,8000',
                        '')
        else:
            TableToTable(nad_address_csv, supportGDB, 'nad_addresses')
        Delete(nad_temp_path)
        # Add leading 0 to any state codes of length 1 (e.g. 1 becomes 01)
        with UpdateCursor('nad_addresses', ['STATECD']) as cursor:
            for row in cursor:
                if len(row[0]) == 1:
                    row[0] = f"0{row[0]}"
                    cursor.updateRow(row)
    except:
        AddError('Something went wrong in the import process. Exiting...')
        exit()
else:
    AddError('Could not find expected NRCS, FSA, and/or NAD Address CSVs in install folders. Exiting...')
    exit()

AddMessage('Address table imports were successful! Exiting...')
