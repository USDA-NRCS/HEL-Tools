from datetime import date
from getpass import getuser
from os import path as os_path, startfile
from sys import exit, path as sys_path
from time import ctime

from arcpy import AddFieldDelimiters, Describe, Exists, GetParameter, GetParameterAsText, SetProgressorLabel
from arcpy.analysis import Statistics
from arcpy.da import SearchCursor
from arcpy.management import AddField, CalculateField, Delete, GetCount, MakeFeatureLayer, Sort
from arcpy.mp import ArcGISProject

from hel_utils import AddMsgAndPrint, errorMsg

# Hack to allow imports of local libraries from python_packages folder
base_dir = os_path.abspath(os_path.dirname(__file__)) #\SUPPORT
sys_path.append(os_path.join(base_dir, 'python_packages'))

from python_packages.docx.api import Document
from python_packages.docxcompose.composer import Composer
from python_packages.docxtpl import DocxTemplate


def logBasicSettings(textFilePath, hel_map_units, where_completed, nrcs_office, fsa_county, fsa_office, consolidate_by_clu):
    with open(textFilePath, 'a+') as f:
        f.write('\n######################################################################\n')
        f.write('Executing Tool: Create Form, Letter, Report\n')
        f.write(f"User Name: {getuser()}\n")
        f.write(f"Date Executed: {ctime()}\n")
        f.write('User Parameters:\n')
        f.write(f"\tHEL Map Units on Tract: {hel_map_units}\n")
        f.write(f"\tWhere Determination Completed: {where_completed}\n")
        f.write(f"\tNRCS Office: {nrcs_office}\n")
        f.write(f"\tFSA County Committee County: {fsa_county}\n")
        f.write(f"\tFSA Office: {fsa_office}\n")
        f.write(f"\tConsolidate by CLU: {consolidate_by_clu}\n")


def add_blank_rows(table_data, max_rows):
    '''Given a list of table data, adds empty rows (dict) up to a specified max'''
    if len(table_data) < max_rows:
        for index in range(max_rows - len(table_data)):
            table_data.append({})
    return table_data


### Initial Tool Validation ###
try:
    aprx = ArcGISProject('CURRENT')
    aprx.listMaps('HEL Determination')[0]
except:
    AddMsgAndPrint('This tool must be run from an ArcGIS Pro project that was developed from the template distributed with this toolbox. Exiting...', 2)
    exit()

AddMsgAndPrint('\nCollecting inputs...')
SetProgressorLabel('Collecting inputs...')

### Input Parameters ###
field_det_lyr = GetParameterAsText(0)
hel_map_units = GetParameterAsText(1)
where_completed = GetParameterAsText(2)
nrcs_office = GetParameterAsText(3)
fsa_county = GetParameterAsText(4)
fsa_office = GetParameterAsText(5)
consolidate_by_clu = GetParameter(6)

AddMsgAndPrint('\nAssigning local variables...')
SetProgressorLabel('Assigning local variables...')

### Paths to Word Templates ###
templates_dir = os_path.join(base_dir, 'Templates')
customer_letter_template_path = os_path.join(templates_dir, 'HELC_Letter_Template.docx')
cpa_026_helc_template_path = os_path.join(templates_dir, 'CPA_026_HELC_Template.docx')
client_report_template_path = os_path.join(templates_dir, 'Client_Report_Template.docx')
planner_summary_template_path = os_path.join(templates_dir, 'Planner_Summary_Template.docx')

### Paths to SUPPORT GDB ###
support_gdb = os_path.join(base_dir, 'SUPPORT.gdb')
nrcs_addresses_table = os_path.join(support_gdb, 'nrcs_addresses')
fsa_addresses_table = os_path.join(support_gdb, 'fsa_addresses')
nad_addresses_table = os_path.join(support_gdb, 'nad_addresses')

### Paths to Site GDB ###
field_det_lyr_path = Describe(field_det_lyr).CatalogPath
site_gdb = field_det_lyr_path[:field_det_lyr_path.find('.gdb')+4]
admin_table = os_path.join(site_gdb, 'Admin_Table')
final_hel_summary_lyr_path = os_path.join(site_gdb, 'Final_HEL_Summary')
final_hel_stats_table_path = os_path.join(site_gdb, 'Final_HEL_Summary_Statistics')
field_det_sorted_lyr_path = os_path.join(site_gdb, 'HELC_Data', 'Field_Determination_Sorted')

### Paths to HEL Project Folder for Outputs ###
hel_dir = os_path.dirname(site_gdb)
customer_letter_output = os_path.join(hel_dir, 'HELC_Letter.docx')
cpa_026_helc_output = os_path.join(hel_dir, 'NRCS-CPA-026-HELC-Form.docx')
client_report_output = os_path.join(hel_dir, 'Client_Report.docx')
planner_summary_output = os_path.join(hel_dir, 'Planner_Summary.docx')

### Path to Log File ###
project_dir = os_path.dirname(hel_dir)
folder_name = os_path.basename(project_dir)
textFilePath = os_path.join(project_dir, f"{folder_name}_log.txt")
logBasicSettings(textFilePath, hel_map_units, where_completed, nrcs_office, fsa_county, fsa_office, consolidate_by_clu)

### Ensure Word Doc Templates Exist ###
#NOTE: Front end validation checks the existance of most geodatabase tables
for path in [customer_letter_template_path, cpa_026_helc_template_path, client_report_template_path, planner_summary_template_path]:
    if not os_path.exists(path):
        AddMsgAndPrint(f"\nFailed to locate required Word template: {path}. Exiting...", 2, textFilePath)
        exit()


### Read and assign values from Admin Table ###
SetProgressorLabel('Reading table data...')
AddMsgAndPrint('\nReading table data...', textFilePath=textFilePath)
if int(GetCount(admin_table)[0]) != 1:
    AddMsgAndPrint('\nAdmin Table has more than one entry. Exiting...', 2, textFilePath)
    exit()

try:
    admin_data = {}
    fields = ['admin_state', 'admin_state_name', 'admin_county', 'admin_county_name', 'state_code', 'state_name', 'county_code', 'county_name', 
        'farm_number', 'tract_number', 'client', 'deter_staff', 'dig_staff', 'request_date', 'request_type', 'comments', 'street', 'street_2', 
        'city', 'state', 'zip']
    with SearchCursor(admin_table, fields) as cursor:
        row = cursor.next()
        admin_data['admin_state'] = row[0] if row[0] else ''
        admin_data['admin_state_name'] = row[1] if row[1] else ''
        admin_data['admin_county'] = row[2] if row[2] else ''
        admin_data['admin_county_name'] = row[3] if row[3] else ''
        admin_data['state_code'] = row[4] if row[4] else ''
        admin_data['state_name'] = row[5] if row[5] else ''
        admin_data['county_code'] = row[6] if row[6] else ''
        admin_data['county_name'] = row[7] if row[7] else ''
        admin_data['farm_number'] = row[8] if row[8] else ''
        admin_data['tract_number'] = row[9] if row[9] else ''
        admin_data['client'] = row[10] if row[10] else ''
        admin_data['deter_staff'] = row[11] if row[11] else ''
        admin_data['dig_staff'] = row[12] if row[12] else ''
        admin_data['request_date'] = row[13].strftime('%m/%d/%Y') if row[13] else ''
        admin_data['request_type'] = row[14] if row[14] else ''
        admin_data['comments'] = row[15] if row[15] else ''
        admin_data['street'] = f'{row[16]}, {row[17]}' if row[17] else row[16]
        admin_data['city'] = row[18] if row[18] else ''
        admin_data['state'] = row[19] if row[19] else ''
        admin_data['zip'] = row[20] if row[20] else ''
except:
    AddMsgAndPrint('\nFailed while retrieving Admin Table data. Exiting...', 2, textFilePath)
    AddMsgAndPrint(errorMsg('Create Form, Letter, Report'), 2, textFilePath)
    exit()


### Read and assign values from NRCS Addresses Table - select row by NRCS Office input ###
# Handle apostrophe in office name for SQL statement
if "'" in nrcs_office:
    nrcs_office = nrcs_office.replace("'", "''")
try:
    nrcs_address = {}
    fields = ['NRCSOffice', 'NRCSAddress', 'NRCSCITY', 'NRCSSTATE', 'NRCSZIP', 'NRCSPHONE', 'NRCSFAX']
    where_clause = """{0}='{1}'""".format(AddFieldDelimiters(support_gdb, 'NRCSOffice'), nrcs_office)
    with SearchCursor(nrcs_addresses_table, fields, where_clause) as cursor:
        row = cursor.next()
        nrcs_address['office'] = row[0] if row[0] else ''
        nrcs_address['street'] = row[1] if row[1] else ''
        nrcs_address['city'] = row[2] if row[2] else ''
        nrcs_address['state'] = row[3] if row[3] else ''
        nrcs_address['zip'] = row[4] if row[4] else ''
        nrcs_address['phone'] = row[5] if row[5] else ''
        nrcs_address['fax'] = row[6] if row[6] else ''
except:
    AddMsgAndPrint('\nFailed while retrieving NRCS Address Table data.\nYou may need to run tool F.Import Office Addresses and then try this tool again. Exiting...', 2, textFilePath)
    AddMsgAndPrint(errorMsg('Create Form, Letter, Report'), 2, textFilePath)
    exit()


### Read and assign values from FSA Addresses Table - select row by FSA Office input ###
# Handle apostrophe in office name for SQL statement
if "'" in fsa_office:
    fsa_office = fsa_office.replace("'", "''")
try:
    fsa_address = {}
    fields = ['FSAOffice', 'FSAAddress', 'FSACITY', 'FSASTATE', 'FSAZIP', 'FSAPHONE', 'FSAFAX', 'FSACounty']
    where_clause = """{0} = '{1}'""".format(AddFieldDelimiters(support_gdb, 'FSAOffice'), fsa_office)
    with SearchCursor(fsa_addresses_table, fields, where_clause) as cursor:
        row = cursor.next()
        fsa_address['office'] = row[0] if row[0] else ''
        fsa_address['street'] = row[1] if row[1] else ''
        fsa_address['city'] = row[2] if row[2] else ''
        fsa_address['state'] = row[3] if row[3] else ''
        fsa_address['zip'] = row[4] if row[4] else ''
        fsa_address['phone'] = row[5] if row[5] else ''
        fsa_address['fax'] = row[6] if row[6] else ''
        fsa_address['county'] = row[7] if row[7] else ''
except:
    AddMsgAndPrint('\nFailed while retrieving FSA Address Table data.\nYou may need to run tool F.Import Office Addresses and then try this tool again. Exiting...', 2, textFilePath)
    AddMsgAndPrint(errorMsg('Create Form, Letter, Report'), 2, textFilePath)
    exit()


### Read and assign values from NAD Addresses Table - select row by State Code from Admin Table ###
if not Exists(nad_addresses_table):
    AddMsgAndPrint('\nNAD Addresses table not found in SUPPORT.gdb. Exiting...', 2, textFilePath)
    exit()
try:
    nad_address = {}
    fields = ['STATECD', 'STATE', 'NADADDRESS', 'NADCITY', 'NADSTATE', 'NADZIP', 'TOLLFREE', 'PHONE', 'TTY', 'FAX']
    where_clause = """{0} = '{1}'""".format(AddFieldDelimiters(support_gdb, 'STATECD'), admin_data['state_code'])
    with SearchCursor(nad_addresses_table, fields, where_clause) as cursor:
        row = cursor.next()
        nad_address['state'] = row[1] if row[1] else ''
        nad_address['street'] = row[2] if row[2] else ''
        nad_address['city'] = row[3] if row[3] else ''
        nad_address['state'] = row[4] if row[4] else ''
        nad_address['zip'] = row[5] if row[5] else ''
        nad_address['toll_free'] = row[6] if row[6] else ''
        nad_address['phone'] = row[7] if row[7] else ''
        nad_address['tty'] = row[8] if row[8] else ''
        nad_address['fax'] = row[9] if row[9] else ''
except:
    AddMsgAndPrint('\nFailed while retrieving NAD Address Table data.\nYou may need to run tool F.Import Office Addresses and then try this tool again. Exiting...', 2, textFilePath)
    AddMsgAndPrint(errorMsg('Create Form, Letter, Report'), 2, textFilePath)
    exit()


### Generate Customer Letter ###
SetProgressorLabel('Generating HELC_Letter.docx...')
try:
    customer_letter_template = DocxTemplate(customer_letter_template_path)
    context = {
        'today_date': date.today().strftime('%A, %B %d, %Y'),
        'admin_data': admin_data,
        'nrcs_address': nrcs_address,
        'fsa_address': fsa_address,
        'fsa_county': fsa_county,
        'nad_address': nad_address
    }
    customer_letter_template.render(context, autoescape=True)
    customer_letter_template.save(customer_letter_output)
    AddMsgAndPrint('\nCreated HELC_Letter.docx...', textFilePath=textFilePath)
except PermissionError:
    AddMsgAndPrint('\nPlease close any open Word documents and try again. Exiting...', 2, textFilePath)
    exit()
except:
    AddMsgAndPrint('\nFailed to create HELC_Letter.docx. Exiting...', 2, textFilePath)
    AddMsgAndPrint(errorMsg('Create Form, Letter, Report'), 2, textFilePath)
    exit()


### Add Numeric Field for CLU Number to Field Determination ###
try:
    SetProgressorLabel('Sorting determination table by numeric CLU field...')
    AddField(field_det_lyr, 'clu_int', 'SHORT')
    CalculateField(field_det_lyr, 'clu_int', 'int(!clu_number!)')
    Sort(field_det_lyr, field_det_sorted_lyr_path, [['clu_int', 'ASCENDING']])
    AddMsgAndPrint('\nAdded CLU integer field and created sorted table...', textFilePath=textFilePath)
except:
    AddMsgAndPrint('\nFailed to create sorted table by CLU. Exiting...', 2, textFilePath)
    AddMsgAndPrint(errorMsg('Create Form, Letter, Report'), 2, textFilePath)
    exit()


### Build Consolidated Determination Table by CLU ###
if consolidate_by_clu:
    try:
        SetProgressorLabel('Consolidating determination data by CLU...')
        table_NHEL_No = os_path.join(site_gdb, 'table_NHEL_No')
        table_NHEL_Yes = os_path.join(site_gdb, 'table_NHEL_Yes')
        table_HEL_No = os_path.join(site_gdb, 'table_HEL_No')
        stats_tables = [table_NHEL_No, table_NHEL_Yes, table_HEL_No]
        stats_fields = [['clu_number','CONCATENATE'], ['clu_calculated_acreage', 'SUM']]
        case_fields = ['HEL_YES', 'sodbust']
        consolidated_table_data = []

        # Table 1: NHEL/No
        where_clause = """{0}='NHEL' and {1}='No'""".format(AddFieldDelimiters(site_gdb, 'HEL_YES'), AddFieldDelimiters(site_gdb, 'sodbust'))
        MakeFeatureLayer(field_det_sorted_lyr_path, 'table1_lyr', where_clause)
        Statistics('table1_lyr', table_NHEL_No, stats_fields, case_fields, ', ')
        try:
            with SearchCursor(table_NHEL_No, ['CONCATENATE_clu_number', 'SUM_clu_calculated_acreage']) as cursor:
                row = cursor.next()
                row_data = {}
                row_data['clu'] = row[0]
                row_data['hel'] = 'NHEL'
                row_data['sodbust'] = 'No'
                row_data['acres'] = f'{row[1]:.2f}'
                consolidated_table_data.append(row_data)
        except StopIteration:
            pass

        # Table 2: NHEL/Yes
        where_clause = """{0}='NHEL' and {1}='Yes'""".format(AddFieldDelimiters(site_gdb, 'HEL_YES'), AddFieldDelimiters(site_gdb, 'sodbust'))
        MakeFeatureLayer(field_det_sorted_lyr_path, 'table2_lyr', where_clause)
        Statistics('table2_lyr', table_NHEL_Yes, stats_fields, case_fields, ', ')
        try:
            with SearchCursor(table_NHEL_Yes, ['CONCATENATE_clu_number', 'SUM_clu_calculated_acreage']) as cursor:
                row = cursor.next()
                row_data = {}
                row_data['clu'] = row[0]
                row_data['hel'] = 'NHEL'
                row_data['sodbust'] = 'Yes'
                row_data['acres'] = f'{row[1]:.2f}'
                consolidated_table_data.append(row_data)
        except StopIteration:
            pass

        # Table 3: HEL/No
        where_clause = """{0}='HEL' and {1}='No'""".format(AddFieldDelimiters(site_gdb, 'HEL_YES'), AddFieldDelimiters(site_gdb, 'sodbust'))
        MakeFeatureLayer(field_det_sorted_lyr_path, 'table3_lyr', where_clause)
        Statistics('table3_lyr', table_HEL_No, stats_fields, case_fields, ', ')
        try:
            with SearchCursor(table_HEL_No, ['CONCATENATE_clu_number', 'SUM_clu_calculated_acreage']) as cursor:
                row = cursor.next()
                row_data = {}
                row_data['clu'] = row[0]
                row_data['hel'] = 'HEL'
                row_data['sodbust'] = 'No'
                row_data['acres'] = f'{row[1]:.2f}'
                consolidated_table_data.append(row_data)
        except StopIteration:
            pass

        # Append HEL/Yes as separate records
        where_clause = """{0}='HEL' and {1}='Yes'""".format(AddFieldDelimiters(site_gdb, 'HEL_YES'), AddFieldDelimiters(site_gdb, 'sodbust'))
        try:
            with SearchCursor(field_det_sorted_lyr_path, ['clu_number', 'clu_calculated_acreage'], where_clause) as cursor:
                for row in cursor:
                    row_data = {}
                    row_data['clu'] = row[0]
                    row_data['hel'] = 'HEL'
                    row_data['sodbust'] = 'Yes'
                    row_data['acres'] = f'{row[1]:.2f}'
                    consolidated_table_data.append(row_data)
        except StopIteration:
            pass

        # Delete Statistics Tables
        for table in stats_tables:
            if Exists(table):
                Delete(table)

        data_026 = consolidated_table_data
        AddMsgAndPrint('\nConsolidated determination data by CLU...', textFilePath=textFilePath)
    except:
        AddMsgAndPrint('\nFailed to consolidate determination data by CLU. Exiting...', 2, textFilePath)
        AddMsgAndPrint(errorMsg('Create Form, Letter, Report'), 2, textFilePath)
        exit()


### Unconsolidated: Assign values from Field Determination feature class ###
else:
    try:
        data_026 = []
        fields = ['clu_number', 'HEL_YES', 'sodbust', 'clu_calculated_acreage']
        with SearchCursor(field_det_sorted_lyr_path, fields) as cursor:
            for row in cursor:
                row_data = {}
                row_data['clu'] = row[0] if row[0] else ''
                row_data['hel'] = row[1] if row[1] else ''
                row_data['sodbust'] = row[2] if row[2] else ''
                row_data['acres'] = f'{row[3]:.2f}' if row[3] else ''
                data_026.append(row_data)
    except:
        AddMsgAndPrint('\nFailed while retrieving CLU Determination table data. Exiting...', 2, textFilePath)
        AddMsgAndPrint(errorMsg('Create Form, Letter, Report'), 2, textFilePath)
        exit()


### Generate Pages 1 and 2 of 026 Form ###
try:
    SetProgressorLabel('Generating NRCS-CPA-026-HELC-Form.docx...')
    cpa_026_helc_template = DocxTemplate(cpa_026_helc_template_path)
    context = {
        'admin_data': admin_data,
        'hel_map_units': hel_map_units,
        'where_completed': where_completed,
        'data_026_pg1': add_blank_rows(data_026, 18) if len(data_026) < 18 else data_026
    }
    cpa_026_helc_template.render(context, autoescape=True)
    cpa_026_helc_template.save(cpa_026_helc_output)
    cpa_026_helc_doc = Document(cpa_026_helc_output)
    cpa_026_helc_composer = Composer(cpa_026_helc_doc)
    AddMsgAndPrint('\nCreated pages 1 and 2 of NRCS-CPA-026-HELC-Form.docx...', textFilePath=textFilePath)
except PermissionError:
    AddMsgAndPrint('\nPlease close any open Word documents and try again. Exiting...', 2, textFilePath)
    exit()
except:
    AddMsgAndPrint('\nFailed to create pages 1 and 2 of NRCS-CPA-026-HELC-Form.docx. Exiting...', 2, textFilePath)
    AddMsgAndPrint(errorMsg('Create Form, Letter, Report'), 2, textFilePath)
    exit()


### Create Summary Statistics Table for Planner Summary Data ###
try:
    Statistics(
        in_table = final_hel_summary_lyr_path,
        out_table = final_hel_stats_table_path,
        statistics_fields = [['Polygon_Acres', 'SUM'], ['clu_calculated_acres', 'MIN']],
        case_field = ['clu_number', 'MUSYM', 'MUHELCL', 'Final_HEL_Value'])
    AddMsgAndPrint('\nCreated Final HEL Summary Statistics table...', textFilePath=textFilePath)
except:
    AddMsgAndPrint('\nFailed to create Final HEL Summary Statistics table. Exiting...', 2, textFilePath)
    AddMsgAndPrint(errorMsg('Create Form, Letter, Report'), 2, textFilePath)
    exit()


### Add Field to Stats Table and Calculate Percentage of Field ###
try:
    AddField(final_hel_stats_table_path, 'percent_of_field', 'DOUBLE')
    CalculateField(final_hel_stats_table_path, 'percent_of_field', '(!SUM_Polygon_Acres!/!MIN_clu_calculated_acres!)*100')
    AddMsgAndPrint('\nCalculated percentage of field in Final HEL Summary Statistics table...', textFilePath=textFilePath)
except:
    AddMsgAndPrint('\nFailed to add field and calculate percentage in Final HEL Summary Statistics table. Exiting...', 2, textFilePath)
    AddMsgAndPrint(errorMsg('Create Form, Letter, Report'), 2, textFilePath)
    exit()


### Package Data into Dictionary for Planner Summary ###
planner_summary_data = {}
with SearchCursor(field_det_sorted_lyr_path, ['clu_number', 'clu_calculated_acreage', 'HEL_YES']) as field_cursor:
    for field_row in field_cursor:
        if field_row[0] not in planner_summary_data: #Should always be true, clu_number should be unique for each row in Field_Determination
            planner_summary_data[field_row[0]] = {'acres': field_row[1], 'class': field_row[2]}
            where_clause = """{0} = '{1}'""".format(AddFieldDelimiters(site_gdb, 'clu_number'), field_row[0])
            rows = []
            with SearchCursor(final_hel_stats_table_path, ['MUHELCL', 'MUSYM', 'Final_HEL_Value', 'SUM_Polygon_Acres', 'percent_of_field'], where_clause) as stats_cursor:
                for stats_row in stats_cursor:
                    rows.append([stats_row[0], stats_row[1], stats_row[2], f"{round(stats_row[3],2):.2f}", f"{round(stats_row[4],2):.2f}"])
                    planner_summary_data[field_row[0]]['rows'] = rows


### Generate Planner Summary ###
SetProgressorLabel('Generating Planner_Summary.docx...')
try:
    planner_summary_template = DocxTemplate(planner_summary_template_path)
    context = {
        'today_date': date.today().strftime('%A, %B %d, %Y'),
        'farm_number': 821,
        'tract_number': 12564,
        'data': planner_summary_data
    }
    planner_summary_template.render(context, autoescape=True)
    planner_summary_template.save(planner_summary_output)
    AddMsgAndPrint('\nCreated Planner_Summary.docx...', textFilePath=textFilePath)
except PermissionError:
    AddMsgAndPrint('\nPlease close any open Word documents and try again. Exiting...', 2, textFilePath)
    exit()
except:
    AddMsgAndPrint('\nFailed to create Planner_Summary.docx. Exiting...', 2, textFilePath)
    AddMsgAndPrint(errorMsg('Create Form, Letter, Report'), 2, textFilePath)
    exit()


# # ### Generate Client Report ### TODO: Create summary stats table and populate Python dict from that
# # SetProgressorLabel('Generating Client_Report.docx...')
# # try:
# #     client_report_template = DocxTemplate(client_report_template_path)
# #     context = {
# #         'today_date': date.today().strftime('%A, %B %d, %Y'),
# #         'farm_number': 821,
# #         'tract_number': 12564,
# #         'data': {
# #             '1': {
# #                 'acres': 10,
# #                 'class': 'HEL',
# #                 'hel': [1, 0.1],
# #                 'hel_phel': [5, 2.5],
# #                 'nhel': [2, 5.12],
# #                 'nhel_phel': [6, 15.2],
# #                 'na': [0, 0]
# #             },
# #             '2': {
# #                 'acres': 26,
# #                 'class': 'NHEL',
# #                 'hel': [1, 0.1],
# #                 'hel_phel': [5, 2.5],
# #                 'nhel': [2, 5.12],
# #                 'nhel_phel': [6, 15.2],
# #                 'na': [0, 0]
# #             }
# #         }
# #     }
# #     client_report_template.render(context, autoescape=True)
# #     client_report_template.save(client_report_output)
# #     AddMsgAndPrint('\nCreated Client_Report.docx...', textFilePath=textFilePath)
# # except PermissionError:
# #     AddMsgAndPrint('\nPlease close any open Word documents and try again. Exiting...', 2, textFilePath)
# #     exit()
# # except:
# #     AddMsgAndPrint('\nFailed to create Client_Report.docx. Exiting...', 2, textFilePath)
# #     AddMsgAndPrint(errorMsg('Create Form, Letter, Report'), 2, textFilePath)
# #     exit()


### Open Customer Letter, 026 Form ###
AddMsgAndPrint('\nFinished generating forms, opening in Microsoft Word. End of script.', textFilePath=textFilePath)
SetProgressorLabel('Finished generating forms, opening in Microsoft Word...')
try:
    startfile(customer_letter_output)
    startfile(cpa_026_helc_output)
    # startfile(client_report_output)
    startfile(planner_summary_output)
except:
    AddMsgAndPrint('\nFailed to open finished forms in Microsoft Word. End of script.', 1, textFilePath)
    AddMsgAndPrint(errorMsg('Create Form, Letter, Report'), 1, textFilePath)
