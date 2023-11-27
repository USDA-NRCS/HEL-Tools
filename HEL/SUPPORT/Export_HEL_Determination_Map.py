from getpass import getuser
from json import loads as json_loads
from os import path, rename, startfile
from time import ctime
from urllib.parse import urlencode
from urllib.request import urlopen

from arcpy import Describe, env, Exists, GetParameter, GetParameterAsText, ListFeatureClasses, ListRasters, ListTables, SetProgressorLabel
from arcpy.da import Editor, SearchCursor
from arcpy.management import Compact, Delete, GetCount
from arcpy.mp import ArcGISProject

from hel_utils import AddMsgAndPrint, errorMsg


def logBasicSettings(textFilePath, zoom_type, imagery, show_location, overwrite_layout):
    with open(textFilePath, 'a+') as f:
        f.write('\n######################################################################\n')
        f.write('Executing Tool: Export HEL Determination Map\n')
        f.write(f"User Name: {getuser()}\n")
        f.write(f"Date Executed: {ctime()}\n")
        f.write('User Parameters:\n')
        f.write(f"\tZoom Method: {zoom_type}\n")
        f.write(f"\tImagery Layer: {imagery}\n")
        if show_location:
            f.write('\tShow PLSS Location Text Box: True\n')
        else:
            f.write('\tShow PLSS Location Text Box: False\n')
        if overwrite_layout:
            f.write('\tOverwrite Determination Map: True\n')
        else:
            f.write('\tOverwrite Determination Map: False\n')


def submitFSquery(url, INparams):
    INparams = INparams.encode('ascii')
    resp = urlopen(url, INparams)
    jsonString = resp.read()
    results = json_loads(jsonString)
    if 'error' in results.keys():
        return False
    else:
        return results


def getPLSS(plss_point):
    # URLs for BLM services
    tr_svc = 'https://gis.blm.gov/arcgis/rest/services/Cadastral/BLM_Natl_PLSS_CadNSDI/MapServer/1/query'   # Town and Range
    sec_svc = 'https://gis.blm.gov/arcgis/rest/services/Cadastral/BLM_Natl_PLSS_CadNSDI/MapServer/2/query'  # Sections

    AddMsgAndPrint('\tChecking input PLSS reference point...')
    plssDesc = Describe(plss_point)
    plss_fc = plssDesc.catalogPath
    if plssDesc.shapeType != 'Point':
        if plss_fc.find('SCRATCH.gdb') > 0:
            Delete(plss_point)
        AddMsgAndPrint('\nThe input PLSS location digitizing was not a Point layer.', 2)
        AddMsgAndPrint('\nPlease digitize a single point in the input point parameter and try again. Exiting...', 2)
        exit()
    else:
        plss_fc = plssDesc.catalogPath
        result = int(GetCount(plss_fc).getOutput(0))
        if result != 1:
            if plss_fc.find('SCRATCH.gdb') > 0:
                Delete(plss_point)
            AddMsgAndPrint('\nThe input PLSS location layer contains a number of points other than 1.', 2)
            AddMsgAndPrint('\nPlease digitize a single point in the input point parameter and try again. Exiting...', 2)
            exit()

    AddMsgAndPrint('\tInput PLSS location reference is a single point. Using point to query BLM PLSS layers...')
    jsonPoint = [row[0] for row in SearchCursor(plss_fc, ['SHAPE@JSON'])][0]

    # Set input parameters for a query to get a count of results
    params = urlencode({
        'f': 'json',
        'geometry':jsonPoint,
        'geometryType':'esriGeometryPoint',
        'returnCountOnly':'true'})
    
    # Run and check the count query
    AddMsgAndPrint('\tQuerying BLM Township and Range Layer...')
    mer_txt = ''
    countQuery = submitFSquery(tr_svc, params)
    if countQuery:
        returned_records = countQuery['count']
        if returned_records > 0:
            # Run actual query for the town and range fields
            params = urlencode({
                'f': 'json',
                'geometry':jsonPoint,
                'geometryType':'esriGeometryPoint',
                'returnGeometry':'false',
                'outFields':'PRINMER,TWNSHPNO,TWNSHPDIR,RANGENO,RANGEDIR'})
            trQuery = submitFSquery(tr_svc,params)
            rdict = trQuery['features'][0]
            adict = rdict['attributes']
            mer_txt = adict['PRINMER']
            town_no = int(adict['TWNSHPNO'])
            town_dir = adict['TWNSHPDIR']
            range_no = int(adict['RANGENO'])
            range_dir = adict['RANGEDIR']

            if len(mer_txt) > 0 and town_no > 0 and range_no > 0:
                # Query count the sections service
                params = urlencode({
                    'f': 'json',
                    'geometry':jsonPoint,
                    'geometryType':'esriGeometryPoint',
                    'returnCountOnly':'true'})
                # Run and check the count query
                AddMsgAndPrint('\tQuerying BLM Sections Layer...')
                countQuery = submitFSquery(sec_svc, params)
                if countQuery:
                    returned_records = countQuery['count']
                    if returned_records > 0:
                        # Run actual query for the section field
                        params = urlencode({
                            'f': 'json',
                            'geometry':jsonPoint,
                            'geometryType':'esriGeometryPoint',
                            'returnGeometry':'false',
                            'outFields':'FRSTDIVNO'})
                        secQuery = submitFSquery(sec_svc, params)
                        rdict = secQuery['features'][0]
                        adict = rdict['attributes']
                        section_no = int(adict['FRSTDIVNO'])
                        if section_no > 0:
                            return f"Location: T{str(town_no)}{town_dir}, R{str(range_no)}{range_dir}, Sec {str(section_no)}\n{mer_txt}"
    return None


### Initial Tool Validation ###
try:
    aprx = ArcGISProject('CURRENT')
    m = aprx.listMaps('HEL Determination')[0]
except:
    AddMsgAndPrint('\nThis tool must be run from an active ArcGIS Pro project that was developed from the template distributed with this toolbox. Exiting...\n', 2)
    exit()


### ESRI Environment Settings ###
env.overwriteOutput = True


### Input Parameters ###
AddMsgAndPrint('Reading inputs...\n')
SetProgressorLabel('Reading inputs...')
field_determination_lyr = GetParameterAsText(0)
zoom_type = GetParameterAsText(1)
zoom_lyr = GetParameterAsText(2)
show_location = GetParameter(3)
plss_point = GetParameterAsText(4)
overwrite_layout = GetParameter(5)
imagery = GetParameterAsText(6)
if '\\' in imagery:
    imagery = imagery.split('\\')[-1]


### Set Local Variables and Paths ###
AddMsgAndPrint('Setting variables...\n')
SetProgressorLabel('Setting variables...')
base_dir = path.abspath(path.dirname(__file__)) #\SUPPORT
scratch_gdb = path.join(base_dir, 'scratch.gdb')
support_gdb = path.join(base_dir, 'SUPPORT.gdb')

helc_fd = path.dirname(Describe(field_determination_lyr).catalogPath)
helc_gdb = path.dirname(helc_fd)
helc_dir = path.dirname(helc_gdb)

userWorkspace = path.dirname(helc_dir)
projectName = path.basename(userWorkspace).replace(' ', '_')
textFilePath = path.join(userWorkspace, f"{projectName}_log.txt")
basedata_gdb = path.join(userWorkspace, f"{projectName}_BaseData.gdb")
projectTable = path.join(basedata_gdb, f"Table_{projectName}")


### Check For Unsaved Edits ###
workspace = helc_gdb
edit = Editor(workspace)
if edit.isEditing:
    AddMsgAndPrint("\nThere are unsaved edits in this project. Please Save or Discard them, then run this tool again. Exiting...", 2)
    exit()

# Start logging to text file
logBasicSettings(textFilePath, zoom_type, imagery, show_location, overwrite_layout)

### Main Procedure ###
try:

    ### Setup Output PDF File Name(s) ###
    SetProgressorLabel("Configuring output file names...")
    outPDF = path.join(helc_dir, f"Determination_Map_{projectName}.pdf")
    # If overwrite existing maps is checked, use standard file name, else enumerate
    if not overwrite_layout:
        if path.exists(outPDF):
            count = 1
            while count > 0:
                outPDF = path.join(helc_dir, f"Determination_Map_{projectName}_{str(count)}.pdf")
                if path.exists(outPDF):
                    count += 1
                else:
                    count = 0


    #### Check if the output PDF file name(s) is currently open in another application.
    # If so, then close it so we can overwrite without an IO error during export later.
    # We can't use Python to check active threads for a file without a special module.
    # We have to get creative with modules available to us, such as os.
    if path.exists(outPDF):
        try:
            rename(outPDF, f"{outPDF}_opentest")
            rename(f"{outPDF}_opentest", outPDF)
            AddMsgAndPrint('\nThe PDF is available to overwrite...', textFilePath=textFilePath)
        except:
            AddMsgAndPrint('\nThe Determination Map PDF file is open or in use by another program. Please close the PDF and try running this tool again. Exiting!', 2, textFilePath)
            exit()
    else:
        AddMsgAndPrint('\nThe Determination Map PDF file does not exist for this project and will be created...', textFilePath=textFilePath)


    ### If Chosen, Get PLSS Data ###
    # Set starting boolean for location text box. Stays false unless all query criteria to show location are met
    display_dm_location = False
    dm_plss_text = ''
        
    if show_location:
        AddMsgAndPrint('\nShow location selected for Determination Map. Processing reference location...', textFilePath=textFilePath)
        SetProgressorLabel('Retrieving PLSS Location...')
        dm_plss_text = getPLSS(plss_point)
        if dm_plss_text != '':
            AddMsgAndPrint('\nThe PLSS query was successful and a location text box will be shown on the Determination Map...', textFilePath=textFilePath)
            display_dm_location = True
               
    # If any part of the PLSS query failed, or if show location was not enabled, then do not show the Location text box
    if display_dm_location == False:
        AddMsgAndPrint('\nEither the Show Location parameter was not enabled or the PLSS query failed...', 1, textFilePath)
        AddMsgAndPrint('\nTown, Range, Section text box will not be shown on the Determination Map...', 1, textFilePath)
            

    ### Gather Data for Header from Project Table ###
    if Exists(projectTable):
        AddMsgAndPrint('\nCollecting header information from project table...', textFilePath=textFilePath)
        fields = ['admin_county_name', 'county_name', 'farm_number', 'tract_number', 'client', 'dig_staff']
        with SearchCursor(projectTable, fields) as cursor:
            row = cursor.next()
            admin_county = row[0] if row[0] else ''
            geo_county = row[1] if row[1] else ''
            farm = row[2] if row[2] else ''
            tract = row[3] if row[3] else ''
            client = row[4] if row[4] else ''
            dig_staff = row[5] if row[5] else ''
    

    ### Retrieve Map Layout Object ###
    try:
        layout = aprx.listLayouts('HEL Determination Layout')[0]
    except:
        AddMsgAndPrint('\nCould not find installed HEL Determination Layout. Exiting!', 2, textFilePath)
        exit()
    
    ### Update Dynamic Text Objects in Layout ###
    AddMsgAndPrint('\nUpdating dynamic text in layout...', textFilePath=textFilePath)
    SetProgressorLabel('Updating dynamica text in layout...')
    try:
        location_element = layout.listElements('TEXT_ELEMENT', 'Location')[0]
        farm_element = layout.listElements('TEXT_ELEMENT', 'Farm')[0]
        tract_element = layout.listElements('TEXT_ELEMENT', 'Tract')[0]
        geoco_element = layout.listElements('TEXT_ELEMENT', 'GeoCo')[0]
        adminco_element = layout.listElements('TEXT_ELEMENT', 'AdminCo')[0]
        customer_element = layout.listElements('TEXT_ELEMENT', 'Customer')[0]
        imagery_element = layout.listElements('TEXT_ELEMENT', 'Imagery Text Box')[0]
    except:
        AddMsgAndPrint(f"\nOne or more expected elements are missing or had its name changed in the {layout.name}...", 2, textFilePath)
        AddMsgAndPrint('\nLayout cannot be updated automatically. Import the appropriate layout from the installation folder and try again. Exiting!', 2, textFilePath)
        exit()

    if dm_plss_text != '' and display_dm_location:
        location_element.text = dm_plss_text
        location_element.visible = True
    else:
        location_element.visible = False
        location_element.text = 'Location: '

    farm_element.text = f"Farm: {farm}" if farm else 'Farm: <Not Found>'
    tract_element.text = f"Tract: {tract}" if tract else 'Tract: <Not Found>'
    geoco_element.text = f"Geographic County: {geo_county}" if geo_county else 'Geographic County: <Not Found>'
    adminco_element.text = f"Administrative County: {admin_county}" if admin_county else 'Administrative County: <Not Found>'
    customer_element.text = f"Customer: {client}" if client else 'Customer: <Not Found>'
    imagery_element.text = f" Image: {imagery}" if imagery else ' Image: '

    ### Configure Layer Visibility and Zoom ###
    AddMsgAndPrint('\nConfiguring layer visibility and layout extent...', textFilePath=textFilePath)
    SetProgressorLabel('Configuring layer visibility and layout extent...')
    
    # Turn off PLSS layer if used
    plss_lyr = ''
    if plss_point:
        plssDesc = Describe(plss_point)
        if plssDesc.dataType == 'FeatureLayer':
            try:
                plss_lyr = m.listLayers(plss_point)[0]
                plss_lyr.visible = False
            except:
                pass

    # Zoom to specified layer extent if applicable
    if zoom_type == 'Zoom to Layer':
        mf = layout.listElements('MAPFRAME_ELEMENT', 'Map Frame')[0]
        lyr = m.listLayers(zoom_lyr)[0]
        ext = mf.getLayerExtent(lyr)
        cam = mf.camera
        cam.setExtent(ext)
        cam.scale = cam.scale*1.25 if cam.scale > 3960 else 3960

    
    # Turn off the imagery element in legend
    # TODO: Configure other legend elements
    legend = layout.listElements('LEGEND_ELEMENT')[0]
    for item in legend.items:
        if item.name == imagery:
            item.visible = False
                

    ### Export Map to PDF ###
    AddMsgAndPrint('\nExporting the Determination Map to PDF...', textFilePath=textFilePath)
    SetProgressorLabel('Exporting Determination Map...')
    layout.exportToPDF(outPDF, resolution=300, image_quality='NORMAL', layers_attributes='LAYERS_ONLY', georef_info=True)


    ### Reset Imagery Text Box to Blank ###
    try:
        imagery_element.text = ' Image: '
    except:
        pass


    ### Clean Up Scratch GDB ###
    AddMsgAndPrint('\nClearing Scratch GDB...', textFilePath=textFilePath)
    SetProgressorLabel('Clearing Scratch GDB...')
    env.workspace = scratch_gdb
    
    fcs = [path.join(scratch_gdb, fc) for fc in ListFeatureClasses('*')]
    for fc in fcs:
        if Exists(fc):
            try:
                Delete(fc)
            except:
                pass

    rasters = [path.join(scratch_gdb, ras) for ras in ListRasters('*')]
    for ras in rasters:
        if Exists(ras):
            try:
                Delete(ras)
            except:
                pass

    tables = [path.join(scratch_gdb, tbl) for tbl in ListTables('*')]
    for tbl in tables:
        if Exists(tbl):
            try:
                Delete(tbl)
            except:
                pass

    
    ### Compact Project's Base Data and HELC GDBs ###
    try:
        AddMsgAndPrint('\nCompacting File Geodatabases...', textFilePath=textFilePath)
        SetProgressorLabel('Compacting File Geodatabases...')
        Compact(basedata_gdb)
        Compact(helc_gdb)
    except:
        pass


    ### Open Customer Letter, 026 Form ###
    AddMsgAndPrint('\nOpening exported PDF file in default PDF viewer...', textFilePath=textFilePath)
    SetProgressorLabel('Opening exported PDF file...')
    try:
        startfile(outPDF)
    except:
        AddMsgAndPrint('\nWARNING: Failed to open PDF file in default PDF viewer...', 1, textFilePath=textFilePath)


    AddMsgAndPrint('\nScript completed successfully', textFilePath=textFilePath)

except SystemExit:
    pass

except:
    errorMsg()
