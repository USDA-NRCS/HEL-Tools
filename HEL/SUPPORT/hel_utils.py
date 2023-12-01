from sys import exc_info
from traceback import format_exception

from arcpy import AddError, AddMessage, AddWarning
from arcpy.management import Delete


def addLyrxByConnectionProperties(map, lyr_name_list, lyrx_layer, gdb_path, visible=True):
    ''' Add a layer to a map by setting the lyrx file connection properties.'''
    if lyrx_layer.name not in lyr_name_list:
        lyrx_cp = lyrx_layer.connectionProperties
        lyrx_cp['connection_info']['database'] = gdb_path
        lyrx_cp['dataset'] = lyrx_layer.name
        lyrx_layer.updateConnectionProperties(lyrx_layer.connectionProperties, lyrx_cp)
        map.addLayer(lyrx_layer)

    lyr_list = map.listLayers()
    for lyr in lyr_list:
        if lyr.longName == lyrx_layer.name:
            lyr.visible = visible


def AddMsgAndPrint(msg, severity=0, textFilePath=None):
    ''' Log messages to text file and ESRI tool messages dialog.'''
    if textFilePath:
        with open(textFilePath, 'a+') as f:
            f.write(f"{msg}\n")
    if severity == 0:
        AddMessage(msg)
    elif severity == 1:
        AddWarning(msg)
    elif severity == 2:
        AddError(msg)


def errorMsg(tool_name):
    ''' Return exception details for logging, ignore sys.exit exceptions.'''
    exc_type, exc_value, exc_traceback = exc_info()
    exc_message = f"\t{format_exception(exc_type, exc_value, exc_traceback)[1]}\n\t{format_exception(exc_type, exc_value, exc_traceback)[-1]}"
    if exc_message.find('sys.exit') > -1:
        pass
    else:
        return f"\n\t------------------------- {tool_name} Tool Error -------------------------\n{exc_message}"


def removeScratchLayers(scratchLayers):
    ''' Delete layers in a given list.'''
    for lyr in scratchLayers:
        try:
            Delete(lyr)
        except:
            continue
