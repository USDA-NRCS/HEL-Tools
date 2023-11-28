from sys import exc_info
from traceback import format_exception

from arcpy import AddError, AddMessage, AddWarning


def AddMsgAndPrint(msg, severity=0, textFilePath=None):
    ''' Log messages to text file and ESRI tool messages dialog.'''
    try:
        if textFilePath:
            with open(textFilePath, 'a+') as f:
                f.write(f"{msg}\n")
        if severity == 0:
            AddMessage(msg)
        elif severity == 1:
            AddWarning(msg)
        elif severity == 2:
            AddError(msg)
    except:
        pass


def errorMsg(tool_name):
    ''' Return exception details for logging, ignore sys.exit exceptions.'''
    exc_type, exc_value, exc_traceback = exc_info()
    exc_message = f"\t{format_exception(exc_type, exc_value, exc_traceback)[1]}\n\t{format_exception(exc_type, exc_value, exc_traceback)[-1]}"
    if exc_message.find('sys.exit') > -1:
        pass
    else:
        return f"\n\t------------------------- {tool_name} Tool Error -------------------------\n{exc_message}"
