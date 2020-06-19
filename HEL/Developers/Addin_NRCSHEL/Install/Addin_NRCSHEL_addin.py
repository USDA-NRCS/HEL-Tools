import arcpy
import pythonaddins

class Button1(object):
    """Implementation for Addin_NRCSHEL_addin.button1 (Button)"""
    def __init__(self):
        self.enabled = True
        self.checked = False
    def onClick(self):
        toolPath = r'C:\HEL\-- NRCS HEL Determination.tbx'
        pythonaddins.GPToolDialog(toolPath,'NRCSHELDeterminationTool')

class Button2(object):
    """Implementation for Addin_NRCSHEL_addin.button2 (Button)"""
    def __init__(self):
        self.enabled = True
        self.checked = False
    def onClick(self):
        toolPath = toolPath = r'C:\HEL\-- NRCS HEL Determination.tbx'
        pythonaddins.GPToolDialog(toolPath,'DownloadNRCSElevationData')

class Button3(object):
    """Implementation for Addin_NRCSHEL_addin.button3 (Button)"""
    def __init__(self):
        self.enabled = True
        self.checked = False
    def onClick(self):
        toolPath = toolPath = r'C:\HEL\-- NRCS HEL Determination.tbx'
        pythonaddins.GPToolDialog(toolPath,'MergeHELSoilData')

class Button4(object):
    """Implementation for Addin_NRCSHEL_addin.button4 (Button)"""
    def __init__(self):
        self.enabled = True
        self.checked = False
    def onClick(self):
        toolPath = toolPath = r'C:\HEL\-- NRCS HEL Determination.tbx'
        pythonaddins.GPToolDialog(toolPath,'MergeLocalDEMs')
