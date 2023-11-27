from os import path
from sys import exc_info
from traceback import format_exception

from arcpy import AddError, AddMessage, AddWarning, CreateScratchName, Describe, env, Exists, \
    ParseFieldName, SetProgressorLabel, SpatialReference

from arcpy.analysis import Buffer
from arcpy.management import Clip as Clip_m, Delete, ProjectRaster


def AddMsgAndPrint(msg, severity=0, textFilePath=None):
    """ Adds tool message to the geoprocessor. Split the message on \n first, so a GPMessage will be added for each line."""
    try:
        if textFilePath:
            f = open(textFilePath, 'a+')
            f.write(f"{msg}\n")
            f.close
            del f
        if severity == 0:
            AddMessage(msg)
        elif severity == 1:
            AddWarning(msg)
        elif severity == 2:
            AddError(msg)
    except:
        pass


def errorMsg():
    """ Print traceback exceptions. If sys.exit was trapped by default exception then ignore traceback message."""
    try:
        exc_type, exc_value, exc_traceback = exc_info()
        theMsg = f"\t{format_exception(exc_type, exc_value, exc_traceback)[1]}\n\t{format_exception(exc_type, exc_value, exc_traceback)[-1]}"
        if theMsg.find('sys.exit') > -1:
            AddMsgAndPrint('\n\n')
            pass
        else:
            AddMsgAndPrint('\n\tNRCS HEL Tool Error: -------------------------', 2)
            AddMsgAndPrint(theMsg, 2)
    except:
        AddMsgAndPrint('Unhandled error in errorMsg method', 2)
        pass


def removeScratchLayers(scratchLayers):
    """ This function is the last task that is executed or gets invoked in an except clause. Removes all temporary scratch layers."""
    try:
        for lyr in scratchLayers:
            try:
                Delete(lyr)
            except:
                AddMsgAndPrint(f"\n\tDeleting Layer: {str(lyr)} failed.", 1)
                continue
    except:
        pass


def FindField(layer, chkField):
    """ Check table or featureclass to see if specified field exists. If fully qualified name is found, return that name;
        otherwise return Set workspace before calling FindField."""
    try:
        if Exists(layer):
            theDesc = Describe(layer)
            theFields = theDesc.fields
            theField = theFields[0]
            for theField in theFields:
                # Parses a fully qualified field name into its components (database, owner name, table name, and field name)
                parseList = ParseFieldName(theField.name) # (null), (null), (null), MUKEY
                # choose the last component which would be the field name
                theFieldname = parseList.split(',')[len(parseList.split(','))-1].strip()  # MUKEY
                if theFieldname.upper() == chkField.upper():
                    return theField.name
            return False
        else:
            AddMsgAndPrint('\tInput layer not found')
            return False
    except:
        errorMsg()
        return False


def extractDEMfromImageService(demSource, fieldDetermination, scratchWS, cluLayer, zFactorList, unitLookUpDict, zUnits):
    """ This function will extract a DEM from a Web Image Service that is in WGS. The CLU will be buffered to 410 meters
        and set to WGS84 GCS in order to clip the DEM. The clipped DEM will then be projected to the same coordinate system as the CLU.
        Eventually code will be added to determine the approximate cell size  of the image service using y-distances from the center of the cells.
        Cell size from a WGS84 service is difficult to calculate. Clip is the fastest however it doesn't honor cellsize so a project is required.
        Original Z-factor on WGS84 service cannot be calculated b/c linear units are unknown. Assume linear units and z-units are the same.
        Returns a clipped DEM and new Z-Factor"""
    try:
        desc = Describe(demSource)
        sr = desc.SpatialReference
        outputCellsize = 3

        AddMsgAndPrint(f"\nInput DEM Image Service: {desc.baseName}")
        AddMsgAndPrint(f"\tGeographic Coordinate System: {sr.Name}")
        AddMsgAndPrint(f"\tUnits (XY): {sr.AngularUnitName}")

        # Set output env variables to WGS84 to project clu
        env.geographicTransformations = 'WGS_1984_(ITRF00)_To_NAD_1983'
        env.outputCoordinateSystem = SpatialReference(4326)

        # Buffer CLU by 410 Meters. Output buffer will be in GCS
        cluBuffer = path.join('in_memory', path.basename(CreateScratchName('cluBuffer_GCS', data_type='FeatureClass', workspace=scratchWS)))
        Buffer(fieldDetermination, cluBuffer, '410 Meters', 'FULL', '', 'ALL', '')

        # Use the WGS 1984 AOI to clip/extract the DEM from the service
        cluExtent = Describe(cluBuffer).extent
        clipExtent = f"{str(cluExtent.XMin)} {str(cluExtent.YMin)} {str(cluExtent.XMax)} {str(cluExtent.YMax)}"

        SetProgressorLabel(f"Downloading DEM from {desc.baseName} Image Service")
        AddMsgAndPrint(f"\n\tDownloading DEM from {desc.baseName} Image Service")

        demClip = path.join('in_memory', path.basename(CreateScratchName('demClipIS', data_type='RasterDataset', workspace=scratchWS)))
        Clip_m(demSource, clipExtent, demClip, '', '', '', 'NO_MAINTAIN_EXTENT')

        # Project DEM subset from WGS84 to CLU coord system
        outputCS = Describe(cluLayer).SpatialReference
        env.outputCoordinateSystem = outputCS
        demProject = path.join('in_memory', path.basename(CreateScratchName('demProjectIS', data_type='RasterDataset', workspace=scratchWS)))
        ProjectRaster(demClip, demProject, outputCS, 'BILINEAR', outputCellsize)
        Delete(demClip)

        # Report new DEM properties
        desc = Describe(demProject)
        newSR = desc.SpatialReference
        newLinearUnits = newSR.LinearUnitName
        newCellSize = desc.MeanCellWidth

        # if zUnits not populated assume it is the same as linearUnits
        if not zUnits: zUnits = newLinearUnits
        newZfactor = zFactorList[unitLookUpDict.get(newLinearUnits)][unitLookUpDict.get(zUnits)]

        AddMsgAndPrint(f"\t\tNew Projection Name: {newSR.Name}")
        AddMsgAndPrint(f"\t\tLinear Units (XY): {newLinearUnits}")
        AddMsgAndPrint(f"\t\tElevation Units (Z): {zUnits}")
        AddMsgAndPrint(f"\t\tCell Size: {str(newCellSize)} {newLinearUnits}")
        AddMsgAndPrint(f"\t\tZ-Factor: {str(newZfactor)}")

        return newLinearUnits, newZfactor, demProject

    except:
        errorMsg()


def extractDEM(cluLayer, inputDEM, fieldDetermination, scratchWS, zFactorList, unitLookUpDict, zUnits):
    """ This function will return a DEM that has the same extent as the CLU selected fields buffered to 410 Meters. The DEM can be a local
        raster layer or a web image server. Datum must be in WGS84 or NAD83 and linear units must be in Meters or Feet otherwise it will exit.
        If the cell size is finer than 3M then the DEM will be resampled. The resampling will happen using the Project Raster tool regardless
        of an actual coordinate system change. If the cell size is 3M then the DEM will be clipped using the buffered CLU. Environment settings
        are used to control the output coordinate system. Function does not check the SR of the CLU. Assumes the CLU is in a projected coordinate
        system and in meters. Should probably verify before reprojecting to a 3M cell.
        Returns a clipped DEM and new Z-Factor"""
    try:
        # Set environment variables
        env.geographicTransformations = 'WGS_1984_(ITRF00)_To_NAD_1983'
        env.resamplingMethod = 'BILINEAR'
        env.outputCoordinateSystem = Describe(cluLayer).SpatialReference
        bImageService = False
        bResample = False
        outputCellSize = 3
        desc = Describe(inputDEM)
        sr = desc.SpatialReference
        linearUnits = sr.LinearUnitName
        cellSize = desc.MeanCellWidth

        if desc.format == 'Image Service':
            if sr.Type == 'Geographic':
                newLinearUnits, newZfactor, demExtract = extractDEMfromImageService(inputDEM, fieldDetermination, scratchWS, cluLayer, zFactorList, unitLookUpDict, zUnits)
                return newLinearUnits, newZfactor, demExtract
            bImageService = True

        # Check DEM properties
        if bImageService:
            AddMsgAndPrint(f"\nInput DEM Image Service: {desc.baseName}")
        else:
            AddMsgAndPrint(f"\nInput DEM Information: {desc.baseName}")

        # DEM must be in a project coordinate system to continue unless DEM is an image service.
        if not bImageService and sr.type != 'Projected':
            AddMsgAndPrint(f"\n\t{str(desc.name)} Must be in a projected coordinate system, Exiting", 2)
            AddMsgAndPrint('\tContact your State GIS Coordinator to resolve this issue', 2)
            return False, False

        # Linear units must be in Meters or Feet
        if not linearUnits:
            AddMsgAndPrint('\n\tCould not determine linear units of DEM....Exiting!', 2)
            AddMsgAndPrint('\tContact your State GIS Coordinator to resolve this issue', 2)
            return False, False

        if linearUnits == 'Meter':
            tolerance = 3.1
        elif linearUnits == 'Foot':
            tolerance = 10.1706
        elif linearUnits == 'Foot_US':
            tolerance = 10.1706
        else:
            AddMsgAndPrint(f"\n\tHorizontal units of {str(desc.baseName)} must be in feet or meters... Exiting!")
            AddMsgAndPrint('\tContact your State GIS Coordinator to resolve this issue', 2)
            return False, False

        # Cell size must be the equivalent of 3 meters.
        if cellSize > tolerance:
            AddMsgAndPrint('\n\tThe cell size of the input DEM must be 3 Meters (9.84252 FT) or less to continue... Exiting!', 2)
            AddMsgAndPrint(f"\t{str(desc.baseName)} has a cell size of {str(desc.MeanCellWidth)} {linearUnits}", 2)
            AddMsgAndPrint('\tContact your State GIS Coordinator to resolve this issue', 2)
            return False, False
        elif cellSize < tolerance:
            bResample = True
        else:
            bResample = False

        # if zUnits not populated assume it is the same as linearUnits
        bZunits = True
        if not zUnits: zUnits = linearUnits; bZunits = False

        # look up zFactor based on units (z,xy -- I should've reversed the table)
        zFactor = zFactorList[unitLookUpDict.get(linearUnits)][unitLookUpDict.get(zUnits)]

        # Print Input DEM properties
        AddMsgAndPrint(f"\tProjection Name: {sr.Name}")
        AddMsgAndPrint(f"\tLinear Units (XY): {linearUnits}")
        AddMsgAndPrint(f"\tElevation Units (Z): {zUnits}")
        if not bZunits:
            AddMsgAndPrint(f"\t\tZ-units were auto set to: {linearUnits}")
        AddMsgAndPrint(f"\tCell Size: {str(desc.MeanCellWidth)} {linearUnits}")
        AddMsgAndPrint(f"\tZ-Factor: {str(zFactor)}")

        # Extract DEM
        SetProgressorLabel('Buffering AOI by 410 Meters')
        cluBuffer = path.join('in_memory', path.basename(CreateScratchName('cluBuffer', data_type='FeatureClass', workspace=scratchWS)))
        Buffer(fieldDetermination, cluBuffer, '410 Meters', 'FULL', 'ROUND')
        env.extent = cluBuffer

        # CLU clip extents
        cluExtent = Describe(cluBuffer).extent
        clipExtent = f"{str(cluExtent.XMin)} {str(cluExtent.YMin)} {str(cluExtent.XMax)} {str(cluExtent.YMax)}"

        # Cell Resolution needs to change; Clip and Project
        if bResample:
            SetProgressorLabel(f"Changing resolution from {str(cellSize)} {linearUnits} to 3 Meters")
            AddMsgAndPrint(f"\n\tChanging resolution from {str(cellSize)} {linearUnits} to 3 Meters")
            demClip = path.join('in_memory', path.basename(CreateScratchName('demClip_resample', data_type='RasterDataset', workspace=scratchWS)))
            Clip_m(inputDEM, clipExtent, demClip, '', '', '', 'NO_MAINTAIN_EXTENT')
            demExtract = path.join('in_memory', path.basename(CreateScratchName('demClip_project', data_type='RasterDataset', workspace=scratchWS)))
            ProjectRaster(demClip, demExtract, env.outputCoordinateSystem, env.resamplingMethod, outputCellSize, '#', '#', sr)
            Delete(demClip)

        # Resolution is correct; Clip the raster
        else:
            SetProgressorLabel('Clipping DEM using buffered CLU')
            AddMsgAndPrint('\n\tClipping DEM using buffered CLU')
            demExtract = path.join('in_memory', path.basename(CreateScratchName('demClip', data_type='RasterDataset', workspace=scratchWS)))
            Clip_m(inputDEM, clipExtent, demExtract, '', '', '', 'NO_MAINTAIN_EXTENT')

        # Report any new DEM properties
        desc = Describe(demExtract)
        newSR = desc.SpatialReference
        newLinearUnits = newSR.LinearUnitName
        newCellSize = desc.MeanCellWidth
        newZfactor = zFactorList[unitLookUpDict.get(newLinearUnits)][unitLookUpDict.get(zUnits)]

        if newSR.name != sr.Name:
            AddMsgAndPrint(f"\t\tNew Projection Name: {newSR.Name}")
        if newCellSize != cellSize:
            AddMsgAndPrint(f"\t\tNew Cell Size: {str(newCellSize)} {newLinearUnits}")
        if newZfactor != zFactor:
            AddMsgAndPrint(f"\t\tNew Z-Factor: {str(newZfactor)}")

        Delete(cluBuffer)
        return newLinearUnits, newZfactor, demExtract

    except:
        errorMsg()
        return False, False, False


class NoProcesingExit(Exception):
    pass
