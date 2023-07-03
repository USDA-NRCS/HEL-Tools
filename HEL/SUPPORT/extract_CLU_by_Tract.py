from json import dumps as json_dumps, loads as json_loads
from os import path
from time import gmtime, sleep, strftime
from urllib import parse, request

import arcpy

from hel_utils import AddMsgAndPrint, errorMsg


def getPortalTokenInfo(portalURL):
    try:
        # i.e. 'https://gis.sc.egov.usda.gov/portal/'
        activePortal = arcpy.GetActivePortalURL()

        # targeted portal is NOT set as default
        if activePortal != portalURL:
            # List of managed portals
            managedPortals = arcpy.ListPortalURLs()

            # portalURL is available in managed list
            if activePortal in managedPortals:
                AddMsgAndPrint("\nYour Active portal is set to: " + activePortal,2)
                AddMsgAndPrint("Set your active portal and sign into: " + portalURL,2)
                return False

            # portalURL must first be added to list of managed portals
            else:
                AddMsgAndPrint("\nYou must add " + portalURL + " to your list of managed portals",2)
                AddMsgAndPrint("Open the Portals Tab to manage portal connections",2)
                AddMsgAndPrint("For more information visit the following ArcGIS Pro documentation:",2)
                AddMsgAndPrint("https://pro.arcgis.com/en/pro-app/help/projects/manage-portal-connections-from-arcgis-pro.htm",1)
                return False

        # targeted Portal is correct; try to generate token
        else:
            # Get Token information
            tokenInfo = arcpy.GetSigninToken()

            # Not signed in.  Token results are empty
            if not tokenInfo:
                AddMsgAndPrint("\nYou are not signed into: " + portalURL,2)
                return False

            # Token generated successfully
            else:
                return tokenInfo

    except:
        errorMsg()
        return False


def submitFSquery(url,INparams):
    """ This function will send a spatial query to a web feature service and convert
        the results into a python structure.  If the results from the service is an
        error due to an invalid token then a second attempt will be sent with using
        a newly generated arcgis token.  If the token is good but the request returned
        with an error a second attempt will be made.  The funciion takes in 2 parameters,
        the URL to the web service and a query string in URLencoded format.

        Error produced with invalid token
        {u'error': {u'code': 498, u'details': [], u'message': u'Invalid Token'}}

        The function returns requested data via a python dictionary"""

    try:
        INparams = INparams.encode('ascii')
        resp = request.urlopen(url,INparams)  # A failure here will probably throw an HTTP exception
        responseStatus = resp.getcode()
        responseMsg = resp.msg
        jsonString = resp.read()

        # json --> Python; dictionary containing 1 key with a list of lists
        results = json_loads(jsonString)

        # Check for expired token; Update if expired and try again
        if 'error' in results.keys():
            if results['error']['message'] == 'Invalid Token':
                AddMsgAndPrint("\tRegenerating ArcGIS Token Information")

                # Get new ArcPro Token, update the original portalToken
                newToken = arcpy.GetSigninToken()
                global portalToken
                portalToken = newToken

                # convert encoded string into python structure and update token
                # by parsing the encoded query strting into list of (name, value pairs)
                # i.e [('f', 'json'),('token','U62uXB9Qcd1xjyX1)]
                # convert to dictionary and update the token in dictionary
                queryString = parseQueryString(INparams)
                requestDict = dict(queryString)
                requestDict.update(token=newToken['token'])
                newParams = urllibEncode(requestDict)
                newParams = newParams.encode('ascii')

                # update incoming parameters just in case a 2nd attempt is needed
                INparams = newParams
                resp = request.urlopen(url,newParams)  # A failure here will probably throw an HTTP exception
                responseStatus = resp.getcode()
                responseMsg = resp.msg
                jsonString = resp.read()
                results = json_loads(jsonString)

        # Check results before returning them; Attempt a 2nd request if results are bad.
        if 'error' in results.keys() or len(results) == 0:
            sleep(5)
            resp = request.urlopen(url,INparams)  # A failure here will probably throw an HTTP exception
            responseStatus = resp.getcode()
            responseMsg = resp.msg
            jsonString = resp.read()
            results = json_loads(jsonString)

            if 'error' in results.keys() or len(results) == 0:
                AddMsgAndPrint("\t2nd Request Attempt Failed - Error Code: " + str(responseStatus) + " -- " + responseMsg + " -- " + str(results),2)
                return False
            else:
                return results

        else:
            return results

    except Exception as e:
        if int(e.code) >= 500:
            pass
        elif int(e.code) >= 400:
            pass
        else:
            AddMsgAndPrint('HTTP ERROR = ' + str(e.code),2)

    except:
        errorMsg()
        return False


def createOutputFC(metadata,outputWS,shape="POLYGON"):
    """ This function will create an empty polygon feature class within the outputWS
        The feature class will be set to the same spatial reference as the Web Feature
        Service. All fields part of the WFS will also be added to the new feature class.
        A field dictionary containing the field names and their property will also be
        returned.  This fieldDict will be used to create the fields in the CLU fc and
        by the getCLUgeometry insertCursor.

        fieldDict ={field:(fieldType,fieldLength,alias)
        i.e {'clu_identifier': ('TEXT', 36, 'clu_identifier'),'clu_number': ('TEXT', 7, 'clu_number')}

        Return the field dictionary and new feature class including the path
        Return False if error ocurred."""

    try:
        # output FC will have the 'CLU_' as a prefix along with the state, county and tract number
        newFC = path.join(outputWS, f"CLU_{str(adminState)}_{str(adminCounty)}_{str(tractNumber)}")

        AddMsgAndPrint("\nCreating New Feature Class: " + path.basename(newFC))
        arcpy.SetProgressorLabel("Creating New Feature Class: " + path.basename(newFC))

        # set the spatial Reference to same as WFS
        # Probably WGS_1984_Web_Mercator_Auxiliary_Sphere
        # {'spatialReference': {'latestWkid': 3857, 'wkid': 102100}
        spatialReferences = metadata['extent']['spatialReference']
        if 'latestWkid' in [sr for sr in spatialReferences.keys()]:
            sr = spatialReferences['latestWkid']
        else:
            sr = spatialReferences['wkid']

        outputCS = arcpy.SpatialReference(sr)

        # fields associated with feature service
        fsFields = metadata['fields']   # {u'alias':u'land_unit_id',u'domain': None, u'name': u'land_unit_id', u'nullable': True, u'editable': True, u'alias': u'LAND_UNIT_ID', u'length': 38, u'type': u'esriFieldTypeString'}
        fieldDict = dict()

        # lookup list for fields that are in DATE field; Date values need to be converted
        # from Unix Epoch format to mm/dd/yyyy format in order to populate a table
        dateFields = list()

        # cross-reference portal attribute description with ArcGIS attribute description
        fldTypeDict = {'esriFieldTypeString':'TEXT','esriFieldTypeDouble':'DOUBLE','esriFieldTypeSingle':'FLOAT',
                       'esriFieldTypeInteger':'LONG','esriFieldTypeSmallInteger':'SHORT','esriFieldTypeDate':'DATE',
                       'esriFieldTypeGUID':'GUID','esriFieldTypeGlobalID':'GUID'}

        # Collect field info to pass to new fc
        for fieldInfo in fsFields:
            # skip the OID field
            if fieldInfo['type'] == 'esriFieldTypeOID':
               continue

            fldType = fldTypeDict[fieldInfo['type']]
            fldAlias = fieldInfo['alias']
            fldName = fieldInfo['name']

            # skip the SHAPE_STArea__ and SHAPE_STLength__ fields
            if fldName.find("SHAPE_ST") > -1:
               continue

            if fldType == 'TEXT':
               fldLength = fieldInfo['length']
            elif fldType == 'DATE':
                 dateFields.append(fldName)
            else:
               fldLength = ""

            fieldDict[fldName] = (fldType,fldLength,fldAlias)

        # Delete newFC if it exists
        if arcpy.Exists(newFC):
            arcpy.Delete_management(newFC)
            AddMsgAndPrint("\t" + path.basename(newFC) + " exists. Deleted")

        # Create empty polygon featureclass with coordinate system that matches AOI.
        arcpy.CreateFeatureclass_management(outputWS, path.basename(newFC), shape, "", "DISABLED", "DISABLED", outputCS)

        # Add fields from fieldDict to mimic WFS
        arcpy.SetProgressor("step", "Adding Fields to " + path.basename(newFC),0,len(fieldDict),1)
        for field,params in fieldDict.items():
            try:
                fldLength = params[1]
                fldAlias = params[2]
            except:
                fldLength = 0
                pass

            arcpy.SetProgressorLabel("Adding Field: " + field)
            arcpy.AddField_management(newFC,field,params[0],"#","#",fldLength,fldAlias)
            arcpy.SetProgressorPosition()

        arcpy.ResetProgressor()
        arcpy.SetProgressorLabel("")
        return fieldDict,newFC

    except:
        errorMsg()
        AddMsgAndPrint("\tFailed to create scratch " + newFC + " Feature Class",2)
        return False


def getCLUgeometryByTractQuery(sqlQuery,fc,RESTurl):
    """ This funciton will retrieve CLU geometry from the CLU WFS and assemble
        into the CLU fc along with the attributes associated with it.
        It is intended to receive requests that will return records that are
        below the WFS record limit"""

    try:
        params = urllibEncode({'f': 'json',
                               'where':sqlQuery,
                               'geometryType':'esriGeometryPolygon',
                               'returnGeometry':'true',
                               'outFields': '*',
                               'token': portalToken['token']})

        # Send request to feature service; The following dict keys are returned:
        # ['objectIdFieldName', 'globalIdFieldName', 'geometryType', 'spatialReference', 'fields', 'features']
        geometry = submitFSquery(RESTurl,params)

        # Error from sumbitFSquery function
        if not geometry:
            return False

        # make sure the request returned geometry; otherwise return False
        if not len(geometry['features']):
            AddMsgAndPrint("\nThere were no CLU fields associated with tract Number " + str(tractNumber),1)
            return False

        # Insert Geometry
        with arcpy.da.InsertCursor(fc, [fld for fld in fields]) as cur:
            arcpy.SetProgressor("step", "Assembling Geometry", 0, len(geometry['features']),1)

            # Iterenate through the 'features' key in geometry dict 'features' contains geometry and attributes
            for rec in geometry['features']:
                arcpy.SetProgressorLabel("Assembling Geometry")
                values = list()    # list of attributes
                polygon = json_dumps(rec['geometry'])   # u'geometry': {u'rings': [[[-89.407702228, 43.334059191999984], [-89.40769642800001, 43.33560779300001]}
                attributes = rec['attributes']          # u'attributes': {u'land_unit_id': u'73F53BC1-E3F8-4747-B51F-E598EE445E47'}}

                for fld in fields:
                    if fld == "SHAPE@JSON":
                        continue

                    # DATE values need to be converted from Unix Epoch format
                    # to dd/mm/yyyy format so that it can be inserted into fc.
                    elif fldsDict[fld][0] == 'DATE':
                        dateVal = attributes[fld]
                        if not dateVal in (None,'null','','Null'):
                            epochFormat = float(attributes[fld]) # 1609459200000
                            # Convert to seconds from milliseconds and reformat
                            localFormat = strftime('%m/%d/%Y',gmtime(epochFormat/1000))   # 01/01/2021
                            values.append(localFormat)
                        else:
                            values.append(None)

                    else:
                        values.append(attributes[fld])

                # geometry goes at the the end
                values.append(polygon)
                cur.insertRow(values)
                arcpy.SetProgressorPosition()

        arcpy.ResetProgressor()
        arcpy.SetProgressorLabel("")
        del cur
        return True

    except:
        try: del cur
        except: pass

        errorMsg()
        return False


def start(state,county,trctNmbr,outSR,outWS,addCLUtoSoftware=False):
    try:
        # Use most of the cores on the machine where ever possible
        arcpy.env.parallelProcessingFactor = "75%"

        global adminState, adminCounty, tractNumber, outSpatialRef, outputWS
        global bUserDefinedSR
        global urllibEncode, parseQueryString
        global portalToken, fldsDict, fields

        adminState = state
        adminCounty = county
        tractNumber = trctNmbr
        outSpatialRef = outSR
        outputWS = outWS

        # user defined an output spatial reference, check if it is an SR object or WKT
        bUserDefinedSR = False
        if outSpatialRef != '':

            # Using a well-known text string representation convert it to a SR object
            if type(outSpatialRef) is str:
                sr = arcpy.SpatialReference()
                sr.loadFromString(outSpatialRef)
                outSpatialRef = sr
            bUserDefinedSR = True

        urllibEncode = parse.urlencode
        parseQueryString = parse.parse_qsl

        # Get the spatial reference of the Active Map from which the tool was invoked
        # and set the WKID as the env.outputCoordSystem
        if not bUserDefinedSR:
            aprx = arcpy.mp.ArcGISProject("CURRENT")
            activeMap = aprx.activeMap

            # If the tool is invoked by the catalog view vs the catalog pane
            # an active map will not be registered so a coordinate system
            # object cannot be obtained.  Exit.
            try:
                activeMapName = activeMap.name
                activeMapSR = activeMap.getDefinition('V2').spatialReference['latestWkid']
                outSpatialRef = arcpy.SpatialReference(activeMapSR)
                arcpy.env.outputCoordinateSystem = outSpatialRef
            except:
                AddMsgAndPrint("Could not obtain Spatial Reference from the ArcGIS Pro Map",2)
                AddMsgAndPrint("Run the tool from an active map or provide a Coordinate System. Exiting!",2)
                exit()

        arcpy.env.overwriteOutput = True
        arcpy.env.geographicTransformations = "WGS_1984_(ITRF00)_To_NAD_1983"

        nrcsPortal = 'https://gis.sc.egov.usda.gov/portal/'
        portalToken = getPortalTokenInfo(nrcsPortal)

        if not portalToken:
           AddMsgAndPrint("Could not generate Portal Token. Exiting!",2)
           exit()

        # URL for Feature Service Metadata (Service Definition) - Dictionary of ;
        cluRESTurl_Metadata = """https://gis.sc.egov.usda.gov/appserver/rest/services/common_land_units/common_land_units/FeatureServer/0"""

        # Used for admin or feature service info; Send POST request
        params = urllibEncode({'f': 'json','token': portalToken['token']})

        # request info about the feature service
        fsMetadata = submitFSquery(cluRESTurl_Metadata,params)

        # Create empty CLU FC with necessary fields
        # fldsDict - {'clu_number': ('TEXT', 7, 'clu_number')}
        fldsDict,cluFC = createOutputFC(fsMetadata,outputWS)

        # Isolate the fields that were inserted into new fc
        fields = fldsDict.keys()

        # Convert to a list b/c Python 3.6 doesn't support .append
        fields = list(fields)

        fields.append('SHAPE@JSON')

        # query by Admin State, Admin County and Tract Number
        cluRESTurl = """https://gis.sc.egov.usda.gov/appserver/rest/services/common_land_units/common_land_units/FeatureServer/0/query"""

        # ADMIN_STATE = 29 AND ADMIN_COUNTY = 017 AND TRACT_NUMBER = 1207
        # This was updated for Alaska purpose only b/c Alaska doesn't use Admin_county, they use county_ansi_code
        if adminState == '02':
            whereClause = "ADMIN_STATE = " + str(adminState) + " AND COUNTY_ANSI_CODE = " + str(adminCounty) + " AND TRACT_NUMBER = " + str(tractNumber)
        else:
            whereClause = "ADMIN_STATE = " + str(adminState) + " AND ADMIN_COUNTY = " + str(adminCounty) + " AND TRACT_NUMBER = " + str(tractNumber)

        AddMsgAndPrint("Querying USDA-NRCS GeoPortal for CLU fields where: " + whereClause)

        # Send geometry request to cluREST API
        if not getCLUgeometryByTractQuery(whereClause,cluFC,cluRESTurl):
            try:
                arcpy.Delete_management(cluFC)
            except:
                pass

            # Script executed directly from ArcGIS Pro
            if addCLUtoSoftware:
                AddMsgAndPrint("Exiting",1)
                exit()

            # Script executed from another script
            else:
                return False

        # Report # of fields assembled
        numOfCLUs = int(arcpy.GetCount_management(cluFC)[0])
        if numOfCLUs > 1:
            AddMsgAndPrint("\nThere are " + str(numOfCLUs) + " CLU fields associated with tract number " + str(tractNumber))
        else:
            AddMsgAndPrint("\nThere is " + str(numOfCLUs) + " CLU field associated with tract number " + str(tractNumber))

        # Project cluFC to user-defined spatial reference or the spatial reference set in the AcrGIS Pro Map
        fromSR = arcpy.Describe(cluFC).spatialReference
        toSR = outSpatialRef

        geoTransformation = arcpy.ListTransformations(fromSR,toSR)
        if len(geoTransformation):
            geoTransformation = geoTransformation[0]
            msg = 1
        else:
            geoTransformation = None
            msg = 0

        projected_CLU = cluFC + "_prj"
        arcpy.Project_management(cluFC,projected_CLU,toSR,geoTransformation)

        arcpy.Delete_management(cluFC)
        arcpy.Rename_management(projected_CLU,projected_CLU[0:-4])
        cluFC = projected_CLU[0:-4]

        AddMsgAndPrint(" ",msg)
        AddMsgAndPrint("\nProjecting CLU Feature class",msg)
        AddMsgAndPrint("FROM: " + str(fromSR.name),msg)
        AddMsgAndPrint("TO: " + str(toSR.name),msg)
        AddMsgAndPrint("Geographic Transformation used: " + str(geoTransformation),msg)

        # Add final CLU layer to either ArcPro or ArcMap
        if addCLUtoSoftware:
            # Add the data to the first ArcPro Map found
            aprx = arcpy.mp.ArcGISProject("CURRENT")
            aprxMaps = aprx.listMaps()

            try:
                activeMap = aprx.activeMap
                activeMapName = activeMap.name
                for map in aprxMaps:
                    if map.name == activeMapName:
                        map.addDataFromPath(cluFC)
                        AddMsgAndPrint(path.basename(cluFC) + " added to " + map.name + " Map")
                        break

            except:
                map = aprx.listMaps()[0]
                map.addDataFromPath(cluFC)
                AddMsgAndPrint(path.basename(cluFC) + " added to " + map.name + " Map")

        else:
            return cluFC

    except:
        errorMsg()
