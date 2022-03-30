# -*- coding: utf-8 -*-
"""
A basic toolbox for summarise_aprx and summarise_mxd.
"""
import os
import arcpy


class Toolbox(object):
    def __init__(self):
        self.label = u'Summarise'
        self.alias = ''
        self.description = 'Demo toolbox for summarise tools'
        # Tool display order is always alphabetical
        self.tools = [SummariseProject]


def set_msg(msg, severity=20):
    """Print a message to Arc tool output window."""
    for line in str(msg).split('\n'):
        line = "    "+line
        if severity == 30:
            arcpy.AddWarning(line)
        elif severity == 40 or severity == 50:
            arcpy.AddError(line)
        else:
            arcpy.AddMessage(line)


class BasePytTool(object):
    # For pyt reload issues, see https://gis.stackexchange.com/questions/91112/refreshing-imported-modules-in-arcgis-python-toolbox
    class ToolValidator(object):
        def __init__(self, parameters):
            """Setup arcpy and the list of tool parameters."""
            self.params = parameters

        def initializeParameters(self):
            """Refine the properties of a tool's parameters.  This method is
            called when the tool is opened."""
            return

        def updateParameters(self):
            """Modify the values and properties of parameters before internal
            validation is performed.  This method is called whenever a parameter
            has been changed."""
            return

        def updateMessages(self):
            """Modify the messages created by internal validation for each tool
            parameter.  This method is called after internal validation."""
            return

    def __init__(self):
        self.label = u'Unnamed Tool'
        self.canRunInBackground = False

    def isLicensed(self):
        return True

    def updateParameters(self, parameters):
        # validator = getattr(self, 'ToolValidator', None)
        # if validator:
        #     return validator(parameters).updateParameters()
        return

    def updateMessages(self, parameters):
        # validator = getattr(self, 'ToolValidator', None)
        # if validator:
        #     return validator(parameters).updateMessages()
        return


class SummariseProject(BasePytTool):
    def __init__(self):
        BasePytTool.__init__(self)
        self.label = u'Summarise Aprx/Mxd'
        arcpy.env.overwriteOutput = True

    def getParameterInfo(self):
        param_inp = arcpy.Parameter()
        param_inp.name = u'Input_Aprx_Mxd'
        param_inp.displayName = u'Input Aprx/Mxd'
        param_inp.parameterType = ESRI_CONSTS.REQUIRED
        param_inp.direction = ESRI_CONSTS.INPUT
        param_inp.datatype = ESRI_CONSTS.DATATYPES.FILE  # ARCMAP_DOCUMENT
        param_inp.filter.list = ['aprx', 'mxd']

        param_outp = arcpy.Parameter()
        param_outp.name = u'Output_File'
        param_outp.displayName = u'Output Xlsx File'
        param_outp.parameterType = ESRI_CONSTS.REQUIRED
        param_outp.direction = ESRI_CONSTS.OUTPUT
        param_outp.datatype = ESRI_CONSTS.DATATYPES.FILE
        param_outp.value = u'C:\\Temp\\arcproj_overview.xlsx'

        param_check_paths = arcpy.Parameter()
        param_check_paths.name = u'Check_Paths'
        param_check_paths.displayName = u'Check Data Paths? (will take a lot longer)'
        param_check_paths.parameterType = ESRI_CONSTS.OPTIONAL
        param_check_paths.direction = ESRI_CONSTS.INPUT
        param_check_paths.datatype = ESRI_CONSTS.DATATYPES.BOOLEAN

        param_open_file = arcpy.Parameter()
        param_open_file.name = u'Open_File'
        param_open_file.displayName = u'Open File When Done'
        param_open_file.parameterType = ESRI_CONSTS.OPTIONAL
        param_open_file.direction = ESRI_CONSTS.INPUT
        param_open_file.datatype = ESRI_CONSTS.DATATYPES.BOOLEAN

        return [param_inp, param_outp, param_check_paths, param_open_file]

    def execute(self, parameters, messages):
        proj = parameters[0].valueAsText
        xls = parameters[1].valueAsText
        check = parameters[2].value
        open = parameters[3].value

        if proj.endswith('.mxd'):
            # sys.path.insert(1, tool_import_path)
            import summarise_mxd
            # sys.path.pop(1)
            summarise_mxd.is_in_arcgis = True
            summarise_mxd.set_msg = set_msg
            summarise_mxd.summarise_mxd(proj, xls, check)

        else:
            # sys.path.insert(1, tool_import_path)
            import summarise_aprx
            # sys.path.pop(1)
            summarise_aprx.is_in_arcgis = True
            summarise_aprx.set_msg = set_msg
            summarise_aprx.summarise_aprx(proj, xls, check)

        if open:
            os.startfile(xls)


class ESRI_CONSTS(object):
    INPUT = u'Input'
    OUTPUT = u'Output'
    REQUIRED = u'Required'
    OPTIONAL = u'Optional'

    # List of datatypes: https://pro.arcgis.com/en/pro-app/arcpy/geoprocessing_and_python/defining-parameter-data-types-in-a-python-toolbox.htm
    class DATATYPES(object):
        ARCMAP_DOCUMENT = u'DEMapDocument'  # u'ArcMap Document'
        BOOLEAN = u'GPBoolean'
        DATASET = u'DEDatasetType'
        DATE = u'GPDate'
        DOUBLE = u'GPDouble'
        FEATURE_CLASS = u'DEFeatureClass'
        FEATURE_DATASET = u'DEFeatureDataset'
        FEATURE_LAYER = u'GPFeatureLayer'
        FILE = u'DEFile'
        FOLDER = u'DEFolder'
        INT = u'GPLong'
        LAYER = u'GPLayer'
        LAYER_FILE = u'DELayer'
        LINEAR_UNIT = u'GPLinearUnit'
        LONG = u'GPLong'
        SPATIAL_REFERENCE = u'GPSpatialReference'  # u'Spatial Reference'
        SQL_EXPRESSION = u'GPSQLExpression'  # u'SQL Expression'
        STRING = u'GPString'
        TABLE = u'DETable'
        WORKSPACE = u'DEWorkspace'
