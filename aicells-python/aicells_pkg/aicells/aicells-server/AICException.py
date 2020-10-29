import traceback
import sys

class AICException(Exception):
    errorTexts = {
        "PARAMETER_UNKNOWN_COLUMN": "Parameter error: Unknown column \"{columnName}\" in: {parameterName}",
        "PARAMETER_INVALID_COLUMN": "Parameter error: Invalid column name in: {parameterName}",
        "PARAMETER_ERROR": "Parameter error: {parameterName}",
        "PARAMETER_COLLISION": "Parameter error: {parameterName}",
        "PARAMETER_SPECIFY_FILL_VALUE": "You must specify a fill 'value'",
        "PARAMETER_NUMERIC_COLUMN_REQUIRED": "Parameter error: numeric column required: {parameterName}",
        "PARAMETER_UNKNOWN": "Unknown parameter: {parameterName}",

        "PARAMETER_INVALID_TYPE": "Parameter error: The type of the parameter value is invalid: {parameterName}",

        "UNKNOWN_FUNCTION": "Unknown AICells function",
        "UNKNOWN_TOOL": "Unknown AICells tool",
        "UNKNOWN_SCORER": "Unknown scorer: {scorer}",
        "SCORER_ERROR": "Scorer ({scorer}) error: {error}",
        "UNKNOWN_METRIC": "Unknown metric: {metric}",
        "METRIC_ERROR": "Metric ({metric}) error: {error}",
        "TOO_MANY_ARRAY_DIMENSION": "The array has more than 2 dimensions",
        "FATAL_ERROR": "Fatal error: {error}",
        "MODEL_ERROR": "Model error: {error}",
        "TOOL_ERROR": "Tool error: {error}",
        "CORR_NON_NUMERIC_COLUMN": "Column selection includes non numeric column(s): {columns}",

        "DATA_SOURCE_UNKNOWN": "Unknown data reader: {data_source}",
        "DATA_SOURCE_ERROR": "Data reader error: {parameterName}",
        "DATA_SOURCE_READ_ONLY": "Can't write, data source is read only.",
    }
    def __init__(self, codeOrErrorList="", parameters={}):
        message = ''
        if isinstance(codeOrErrorList, list):
            self.errorList = []
            self.errorListRaw = []
            for err in codeOrErrorList:
                if err[0] in self.errorTexts:
                    message = self.errorTexts[err[0]].format(**err[1])
                else:
                    message = "Unknown error code"
                self.errorList.append([err[0], message])
                self.errorListRaw.append(err)
        else:
            code = codeOrErrorList
            if code in self.errorTexts:
                message = self.errorTexts[code].format(**parameters)
            else:
                message = "Unknown error code"
            self.errorList = [[code, message]]
            self.errorListRaw = [[codeOrErrorList, parameters]]


        sys.stderr.write(f"AICException:\n{message}\n" + traceback.format_exc())
    def GetErrorList(self):
        return self.errorList
    def GetErrorListRaw(self):
        return self.errorListRaw

class AICEParameterError(AICException):
    pass

class AICEUnknownParameter(AICException):
    pass

def AICErrorToExcelRange(error):
    return [['#AICELLS-ERROR!', '#AICELLS-ERROR@']] + error
