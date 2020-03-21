# AIcells (https://github.com/aicells/aicells) - Copyright 2020 Gergely Szerovay, László Siller
#
# Licensed under the Apache License, Version 2.0 (the "License");
# you may not use this file except in compliance with the License.
# You may obtain a copy of the License at
#
#     http://www.apache.org/licenses/LICENSE-2.0
#
# Unless required by applicable law or agreed to in writing, software
# distributed under the License is distributed on an "AS IS" BASIS,
# WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
# See the License for the specific language governing permissions and
# limitations under the License.

import importlib
import traceback

import numpy
import pandas

import sys

class AICException(Exception):
    errorTexts = {
        "PARAMETER_UNKNOWN_COLUMN": "Parameter error: Unknown column: {parameterName} {columnName}",
        "PARAMETER_INVALID_COLUMN": "Parameter error: Invalid column name in: {parameterName}",
        "PARAMETER_ERROR": "Parameter error: {parameterName}",
        "PARAMETER_COLLISION": "Parameter error: {parameterName}",
        "PARAMETER_SPECIFY_FILL_VALUE": "You must specify a fill 'value'",
        "PARAMETER_NUMERIC_COLUMN_REQUIRED": "Parameter error: numeric column required: {parameterName}",
        "UNKNOWN_FUNCTION": "Unknown AICells function",
        "UNKNOWN_TOOL": "Unknown AICells tool",
        "UNKNOWN_SCORER": "Unknown scorer: {scorer}",
        "SCORER_ERROR": "Scorer ({scorer}) error: {error}",
        "UNKNOWN_METRIC": "Unknown metric: {metric}",
        "METRIC_ERROR": "Metric ({metric}) error: {error}",
        "TOO_MANY_ARRAY_DIMENSION": "The array has more than 2 dimensions",
        "FATAL_ERROR": "Fatal error: {error}",
        "MODEL_ERROR": "Model error: {error}",
        "CORR_NON_NUMERIC_COLUMN": "Column selection includes non numeric column(s): {columns}",
    }
    def __init__(self, codeOrErrorList, parameters={}):
        message = ''
        if isinstance(codeOrErrorList, list):
            self.errorList = []
            for err in codeOrErrorList:
                if err[0] in self.errorTexts:
                    message = self.errorTexts[err[0]].format(**err[1])
                else:
                    message = "Unknown error code"
                self.errorList.append([err[0], message])
        else:
            code = codeOrErrorList
            if code in self.errorTexts:
                message = self.errorTexts[code].format(**parameters)
            else:
                message = "Unknown error code"
            self.errorList = [[code, message]]

        sys.stderr.write(f"AICException:\n{message}\n" + traceback.format_exc())
    def GetErrorList(self):
        return self.errorList

class AICEParameterError(AICException):
    pass

def AICErrorToExcelRange(error):
    return [['#AICELLS-ERROR!', '#AICELLS-ERROR@']] + error



#import udf-server.classes

class AICFactory:
    classTypeCache = {}

    def CreateInstance(self, classNameWithNamespace, args={}):

        #classNameWithNamespace = classNameWithNamespace

        if classNameWithNamespace in self.classTypeCache:
            classType = self.classTypeCache[classNameWithNamespace]
        else:
            classNameWithNamespaceAsList = classNameWithNamespace.split('.')
            className = classNameWithNamespaceAsList.pop()

            namespace = ''
            if len(classNameWithNamespaceAsList) != 0:
                namespace = '.'+'.'.join(classNameWithNamespaceAsList)

            #module = importlib.import_module('.'+className, package=__name__+'.classes' + namespace)
            module = importlib.import_module('.'+className, package='udf-server.classes' + namespace)

            classType = getattr(module, className)
            self.classTypeCache[classNameWithNamespace] = classType

        classInstance = classType(**args)
        classInstance.factory = self
        return classInstance

# xlwings default behaviour:
# two Excel rows = [[1,2,3,4,5,6], [1,2,3,4,5,6]]
# one Excel column = [1,2,3,4,5,6]

def ReturnDataFrame(df, columnHeader, rowHeader, transpose=False):
    
    if rowHeader:
        df = df.rename_axis('').reset_index()

    columns = [df.columns.tolist()]
    values = df.values.tolist()

    if columnHeader:
        l = columns + values
    else:
        l = values

    # replace empty strings with None
    if df.isnull().any().any(): # NaN in numeric arrays, None or NaN in object arrays, NaT in datetimelike
        for idx1, v1 in enumerate(l):
            for idx2, v2 in enumerate(l[idx1]):
                if pandas.isnull(v2):
                    l[idx1][idx2] = ""

    if transpose:
        return Transpose2DList(l)
    else:
        return l

def ReturnSeries(series, columnHeader=False, transpose=False):
    if columnHeader:
        l = [series.index.values.tolist(), series.tolist()]
    else:
        l = [series.tolist()]

    # replace empty strings with None
    if series.isnull().any():
        for idx1, v1 in enumerate(l):
            if pandas.isnull(v1):
                l[idx1] = ""

    if transpose:
        return Transpose2DList(l)
    else:
        return l

def ReturnNumpyArray(npArray, transpose=False):
    if npArray.ndim == 1:
        npArray = numpy.reshape(npArray, (-1, 1))

    if npArray.ndim == 2:
        if transpose:
            l = npArray.T.tolist()
        else:
            l = npArray.tolist()

        if pandas.isnull(npArray):
            # replace empty strings with None
            for idx1, v1 in enumerate(l):
                for idx2, v2 in enumerate(l[idx1]):
                    if pandas.isnull(v2):
                        l[idx1][idx2] = ""

        return l

    raise AICException("TOO_MANY_ARRAY_DIMENSION")

def ReturnList(l, transpose=False):
    is2d = False
    if len(l) > 0:
        if isinstance(l[0], list):
            is2d = True

    if not is2d:
        l = [l]
        transpose = not transpose

    # replace empty strings with None
    for idx1, v1 in enumerate(l):
        for idx2, v2 in enumerate(l[idx1]):
            if pandas.isnull(v2):
                l[idx1][idx2] = ""

    if transpose:
        return Transpose2DList(l)
    else:
        return l


def Transpose2DList(l):
    # ret = numpy.array(l).T.tolist() # data type problems
    ret = list(map(list, zip(*l)))
    return ret

def Transpose1DList(l):
    return Transpose2DList([l])

