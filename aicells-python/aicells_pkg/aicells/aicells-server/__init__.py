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

import sys
import yaml
import os

import pandas
import xlwings
import multiprocessing
import time

from . import UDFUtils
from . import AICFactory
from . import AICException

aicConfig = {}

globalProcess = None
globalQueue = None

def RunUDFProcessWrapper(args, timeStart, q=None):
    result = RunUDF(args, q)
    print(timeStart)
    print(time.time())
    diff = time.time() - timeStart
    q.put(['debug', f"Python function run time (wo. data transfer): {diff:.2f}s"])
    q.put(['result', result])

def RunUDF(args, q=None):
    # [0] - udf name
    # [1] - udf parameters
    # [2] - namespace list
    # [3-] - ranges

    global globalOutputRange

    args = list(args) # tuple to list

    # build a dictionary from the ranges, the keys are the names in the namespace
    rangeDict = {}
    for i in range(0, 10):
        if not (args[2][i] is None):
            idx0 = i + 3

            # single value passed, convert to 2d
            if (not isinstance(args[idx0], list)):
                args[idx0] = [[args[idx0]]]
            else:
                # single 1d row passed, convert to 2d
                if (not isinstance(args[idx0][0], list)):
                    args[idx0] = [args[idx0]]

            # replace empty strings with None
            for idx1, v1 in enumerate(args[idx0]):
                for idx2, v2 in enumerate(v1):
                    if isinstance(v2, str):
                        if v2 == "":
                            args[idx0][idx1][idx2] = None

            rangeDict[args[2][i]] = args[idx0]

    # merge the udf arguments into the base namespace
    try:
        if not ("parameters" in rangeDict):
            rangeDict['parameters'] = []

        if len(rangeDict['parameters']) == 2:
            if len(rangeDict['parameters'][0]) != 2:
                # vertical parameter table => transpose it
                rangeDict['parameters'] = UDFUtils.Transpose2DList(rangeDict['parameters'])

        if isinstance(args[1], list):
            if len(args[1]) > 0:
                if (len(args[1]) == 2) and (not isinstance(args[1][0], list)):
                    args[1] = [args[1]]

                for i in range(0, len(args[1])):
                    # parameters range should have exactly 2 columns
                    if len(args[1][i]) != 2:
                        raise AICException.AICEParameterError("PARAMETER_ERROR", {"parameterName": "parameters"})
                    key = args[1][i][0]
                    value = args[1][i][1]
                    match = False
                    for j in range(0, len(rangeDict['parameters'])):
                        if key == rangeDict['parameters'][j][0]:
                            match = True
                            if value == '@AICELLS-RANGE@':
                                rangeDict['parameters'][j][1] = value
                            elif not (value is None):
                                # if (not (rangeDict['parameters'][j][1] is None)) and (rangeDict['parameters'][j][1] != ""):
                                if not (rangeDict['parameters'][j][1] is None):
                                    raise AICException.AICEParameterError("PARAMETER_COLLISION", {"parameterName": key})
                                rangeDict['parameters'][j][1] = value
                    if not match:
                        rangeDict['parameters'].append([key, value])
    except AICException.AICEParameterError as e:
        return AICException.AICErrorToExcelRange(e.GetErrorList())

    dataSourceClass = None

    try:
        c = factory.CreateInstance("function-class." + args[0])
        c.SetQueue(q)
        c.SetConfig(aicConfig)
    except Exception as e:
        e = AICException.AICException("UNKNOWN_FUNCTION")
        return AICException.AICErrorToExcelRange(e.GetErrorList())

    if 'parameters.output' in rangeDict:
        try:
            dataSourceClass = c.GetDataSourceClass(rangeDict['parameters.output'])
        except Exception as e:
            pass

    try:
        result = c.Run(rangeDict)
    except AICException.AICException as e:
        return AICException.AICErrorToExcelRange(e.GetErrorList())

    if 'svg' in c.GetTag():
        return [['#AICELLS-SVG!'] + result]

    if dataSourceClass is None:
        return result
    else:
        try:
            try:
                dataSource = factory.CreateInstance('tool-class.' + dataSourceClass.replace('.', '_'))
            except Exception as e:
                raise AICException.AICEParameterError("DATA_SOURCE_UNKNOWN", {"data_source": 'data_source'})

            dataSourceArguments = dataSource.ProcessArguments(rangeDict, 'parameters.output')

            try:
                header = False
                if 'header' in dataSourceArguments:
                    if dataSourceArguments['header']:
                        header = True
                if header:
                    columns = result[0]
                    if not isinstance(columns, list): # single column
                        columns = [columns]
                    df = pandas.DataFrame(result[1:], columns=columns)
                else:
                    df = pandas.DataFrame(result)
                fn = dataSource.Write(df, c.workbookPath, dataSourceArguments, 'output')
                return fn # 'Output saved: '
            except AICException.AICException as e:
                raise

            # raise AICException.AICException("DATA_SOURCE_ERROR", {"parameterName": 'data_source'})

        except AICException.AICException as e:
            return AICException.AICErrorToExcelRange(e.GetErrorList())
        except Exception as e:
            errorMessage = ""
            if len(e.args) > 0:
                errorMessage = e.args[0]
            e2 = AICException.AICException("FATAL_ERROR", {"error": errorMessage})
            return AICException.AICErrorToExcelRange(e2.GetErrorList())

@xlwings.func
def aicUDFRunner(*args):
    # [0] - udf name
    # [1] - udf parameters
    # [2] - namespace list
    # [3-] - ranges

    sys.stderr.write("Call from VBA UDF " + args[0] + ":\n")

    return RunUDF(args)

@xlwings.func
def aicProcessRunner(*args):
    # [0] - udf name
    # [1] - udf parameters
    # [2] - namespace list
    # [3-] - ranges
    global globalProcess, globalQueue

    sys.stderr.write("Call from VBA Runner tool " + args[0] + ":\n")

    if not (globalProcess is None):
        return 'ERROR'

    globalQueue = multiprocessing.Queue()
    globalProcess = multiprocessing.Process(target=RunUDFProcessWrapper, args=(args, time.time(), globalQueue))
    globalProcess.start()

    return 'OK'

@xlwings.func
def aicQueueGet(*args):
    global globalProcess, globalQueue

    if globalProcess is None:
        return ['empty', 'empty']

    if not (globalQueue is None):
        if not globalQueue.empty():
            queueItem = globalQueue.get()
            if queueItem[0] == 'result':
                globalProcess = None
                globalQueue = None
            return queueItem
        else:
            return ['empty', 'empty']

@xlwings.func
def aicAbortProcess(*args):
    global globalProcess, globalQueue

    sys.stderr.write("Call from VBA Dialog Runner: aicAbortProcess\n")

    if not (globalProcess is None):
        if globalProcess.is_alive():
            globalProcess.terminate()
    globalProcess = None
    globalQueue = None


print('Loading aicells-config.yml ...')

dir = os.path.dirname(os.path.realpath(__file__))
yamlFile = dir + '\\..\\..\\..\\..\\aicells-config.yml'
with open(yamlFile) as file:
    aicConfig = yaml.load(file, Loader=yaml.FullLoader)

factory = AICFactory.AICFactory(aicConfig)
