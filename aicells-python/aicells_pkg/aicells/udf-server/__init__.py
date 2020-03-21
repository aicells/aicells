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

import numpy
import xlwings
import random
import pandas
from . import udfutils

factory = udfutils.AICFactory()

def RunUDF(args):
    # [0] - udf name
    # [1] - udf parameters
    # [2] - namespace list
    # [3-] - ranges

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
        if isinstance(args[1], list):
            if len(args[1]) > 0:
                if (len(args[1]) == 2) and (not isinstance(args[1][0], list)):
                    args[1] = [args[1]]
                for i in range(0, len(args[1])):
                    # parameters range should have exactly 2 columns
                    if len(args[1][i]) != 2:
                        raise udfutils.AICEParameterError("PARAMETER_ERROR", {"parameterName": "parameters"})
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
                                    raise udfutils.AICEParameterError("PARAMETER_COLLISION", {"parameterName": key})
                                rangeDict['parameters'][j][1] = value
                    if not match:
                        rangeDict['parameters'].append([key, value])
    except udfutils.AICEParameterError as e:
        return udfutils.AICErrorToExcelRange(e.GetErrorList())

    try:
        c = factory.CreateInstance("udf." + args[0])
    except Exception as e:
        e = udfutils.AICException("UNKNOWN_FUNCTION")
        return udfutils.AICErrorToExcelRange(e.GetErrorList())

    try:
        result = c.Run(rangeDict)
    except udfutils.AICException as e:
        return udfutils.AICErrorToExcelRange(e.GetErrorList())
    else:
        return result

@xlwings.func
def aicRaw(*args):
    # [0] - udf name
    # [1] - udf parameters
    # [2] - namespace list
    # [3-] - ranges
    sys.stderr.write("Call from VBA " + args[0] + ":\n")

    return RunUDF(args)
    # ret = numpy.random.rand(4, 3)

    # ret = [
    #     ['#AICELLS-ERROR!', '#AICELLS-ERROR@'],
    #     ['code1', 'message1'],
    #     ['code2', 'message2']
    # ]
    # return ret

