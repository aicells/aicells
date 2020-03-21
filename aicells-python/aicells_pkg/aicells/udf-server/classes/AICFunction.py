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

import yaml
import pathlib
import os
from ..udfutils import AICEParameterError, AICException
import pandas


class AICFunction:
    def __init__(self):
        self.parameters = []

        filesDir = str(pathlib.Path(__file__).parent.absolute())
        directories = [
            'tool-1-draft',
            'tool-2-alpha',
            'tool-3-beta',
            'tool-4-production',
            'udf-1-draft',
            'udf-2-alpha',
            'udf-3-beta',
            'udf-4-production',
        ]

        self.y = None
        for d in directories:
            yamlFile = os.path.join(filesDir, '..', 'yml\\' + d, self.__class__.__name__ + '.yml')
            if os.path.isfile(yamlFile):
                with open(yamlFile) as file:
                    self.y = yaml.load(file, Loader=yaml.FullLoader)

        if self.y is None:
            raise AICException("FATAL_ERROR", {"error": "YML file not found, class: " + self.__class__.__name__})

        self.parameters = []
        if 'parameters' in self.y:
            if isinstance(self.y['parameters'], list):
                self.parameters = self.y['parameters']

        self.parameterNameList = []
        self.parameterTypeList = []
        self.parameterDefaultList = []
        self.parameterDescriptionList = []

        for parameter in self.parameters:
            self.parameterNameList.append(parameter['parameterName'])
            self.parameterTypeList.append(", ".join(parameter['type']))
            if 'default' in parameter:
                self.parameterDefaultList.append(parameter['default'])
            else:
                self.parameterDefaultList.append('(required)')

            if 'description' in parameter:
                self.parameterDescriptionList.append(parameter['description'])
            else:
                self.parameterDescriptionList.append('')

    def LoadDataYML(self, fileName):
        filesDir = str(pathlib.Path(__file__).parent.absolute())
        yamlFile = os.path.join(filesDir, '..', 'yml\\', fileName)
        if os.path.isfile(yamlFile):
            with open(yamlFile) as file:
                return yaml.load(file, Loader=yaml.FullLoader)
        raise AICException("FATAL_ERROR", {"error": "YML file not found: " + fileName})

    def GetDescription(self):
        self.parameters = []
        if 'description' in self.y:
            return self.y['description']
        else:
            return ''

    def GetParameterNameList(self):
        return self.parameterNameList

    def GetParameterTypeList(self):
        return self.parameterTypeList

    def GetParameterDefaultList(self):
        return self.parameterDefaultList

    def GetParameterDescritpionList(self):
        return self.parameterDescriptionList

    def FlatternList(self, l):
        if not isinstance(l, list):
            return [l]
        flatList = []
        for subList in l:
            if isinstance(subList, list):
                for item in subList:
                    flatList.append(item)
            else:
                flatList.append(subList)
        return flatList

    def ProcessArguments(self, args, argsKey):
        # self.args = args
        kwargs = {}
        for r in args[argsKey]:
            try:
                kwargs[r[0]] = r[1]
            except Exception as e:
                raise AICEParameterError("PARAMETER_ERROR", {'parameterName': argsKey})
        kwargsCleaned = {}
        errors = []
        for parameter in self.parameters:
            parameterName = parameter['parameterName']

            if 'default' in parameter:
                if not (parameterName in kwargs):
                    kwargs[parameterName] = parameter['default']
                elif kwargs[parameterName] is None:
                    kwargs[parameterName] = parameter['default']

            parameterTypes = parameter['type']
            if not isinstance(parameterTypes, list):
                parameterTypes = [parameterTypes]

            for parameterType in parameterTypes:
                if (parameterName in kwargs) and (not (parameterName in kwargsCleaned)):
                    try:
                        if parameterType == 'string':
                            kwargsCleaned[parameterName] = self._String(parameterName, kwargs[parameterName])
                        if parameterType == 'set':
                            if not ('setValues' in parameter):
                                raise AICException("FATAL_ERROR", {"error": "setValues not defined"})
                            kwargsCleaned[parameterName] = self._Set(parameterName, kwargs[parameterName], parameter['setValues'])
                        if parameterType == 'parameters':
                            kwargsCleaned[parameterName] = None
                        # if parameterType == 'model_parameters':
                        #     kwargsCleaned[parameterName] = self._ModelParameters(parameterName, kwargs[parameterName])
                        if parameterType == 'float':
                            kwargsCleaned[parameterName] = self._Float(parameterName, kwargs[parameterName])
                        if parameterType == 'boolean':
                            kwargsCleaned[parameterName] = self._Boolean(parameterName, kwargs[parameterName])
                        if parameterType == 'integer':
                            kwargsCleaned[parameterName] = self._Integer(parameterName, kwargs[parameterName])
                        if (parameterType == 'list') or (parameterType == 'series'):
                            listItems = None
                            listKey = argsKey + '.' + parameterName
                            if listKey in args:
                                listItems = args[listKey]
                            if parameterType == 'list':
                                kwargsCleaned[parameterName] = self._List(parameterName, kwargs[parameterName], listItems)
                            if parameterType == 'series':
                                l = self._List(parameterName, kwargs[parameterName], listItems)
                                if isinstance(l, list):
                                    kwargsCleaned[parameterName] = pandas.Series(l)
                        if parameterType == 'dataframe':
                            list2d = None
                            listKey = argsKey + '.' + parameterName
                            if listKey in args:
                                list2d = args[listKey]
                                # TODO: ranges with 1 rows?
                            columnHeader = False
                            if 'columnHeader' in parameter:
                                if parameter['columnHeader'] == True:
                                    columnHeader = True
                            kwargsCleaned[parameterName] = self._DataFrame(parameterName, kwargs[parameterName], list2d, columnHeader)
                        if parameterType == 'Null':
                            kwargsCleaned[parameterName] = self._None(parameterName, kwargs[parameterName])
                        if parameterType == 'False':
                            kwargsCleaned[parameterName] = self._False(parameterName, kwargs[parameterName])
                    except AICEParameterError as e:
                        pass
            # } for parameterType in parameterTypes:

            if not (parameterName in kwargsCleaned):
                errors += [["PARAMETER_ERROR", {'parameterName': parameterName}]]

        if len(errors) != 0:
            raise AICEParameterError(errors)

        # self.kwargs = kwargsCleaned
        return kwargsCleaned


    def _Set(self, parameterName, x, setValues):
        if not isinstance(x, str):
            raise AICEParameterError("PARAMETER_ERROR", {"parameterName": parameterName})
        else:
            if x == "@AICELLS-RANGE@":
                raise AICEParameterError("PARAMETER_ERROR", {"parameterName": parameterName})
            if not (x in setValues):
                raise AICEParameterError("PARAMETER_ERROR", {"parameterName": parameterName})
            return x

    def _String(self, parameterName, x):
        if not isinstance(x, str):
            raise AICEParameterError("PARAMETER_ERROR", {"parameterName": parameterName})
        else:
            if x == "@AICELLS-RANGE@":
                raise AICEParameterError("PARAMETER_ERROR", {"parameterName": parameterName})
            return x

    # def _ModelParameters(self, parameterName, x):
    #     if not (x is None):
    #         if x[0][0] != False:
    #             if len(x[0]) != 2:
    #                 raise AICEParameterError("PARAMETER_ERROR", {"parameterName": parameterName})
    #             else:
    #                 parameterDict = {}
    #                 for row in x:
    #                     parameterDict[row[0]] = row[1]
    #                 return parameterDict

    def _Float(self, parameterName, x):
        if not (x is None):
            if isinstance(x, float):
                return x
            if isinstance(x, int):
                return float(x)
        raise AICEParameterError("PARAMETER_ERROR", {"parameterName": parameterName})

    def _Boolean(self, parameterName, x):
        if not (x is None):
            if isinstance(x, bool):
                return x
        raise AICEParameterError("PARAMETER_ERROR", {"parameterName": parameterName})

    def _Integer(self, parameterName, x):
        if not (x is None):
            if isinstance(x, int):
                return x
            if isinstance(x, float):
                if x.is_integer():
                    return int(x)
            raise AICEParameterError("PARAMETER_ERROR", {"parameterName": parameterName})

    def _List(self, parameterName, x, listItems):
        if not (listItems is None):
            return self.FlatternList(listItems)
        raise AICEParameterError("PARAMETER_ERROR", {"parameterName": parameterName})

    def _DataFrame(self, parameterName, x, list2d, columnHeader):
        if not (list2d is None):
            if columnHeader:
                columns = list2d[0]
                if not isinstance(columns, list):
                    columns = [columns]
                df = pandas.DataFrame(list2d[1:], columns=columns)
            else:
                df = pandas.DataFrame(list2d)
            return df
        raise AICEParameterError("PARAMETER_ERROR", {"parameterName": parameterName})

    def _None(self, parameterName, x):
        if x is None:
            return x
        raise AICEParameterError("PARAMETER_ERROR", {"parameterName": parameterName})

    def _False(self, parameterName, x):
        if not (x is None):
            if isinstance(x, bool):
                if x == False:
                    return x
        raise AICEParameterError("PARAMETER_ERROR", {"parameterName": parameterName})
