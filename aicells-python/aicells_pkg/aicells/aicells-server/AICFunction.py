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
# import pathlib
import os
from .AICException import AICEParameterError, AICException, AICEUnknownParameter
import pandas
from . import UDFUtils

class AICEParameterNotMatch(AICException):
    pass

class AICFunction:
    def Init(self):
        self.queue = None
        self.parameters = []
        self.workbookPath = None

        self.y = None
        for d in ['function', 'tool']:
            yamlFile = os.path.join(os.path.dirname(self.classFile), '..', d + '-yml\\', self.__class__.__name__ + '.yml')
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
        yamlFile = os.path.join(os.path.dirname(self.classFile), '..', 'yml\\', fileName)
        if os.path.isfile(yamlFile):
            with open(yamlFile) as file:
                return yaml.load(file, Loader=yaml.FullLoader)
        raise AICException("FATAL_ERROR", {"error": "YML file not found: " + fileName})

    def GetTag(self):
        return self.y['tag']

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

    def SetQueue(self, q):
        self.queue = q

    def SetConfig(self, config):
        self.config = config

    def Progress(self, text):
        if self.queue:
            self.queue.put(['progress', text])

    def ProcessArguments(self, args, argsKey):
        errors = []
        kwargs = {}

        parameterNamePrefix = argsKey[11:].strip(".")
        if parameterNamePrefix != "":
            parameterNamePrefix += '.'

        if len(args[argsKey]) == 2:
            if len(args[argsKey][0]) != 2:
                # vertical parameter table => transpose it
                args[argsKey] = UDFUtils.Transpose2DList(args[argsKey])

        for r in args[argsKey]:
            if not isinstance(r, list):
                errors += [["PARAMETER_ERROR", {'parameterName': argsKey}]]
            elif len(r) != 2:
                errors += [["PARAMETER_ERROR", {'parameterName': argsKey}]]
            else:
                if (r[0] is None) or (r[0] == "") or (r[0] in ['function', 'output']):
                    pass
                elif r[0] == "_workbook_path":
                    self.workbookPath = r[1]
                else:
                    if not (r[0] in self.parameterNameList):
                        raise AICEUnknownParameter("PARAMETER_UNKNOWN", {'parameterName': parameterNamePrefix + r[0]})
                    kwargs[r[0]] = r[1]

        kwargsCleaned = {}
        # errors = []
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
                        if parameterType == 'data_source':
                            kwargsCleaned[parameterName] = self._DataSource(parameterName, kwargs[parameterName], parameterNamePrefix, args, self.workbookPath)
                        if parameterType == 'string':
                            kwargsCleaned[parameterName] = self._String(parameterName, kwargs[parameterName])
                        if parameterType == 'set':
                            if not ('setValues' in parameter):
                                raise AICException("FATAL_ERROR", {"error": "setValues not defined"})
                            kwargsCleaned[parameterName] = self._Set(parameterName, kwargs[parameterName], parameter['setValues'])
                        if parameterType == 'parameters':
                            kwargsCleaned[parameterName] = None
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
                            isDataSource = False
                            try:
                                kwargsCleaned[parameterName] = self._DataSource(parameterName, kwargs[parameterName], parameterNamePrefix, args, self.workbookPath)
                                isDataSource = True
                            except AICEParameterNotMatch as e:
                                pass

                            if not isDataSource:
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
                    except AICEParameterNotMatch as e:
                        #errors += e.GetErrorListRaw()
                        pass
            # } for parameterType in parameterTypes:

            if not (parameterName in kwargsCleaned):
                errors += [["PARAMETER_INVALID_TYPE", {'parameterName': parameterName}]]

        if len(errors) != 0:
            raise AICEParameterError(errors)

        # self.kwargs = kwargsCleaned
        return kwargsCleaned


    def _Set(self, parameterName, x, setValues):
        if not isinstance(x, str):
            AICEParameterNotMatch()
        else:
            if x == "@AICELLS-RANGE@":
                AICEParameterNotMatch()
            if not (x in setValues):
                AICEParameterNotMatch()
            return x
        
    def GetDataSourceClass(self, arr2d):
        if not isinstance(arr2d, list):
            raise AICException()
        if len(arr2d) == 0:
            raise AICException()
        if not isinstance(arr2d[0], list):
            raise AICException()

        if not isinstance(arr2d[0][0], str):
            raise AICException()
        if arr2d[0][0] != 'data_source':
            raise AICException()

        if len(arr2d[0]) == 2:
            # horizontal parameter range
            if not isinstance(arr2d[0][1], str):
                raise AICException()
            dataSourceClass = arr2d[0][1]
        elif len(arr2d) == 2:
            # vertical parameter range
            if not isinstance(arr2d[1][0], str):
                raise AICException()
            dataSourceClass = arr2d[1][0]
        return dataSourceClass

    def _DataSource(self, parameterName, x, parameterNamePrefix, args, workbookPath):
        if not 'parameters.' + parameterNamePrefix + parameterName in args:
            raise AICEParameterNotMatch()

        try:
            dataSourceClass = self.GetDataSourceClass(args['parameters.' + parameterNamePrefix + parameterName])
        except Exception as e:
            raise AICEParameterNotMatch()
        
        try:
            dataSource = self.factory.CreateInstance('tool-class.' + dataSourceClass.replace('.', '_'))
        except Exception as e:
            raise AICEParameterError("DATA_SOURCE_UNKNOWN", {"dataSource": args['parameters.' + parameterNamePrefix + parameterName][0][1]})

        dataSourceArguments = dataSource.ProcessArguments(args, 'parameters.' + parameterNamePrefix + parameterName)

        try:
            return dataSource.Read(workbookPath, dataSourceArguments, parameterNamePrefix + parameterName)
        except AICException as e:
            raise
        except Exception as e:
            raise AICEParameterError("DATA_SOURCE_ERROR", {"parameterName": parameterNamePrefix + parameterName})

    def _String(self, parameterName, x):
        if not isinstance(x, str):
            raise AICEParameterNotMatch()
        else:
            if x == "@AICELLS-RANGE@":
                raise AICEParameterNotMatch()
            return x

    def _Float(self, parameterName, x):
        if not (x is None):
            if isinstance(x, float):
                return x
            if isinstance(x, int):
                return float(x)
        raise AICEParameterNotMatch()

    def _Boolean(self, parameterName, x):
        if not (x is None):
            if isinstance(x, bool):
                return x
        raise AICEParameterNotMatch()

    def _Integer(self, parameterName, x):
        if not (x is None):
            if isinstance(x, int):
                return x
            if isinstance(x, float):
                if x.is_integer():
                    return int(x)
            raise AICEParameterNotMatch()

    def _List(self, parameterName, x, listItems):
        if not (listItems is None):
            return self.FlatternList(listItems)
        raise AICEParameterNotMatch()

    def _DataFrame(self, parameterName, x, list2d, columnHeader):
        if not (list2d is None):
            if columnHeader:
                columns = list2d[0]
                if not isinstance(columns, list): # single column
                    columns = [columns]
                df = pandas.DataFrame(list2d[1:], columns=columns)
            else:
                df = pandas.DataFrame(list2d)
            return df
        raise AICEParameterNotMatch()

    def _None(self, parameterName, x):
        if x is None:
            return x
        raise AICEParameterNotMatch()

    def _False(self, parameterName, x):
        if not (x is None):
            if isinstance(x, bool):
                if x == False:
                    return x
        raise AICEParameterNotMatch()
