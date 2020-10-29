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

from ....AICFunction import AICFunction
from .... import AICException
from .... import UDFUtils

class aicSaveToDataSource(AICFunction):
    def Run(self, arguments):
        a = self.ProcessArguments(arguments, 'parameters')

        try:
            dataSourceClass = self.GetDataSourceClass(arguments['parameters.data_source'])
        except Exception as e:
            raise AICException.AICEParameterError("DATA_SOURCE_UNKNOWN", {"dataSource": 'data_source'})

        try:
            dataSource = self.factory.CreateInstance('tool-class.' + dataSourceClass.replace('.', '_'))
        except Exception as e:
            raise AICException.AICEParameterError("DATA_SOURCE_UNKNOWN", {"dataSource": 'data_source'})

        dataSourceArguments = dataSource.ProcessArguments(arguments, 'parameters.data_source')

        try:
            fn = dataSource.Write(a['input_data'], self.workbookPath, dataSourceArguments, 'data_source')
            return 'Data saved: ' + fn
        except AICException.AICException as e:
            raise

        raise AICException.AICException("DATA_SOURCE_ERROR", {"parameterName": 'data_source'})

