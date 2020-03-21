# AIcells (https://github.com/aicells/aicells) - Copyright 2020 László Siller, Gergely Szerovay
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

from ..AICFunction import AICFunction
from ... import udfutils
import pandas


class aicDescribe(AICFunction):
    def Run(self, arguments):
        a = self.ProcessArguments(arguments, 'parameters')

        for p in ['percentile1', 'percentile2', 'percentile3']:
            if a[p] < 0 or a[p] > 1:
                raise udfutils.AICEParameterError("PARAMETER_ERROR", {"parameterName": p})

        allColumns = ['data type', 'count', 'blank', 'unique', 'top', 'freq', 'mean', 'std', 'min', str(int(a['percentile1']*100)) + '%',
                                 str(int(a['percentile2']*100)) + '%', str(int(a['percentile3']*100)) + '%', 'max']

        inputData = a['input_data']

        # validate column names
        if not (a['selected_columns'] is None):
            for column in a['selected_columns']:
                if not isinstance(column, str):
                    raise udfutils.AICEParameterError("PARAMETER_INVALID_COLUMN",
                                                      {"parameterName": 'selected_columns'})
                if not (column in inputData.columns):
                    raise udfutils.AICEParameterError("PARAMETER_UNKNOWN_COLUMN",
                                                      {"parameterName": 'selected_columns', "columnName": column})

        skeleton = pandas.DataFrame(columns=allColumns)

        if a['selected_columns'] is None:
            selectedData = inputData
        else:
            selectedData = inputData[a['selected_columns']]

        dataType = pandas.DataFrame(selectedData).dtypes.astype(str).rename('data type')
        dataType = dataType.replace('float64', 'number')
        dataType = dataType.replace('object', 'text')
        dataType = dataType.replace('datetime64[ns]', 'date')
        dataType = dataType.replace('bool', 'logical')

        isBlank = selectedData.isnull().sum().sort_values(ascending=False).rename('empty cell')

        percentiles = [a['percentile1'], a['percentile2'], a['percentile3']]

        # describe = selectedData.describe(percentiles=a['percentiles'], include=a['include'], exclude=a['exclude'])

        describe = selectedData.describe(include='all', percentiles=percentiles).transpose()

        describe = pandas.concat([dataType, isBlank, describe], axis=1, sort=False)
        describe = pandas.concat([skeleton, describe], axis=0, sort=False)

        if a['selected_statistics'] is None:
            describe = describe[allColumns]
        else:
            selectedStatistics = a['selected_statistics']
            describe = describe[selectedStatistics]

        return udfutils.ReturnDataFrame(describe, columnHeader=a['display_column_headers'], rowHeader=a['display_row_headers'],
                                        transpose=a['transpose']) # rowHeader=a['display_row_headers']

#  'first', 'last',
#