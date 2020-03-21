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
import numpy


class aicCorrelationMatrix(AICFunction):
    def Run(self, arguments):
        global returnRange
        a = self.ProcessArguments(arguments, 'parameters')

        inputData = a['input_data']

        if not (a['selected_columns_1'] is None):
            # validate column names
            for column in a['selected_columns_1']:
                if not isinstance(column, str):
                    raise udfutils.AICEParameterError("PARAMETER_INVALID_COLUMN",
                                                      {"parameterName": 'selected_columns_1'})
                if not (column in inputData.columns):
                    raise udfutils.AICEParameterError("PARAMETER_UNKNOWN_COLUMN",
                                                      {"parameterName": 'selected_columns_1', "columnName": column})

        if not (a['selected_columns_2'] is None):
            for column in a['selected_columns_2']:
                if not isinstance(column, str):
                    raise udfutils.AICEParameterError("PARAMETER_INVALID_COLUMN",
                                                      {"parameterName": 'selected_columns_2'})
                if not (column in inputData.columns):
                    raise udfutils.AICEParameterError("PARAMETER_UNKNOWN_COLUMN",
                                                      {"parameterName": 'selected_columns_2', "columnName": column})

        corrColumns = []

        if a['selected_columns_1'] is None:
            corrColumns += inputData.columns.tolist()
        else:
            corrColumns += a['selected_columns_1']

        if a['selected_columns_2'] is None:
            corrColumns += inputData.columns.tolist()
        else:
            corrColumns += a['selected_columns_2']

        corrColumnsUnique = []
        for x in corrColumns:
            if x not in corrColumnsUnique:
                corrColumnsUnique.append(x)

        correlationMatrix = inputData[corrColumnsUnique].corr(method=a['method'], min_periods=a['min_periods'])

        diff = frozenset(corrColumnsUnique).difference(correlationMatrix.columns.tolist())
        if len(diff) != 0:
            raise udfutils.AICEParameterError("CORR_NON_NUMERIC_COLUMN",
                                              {"clolumns": ", ".join(diff)})

        if not (a['selected_columns_1'] is None):
            correlationMatrix = correlationMatrix[a['selected_columns_1']]

        if not (a['selected_columns_2'] is None):
            correlationMatrix = correlationMatrix.loc[a['selected_columns_2'], :]

        if not (a['selected_columns_1'] is None):
            if len(a['selected_columns_1']) == 1:
                correlationMatrix = correlationMatrix.sort_values(a['selected_columns_1'], ascending=False)

        if not (a['selected_columns_2'] is None):
            if len(a['selected_columns_2']) == 1:
                correlationMatrix = correlationMatrix.sort_values(a['selected_columns_2'], ascending=False, axis=1)

        if a['absolute_values']:
            correlationMatrix = correlationMatrix.abs()

        return udfutils.ReturnDataFrame(correlationMatrix, columnHeader=a['display_column_headers'],
                                        rowHeader=a['display_row_headers'], transpose=a['transpose'])

