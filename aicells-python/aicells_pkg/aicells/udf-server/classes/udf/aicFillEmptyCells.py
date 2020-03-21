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

from ..AICFunction import AICFunction
from ... import udfutils
import pandas
import numpy

class aicFillEmptyCells(AICFunction):
    def Run(self, arguments):
        a = self.ProcessArguments(arguments, 'parameters')

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


        if a['selected_columns'] is None:
            a['selected_columns'] = inputData.columns.tolist()

        selectedData = inputData[a['selected_columns']]

        # mean and median works only on numeric columns
        if a['method'] in ['mean', 'median']:
            for c in selectedData.columns:
                if not numpy.issubdtype(selectedData[c].dtype, numpy.number):
                    raise udfutils.AICEParameterError("PARAMETER_NUMERIC_COLUMN_REQUIRED",
                                                      {"parameterName": c})

        if a['method'] == 'value':
            a['method'] = None
            # "ValueError: Must specify a fill 'value' or 'method'.
            if a['value'] is None:
                raise udfutils.AICEParameterError("PARAMETER_SPECIFY_FILL_VALUE",
                                                  {"parameterName": 'value'})

        if a['method'] == 'mean':
            filledData = selectedData.fillna(selectedData.mean(axis=0), limit=a['limit'])
        elif a['method'] == 'median':
            filledData = selectedData.fillna(selectedData.median(axis=0), limit=a['limit'])
        else:
            filledData = selectedData.fillna(value=a['value'], method=a['method'], limit=a['limit'])

        if a['return_all_columns']:
            notSelectedData = inputData.drop(a['selected_columns'], axis=1)
            filledData = pandas.concat([filledData, notSelectedData], axis=1)[inputData.columns.tolist()]

        return udfutils.ReturnDataFrame(filledData, rowHeader=False, columnHeader=True, transpose=a['transpose'])
