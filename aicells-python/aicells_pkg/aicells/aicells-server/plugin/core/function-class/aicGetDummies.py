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

from ....AICFunction import AICFunction
from .... import AICException
from .... import UDFUtils
import pandas

class aicGetDummies(AICFunction):
    def Run(self, arguments):
        global returnRange
        a = self.ProcessArguments(arguments, 'parameters')

        inputData = a['input_data']

        if not (a['selected_columns'] is None):
            # validate column names
            for column in a['selected_columns']:
                if not isinstance(column, str):
                    raise AICException.AICEParameterError("PARAMETER_INVALID_COLUMN", {"parameterName": 'selected_columns'})
                if not (column in inputData.columns):
                    raise AICException.AICEParameterError("PARAMETER_UNKNOWN_COLUMN",
                                                      {"parameterName": 'selected_columns', "columnName": column})

        # filter by column names
        dummiesInputData = inputData[a['selected_columns']]
        restInputData = inputData.drop(columns=a['selected_columns'])

        # TODO: adv_param teszt

        dummiesData = pandas.get_dummies(dummiesInputData, prefix=a['prefix'], prefix_sep= a['prefix_sep'],
                                         dummy_na=a['dummy_na'],
                                         drop_first=a['drop_first']) # , sparse=a['sparse']columns=a['columns'],

        if a['full_table']:
            frames = [restInputData, dummiesData]
            returnRange = pandas.concat(frames, axis=1)
        if not a['full_table']:
            returnRange = dummiesData

        return UDFUtils.ReturnDataFrame(returnRange, columnHeader=a['column_header'], rowHeader=False,
                                        transpose=a['transpose']) # rowHeader=a['row_header']