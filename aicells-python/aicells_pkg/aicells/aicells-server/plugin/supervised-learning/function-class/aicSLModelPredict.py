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
import random
import numpy


class aicSLModelPredict(AICFunction):
    def Run(self, arguments):
        a = self.ProcessArguments(arguments, 'parameters')

        try:
            tool = self.factory.CreateInstance('tool-class.' + a['tool_name'].replace('.', '_'))
        except Exception as e:
            raise AICException.AICException("UNKNOWN_TOOL")

        toolA = tool.ProcessArguments(arguments, 'parameters.tool_parameters')

        model = tool.GetModel(toolA)

        trainData = a['train_data']

        predictData = a['predict_data']

        # validate column names
        if not (a['selected_features'] is None):
            for column in a['selected_features']:
                if not isinstance(column, str):
                    raise AICException.AICEParameterError("PARAMETER_INVALID_COLUMN",
                                                          {"parameterName": 'selected_features'})
                if not (column in trainData.columns):
                    raise AICException.AICEParameterError("PARAMETER_UNKNOWN_COLUMN",
                                                          {"parameterName": 'selected_features', "columnName": column})

        # validate target column name
        if not isinstance(a['selected_target'], str):
            raise AICException.AICEParameterError("PARAMETER_INVALID_COLUMN", {"parameterName": 'selected_target'})
        if not (a['selected_target'] in trainData.columns):
            raise AICException.AICEParameterError("PARAMETER_UNKNOWN_COLUMN",
                                                  {"parameterName": 'selected_target',
                                                   "columnName": a['selected_target']})

        if not (a['selected_features'] is None):
            for column in a['selected_features']:
                if not isinstance(column, str):
                    raise AICException.AICEParameterError("PARAMETER_INVALID_COLUMN",
                                                          {"parameterName": 'selected_features'})
                if not (column in predictData.columns):
                    raise AICException.AICEParameterError("PARAMETER_UNKNOWN_COLUMN",
                                                          {"parameterName": 'selected_features', "columnName": column})

        # validate target column name
        # if not isinstance(a['selected_target'], str):
        #     raise AICException.AICEParameterError("PARAMETER_INVALID_COLUMN", {"parameterName": 'selected_target'})
        # if not (a['selected_target'] in predictData.columns):
        #     raise AICException.AICEParameterError("PARAMETER_UNKNOWN_COLUMN",
        #                                       {"parameterName": 'selected_target', "columnName": a['selected_target']})

        # TODO 0 < test_size < 1
        # TODO input_data minimum row number

        # Train Data

        if not (a['selected_features'] is None):
            XTrain = trainData[a['selected_features']]
        else:
            XTrain = trainData.drop(a['selected_target'], axis=1)

        yTrain = trainData[a['selected_target']]

        if not (a['seed'] is None):
            random.seed(a['seed'])
            numpy.random.seed(a['seed'])

        # Predict Data

        if not (a['selected_features'] is None):
            XPredict = predictData[a['selected_features']]
        else:
            XPredict = predictData.drop(a['selected_target'], axis=1)

        try:
            model.fit(XTrain, yTrain)
        except Exception as e:
            errorMessage = ""
            if len(e.args) > 0:
                errorMessage = e.args[0]
            raise AICException.AICException("MODEL_ERROR", {"error": errorMessage})

        yPredict = model.predict(XPredict)

        # ToDO: Numpy array

        yPredict = pandas.DataFrame(yPredict)

        # results_summary is a dataframe
        return UDFUtils.ReturnDataFrame(yPredict, False, a['transpose'])


