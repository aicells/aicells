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
import sklearn
import pandas
import numpy
import random

class aicSLModelMetrics(AICFunction):
    def Run(self, arguments):
        a = self.ProcessArguments(arguments, 'parameters')

        try:
            tool = self.factory.CreateInstance('tool.'+a['AIcells_tool_name'].replace('.', '_'))
        except Exception as e:
            raise udfutils.AICException("UNKNOWN_TOOL")

        tool_a = tool.ProcessArguments(arguments, 'parameters.tool_parameters')
        
        model = tool.GetModel(tool_a)

        input_data = a['input_data']

        # validate column names
        if not (a['selected_features'] is None):
            for column in a['selected_features']:
                if not isinstance(column, str):
                    raise udfutils.AICEParameterError("PARAMETER_INVALID_COLUMN", {"parameterName": 'selected_features'})
                if not (column in input_data.columns):
                    raise udfutils.AICEParameterError("PARAMETER_UNKNOWN_COLUMN", {"parameterName": 'selected_features', "columnName": a['selected_target']})

        # validate target column name
        if not isinstance(a['selected_target'], str):
            raise udfutils.AICEParameterError("PARAMETER_INVALID_COLUMN", {"parameterName": 'selected_target'})
        if not (a['selected_target'] in input_data.columns):
            raise udfutils.AICEParameterError("PARAMETER_UNKNOWN_COLUMN",
                                              {"parameterName": 'selected_target', "columnName": a['selected_target']})

        # TODO 0 < test_size <  1
        # TODO input_data minimum row number

        if not (a['selected_features'] is None):
            X = input_data[a['selected_features']]
        else:
            X = input_data.drop(a['selected_target'], axis=1)

        if not (a['seed'] is None):
            random.seed(a['seed'])
            numpy.random.seed(a['seed'])

        Y = input_data[a['selected_target']]

        X_train, X_test, y_train, y_test = sklearn.model_selection.train_test_split(
            X,
            Y,
            test_size=a['test_size']
        )

        try:
            model.fit(X_train, y_train)
        except Exception as e:
            errorMessage = ""
            if len(e.args) > 0:
                errorMessage = e.args[0]
            raise udfutils.AICException("MODEL_ERROR", {"error": errorMessage})

        y_pred_train = model.predict(X_train)
        y_pred_test = model.predict(X_test)

        # mae_percent_train = sklearn.metrics.mean_absolute_error(y_train, y_pred_train,
        #                                                         multioutput='raw_values') / numpy.mean(y_train)
        # mae_percent_test = sklearn.metrics.mean_absolute_error(y_test, y_pred_test,
        #                                                        multioutput='raw_values') / numpy.mean(y_test)

        # stirng (=regression currently) or list of metrics
        if isinstance(a['selected_metrics'], str):
            if a['selected_metrics'] in self.y['groups']:
                metricList = self.y['groups'][a['selected_metrics']]
            else:
                metricList = [a['selected_metrics']]
        else:
            metricList = a['selected_metrics']

        # prefix metrics list with train_ and test_
        metricListTrainAndTest = []
        for m in metricList:
            metricListTrainAndTest.append(("train_"+m).replace('_', ' '))
            metricListTrainAndTest.append(("test_"+m).replace('_', ' '))

        # crete an empty dataframe
        dataFrame = pandas.DataFrame(columns=metricListTrainAndTest)

        # a single row of metrics in the dataframe
        rowDict = {}
        # loop on each selected metrics
        for selectedMetric in metricList:
            # search for a given metric in the yml
            for metric in self.y['metrics']:
                if metric['name'] == selectedMetric:
                    kwargs = {}
                    if 'arguments' in metric:
                        if isinstance(metric['arguments'], dict):
                            kwargs = metric['arguments']

                    sign = 1
                    if a['greater_is_better']:
                        if 'greater_is_better' in metric:
                            if not metric['greater_is_better']:
                                sign = -1

                    function = metric['function'].split('(')[0].split('.') # eg. sklearn, metrics, median_absolute_error
                    f1 = getattr(sklearn, function[1]) # eg. metrics
                    f2 = getattr(f1, function[2]) # eg. median_absolute_error

                    try:
                        rowDict[("train_"+selectedMetric).replace('_', ' ')] = sign * f2(y_train, y_pred_train, **kwargs)
                        rowDict[("test_"+selectedMetric).replace('_', ' ')] = sign * f2(y_test, y_pred_test, **kwargs)
                    except Exception as e:
                        errorMessage = ""
                        if len(e.args) > 0:
                            errorMessage = e.args[0]
                        raise udfutils.AICException("METRIC_ERROR", {"error": errorMessage, 'metric': selectedMetric})

            if not (("train_"+selectedMetric).replace('_', ' ') in rowDict):
                raise udfutils.AICException("UNKNOWN_METRIC", {'scorer': selectedMetric})

        dataFrame = dataFrame.append(rowDict, ignore_index=True)

        # results_summary is a dataframe
        return udfutils.ReturnDataFrame(dataFrame, a['display_column_headers'], False, a['transpose'])


