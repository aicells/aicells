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

import sklearn
import sklearn.linear_model

from ....AICFunction import AICFunction
from .... import AICException
from .... import UDFUtils

class sklearn_Ridge(AICFunction):
    def GetModel(self,
            model_parameters
            ):

        return sklearn.linear_model.Ridge(**model_parameters)


