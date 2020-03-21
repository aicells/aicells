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

class aicUDFParameterDefaultValues(AICFunction):
    def Run(self, arguments):
        a = self.ProcessArguments(arguments, 'parameters')

        try:
            c = self.factory.CreateInstance("udf." + a['AIcells_UDF_name'])
        except Exception as e:
            raise udfutils.AICException("UNKNOWN_FUNCTION")
        else:
            return udfutils.Transpose1DList(c.GetParameterDefaultList())
