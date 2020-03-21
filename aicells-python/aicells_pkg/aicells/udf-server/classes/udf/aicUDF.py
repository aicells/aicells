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

import glob
import pathlib
import os

class aicUDF(AICFunction):
    def Run(self, arguments):
        a = self.ProcessArguments(arguments, 'parameters')

        filesDir = str(pathlib.Path(__file__).parent.absolute())

        directories = [
            # 'udf-1-draft',
            'udf-2-alpha',
            'udf-3-beta',
            'udf-4-production',
        ]

        fileList = []
        for d in directories:
            fileList = fileList + glob.glob(filesDir + f"\\..\\..\\yml\\{d}\\*.yml")

        UDFList = []
        for f in fileList:
            UDFList.append(os.path.basename(f).replace('.yml', ''))

        UDFList.sort()

        return udfutils.Transpose1DList(UDFList)
