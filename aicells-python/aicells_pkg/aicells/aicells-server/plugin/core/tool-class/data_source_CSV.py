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

import os
import pandas
import zlib

class data_source_CSV(AICFunction):
    def CRC32File(self, file):
        crc32 = 0
        for line in open(file, "rb"):
            crc32 = zlib.crc32(line, crc32)

        # To generate the same numeric value across all Python versions and platforms, use crc32(data) & 0xffffffff.
        crc32 = crc32 & 0xFFFFFFFF
        return crc32

    def Write(self, inputData, workbookPath, a, fullParameterName):
        # TODO error handling

        if a['read_only']:
            raise AICException.AICException("DATA_SOURCE_READ_ONLY", {})

        kwargs = a.copy()

        file = os.path.join(os.path.dirname(workbookPath), kwargs['file'])
        del kwargs['file']
        del kwargs['data_source']
        del kwargs['selected_columns']
        del kwargs['read_only']
        del kwargs['update_on_referenced_hash_change']
        del kwargs['hash']

        # these only used at reading
        del kwargs['skiprows']
        del kwargs['nrows']
        del kwargs['thousands']
        del kwargs['comment']
        del kwargs['na_values']

        kwargs['sep'] = kwargs['delimiter']
        del kwargs['delimiter']

        # treat empty cell as empty string
        if kwargs['na_rep'] is None:
            kwargs['na_rep'] = ''

        # validate column names
        if not (a['selected_columns'] is None):
            for column in a['selected_columns']:
                if not (isinstance(column, str) or isinstance(column, int) or isinstance(column, float)):
                    raise AICException.AICEParameterError("PARAMETER_INVALID_COLUMN",
                                                      {"parameterName": fullParameterName + '.selected_columns'})
                if not (column in inputData.columns):
                    raise AICException.AICEParameterError("PARAMETER_UNKNOWN_COLUMN",
                                                      {"parameterName": fullParameterName + '.selected_columns', "columnName": column})

        if a['selected_columns'] is None:
            selectedData = inputData
        else:
            selectedData = inputData[a['selected_columns']]

        kwargs['index'] = False
        #kwargs['header'] = False
        kwargs['path_or_buf'] = file
        selectedData.to_csv(**kwargs)

        crc32 = self.CRC32File(file)
        return "{" + f"{crc32:08x}" + "} " + os.path.basename(file)


    def Read(self, workbookPath, a, fullParameterName):
        # TODO error handling
        kwargs = a.copy()

        file = os.path.join(os.path.dirname(workbookPath), kwargs['file'])
        del kwargs['file']
        del kwargs['data_source']
        del kwargs['selected_columns']
        del kwargs['read_only']
        del kwargs['update_on_referenced_hash_change']
        del kwargs['hash']

        # these only used at writing
        del kwargs['float_format']
        del kwargs['date_format']
        del kwargs['na_rep']

        if kwargs['header']:
            kwargs['header'] = 0
        else:
            kwargs['header'] = None

        if kwargs['delimiter'] == "\\t":
            kwargs['delimiter'] = "\t"

        kwargs['filepath_or_buffer'] = file
        inputData = pandas.read_csv(**kwargs)

        # validate column names
        if not (a['selected_columns'] is None):
            for column in a['selected_columns']:
                if not (isinstance(column, str) or isinstance(column, int) or isinstance(column, float)):
                    raise AICException.AICEParameterError("PARAMETER_INVALID_COLUMN",
                                                      {"parameterName": fullParameterName + '.selected_columns'})
                if not (column in inputData.columns):
                    raise AICException.AICEParameterError("PARAMETER_UNKNOWN_COLUMN",
                                                      {"parameterName": fullParameterName + '.selected_columns', "columnName": column})

        if a['selected_columns'] is None:
            selectedData = inputData
        else:
            selectedData = inputData[a['selected_columns']]


        return selectedData
