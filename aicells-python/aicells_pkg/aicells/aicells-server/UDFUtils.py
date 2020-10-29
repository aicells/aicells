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


import numpy
import pandas
from . import AICException

# xlwings default behaviour:
# two Excel rows = [[1,2,3,4,5,6], [1,2,3,4,5,6]]
# one Excel column = [1,2,3,4,5,6]

def ReturnDataFrame(df, columnHeader, rowHeader, transpose=False):
    
    if rowHeader:
        df = df.rename_axis('').reset_index()

    columns = [df.columns.tolist()]
    values = df.values.tolist()

    if columnHeader:
        l = columns + values
    else:
        l = values

    # replace empty strings with None
    if df.isnull().any().any(): # NaN in numeric arrays, None or NaN in object arrays, NaT in datetimelike
        for idx1, v1 in enumerate(l):
            for idx2, v2 in enumerate(l[idx1]):
                if pandas.isnull(v2):
                    l[idx1][idx2] = ""

    if transpose:
        return Transpose2DList(l)
    else:
        return l

def ReturnSeries(series, columnHeader=False, transpose=False):
    if columnHeader:
        l = [series.index.values.tolist(), series.tolist()]
    else:
        l = [series.tolist()]

    # replace empty strings with None
    if series.isnull().any():
        for idx1, v1 in enumerate(l):
            if pandas.isnull(v1):
                l[idx1] = ""

    if transpose:
        return Transpose2DList(l)
    else:
        return l

def ReturnNumpyArray(npArray, transpose=False):
    if npArray.ndim == 1:
        npArray = numpy.reshape(npArray, (-1, 1))

    if npArray.ndim == 2:
        if transpose:
            l = npArray.T.tolist()
        else:
            l = npArray.tolist()

        if pandas.isnull(npArray):
            # replace empty strings with None
            for idx1, v1 in enumerate(l):
                for idx2, v2 in enumerate(l[idx1]):
                    if pandas.isnull(v2):
                        l[idx1][idx2] = ""

        return l

    raise AICException.AICException("TOO_MANY_ARRAY_DIMENSION")

def ReturnList(l, transpose=False):
    is2d = False
    if len(l) > 0:
        if isinstance(l[0], list):
            is2d = True

    if not is2d:
        l = [l]
        transpose = not transpose

    # replace empty strings with None
    for idx1, v1 in enumerate(l):
        for idx2, v2 in enumerate(l[idx1]):
            if pandas.isnull(v2):
                l[idx1][idx2] = ""

    if transpose:
        return Transpose2DList(l)
    else:
        return l


def Transpose2DList(l):
    # ret = numpy.array(l).T.tolist() # data type problems
    ret = list(map(list, zip(*l)))
    return ret

def Transpose1DList(l):
    return Transpose2DList([l])

