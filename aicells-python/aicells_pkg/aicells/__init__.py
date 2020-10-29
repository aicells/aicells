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

import psutil
import xlwings
import os
import sys

__version__ = '0.0.1'

def IsProcessAlreadyRunning(script):
    for process in psutil.process_iter():
        if process.name().startswith('python'):
            if len(process.cmdline()) > 1:
                if script in process.cmdline()[1] and process.pid != os.getpid():
                    return True
    return False

def StartCOMServer(clsid):
    if not IsProcessAlreadyRunning('aicells-server.py'):
        sys.path.insert(0, os.path.dirname(os.path.realpath(__file__)))

        xlwings.serve(clsid=clsid)
    else:
        print('aicells-server.py is already running!')

if __name__ == "__main__":
    StartCOMServer()

