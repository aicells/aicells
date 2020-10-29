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
from .... import UDFUtils

import seaborn
import matplotlib.pyplot as plt
import matplotlib
#from matplotlib import rcParams
import matplotlib.colors
import sys
from io import BytesIO
import os

import random

def SetSeabornStyle():

    matplotlib.rcParams['figure.dpi'] = 300
    matplotlib.rcParams['figure.figsize'] = 5, 5

    cBlack = '#080808'
    cWhite = '#F0F0F0'

    plt.rcParams['savefig.facecolor'] = cBlack

    customStyle = {
        'aicells_1': 11,
        # "xtick.major.size": 10, "ytick.major.size": 10,
        'xtick.bottom': True,
        'xtick.top': False,
        'ytick.left': True,
        'ytick.right': False,

        'patch.edgecolor': cBlack,
        # 'patch.force_edgecolor': False,

        'axes.facecolor': cBlack,
        'figure.facecolor': cBlack,
        # 'legend.frameon': True,
        'axes.labelcolor': cWhite,
        'axes.edgecolor': cWhite,
        'grid.color': '#505050',
        'text.color': cWhite,
        'xtick.color': cWhite,
        'ytick.color': cWhite}

    seaborn.set_style("darkgrid", rc=customStyle)

    matplotlib.rcParams['figure.figsize'] = 5, 5

    # Office 365 Excel Color Palette, Accent 1-6
    palette = ["#4472C4", "#ED7D31", "#FFC000", "#A5A5A5", "#5B9BD5", "#70AD47"]

    seaborn.set_palette(palette)


    cmBlueToWhite = matplotlib.colors.LinearSegmentedColormap.from_list("n", [palette[0], cWhite])
    cmGreenToWhite = matplotlib.colors.LinearSegmentedColormap.from_list("n", [palette[5], cWhite])


def getSVG():
    # sio = BytesIO()
    # plt.savefig(sio, format='svg', bbox_inches="tight")
    # svgData = sio.getvalue()
    # sio.close()
    # plt.close()
    # return svgData.decode("utf-8")

    dir = os.path.dirname(os.path.realpath(__file__))
    svgFile = dir + '\\..\\..\\..\\..\\..\\..\\..\\aicells-temp\\output.svg'

    # 1 SVG point = 1/72 inch
    plt.savefig(svgFile, format='svg') # , bbox_inches="tight"

    return svgFile

class aicSeabornPlot(AICFunction):

    def Run(self, arguments):
        a = self.ProcessArguments(arguments, 'parameters')

        try:
            tool = self.factory.CreateInstance('tool-class.'+a['aicells_tool_name'].replace('.', '_'))
        except Exception as e:
            raise AICException.AICException("UNKNOWN_TOOL")

        tool_a = tool.ProcessArguments(arguments, 'parameters.tool_parameters')

        a['data'] = a['input_data']
        picture_name = a['picture_name']
        title = a['title']
        title_height_ratio = a['title_height_ratio']

        del a['aicells_tool_name']
        del a['input_data']
        del a['parameters']
        del a['picture_name']
        del a['tool_parameters']
        del a['title']
        del a['title_height_ratio']

        SetSeabornStyle()

        tips = seaborn.load_dataset("tips")

        # tips, col = "time", row = "sex", height = 5, aspect = 1, legend_out = True
        g = seaborn.FacetGrid(**a)

        # g.map_dataframe(seaborn.scatterplot, x="total_bill", y="tip", hue="day", linewidth=0)
        g.map_dataframe(**tool.GetMapArgs(tool_a))

        if title != '':
            g.fig.subplots_adjust(top=1 - title_height_ratio)
            g.fig.suptitle(title + " " + str(random.randint(1000, 9999)), fontsize=16)

        g.add_legend()

        return [picture_name, getSVG()]