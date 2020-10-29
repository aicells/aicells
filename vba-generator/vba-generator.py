import yaml
import os
import pathlib
import glob
import xml.sax.saxutils

def LimitString255(x):
    if len(x) > 255:
        x = x[0:255]
        print(f"String was limited to 255 characters: {x}")
    return x

intellisenseXML = "<IntelliSense xmlns=\"http://schemas.excel-dna.net/intellisense/1.0\">\n  <FunctionInfo>\n"

generatedCode = ""
processedUDFs = []

dir = os.path.dirname(os.path.realpath(__file__))
yamlFile = dir + '\\..\\aicells-config.yml'
with open(yamlFile) as file:
    aicConfig = yaml.load(file, Loader=yaml.FullLoader)

# directories = [
#     'core',
#     'correlation',
#     'random',
#     'supervised-learning',
#     'seaborn'
# ]

directories = aicConfig['plugins']

filesDir = str(pathlib.Path(__file__).parent.absolute())

fileList = []
for d in directories:
    fileList = fileList + glob.glob(filesDir + f"\\..\\aicells-python\\aicells_pkg\\aicells\\aicells-server\\plugin\\{d}\\function-yml\\*.yml")

for yamlFile in fileList:
    className = os.path.basename(yamlFile).replace('.yml', '')


    with open(yamlFile) as file:
        y = yaml.load(file, Loader=yaml.FullLoader)

    if not ('udf' in y['tag']):
        print(f"{className} [macro only]")
        continue

    print(f"{className}")

    typeMap = {
        "data_source": "range",
        "string": "String",
        "set": "String",
        "parameters": "range",
        "float": "Double",
        "boolean": "Boolean",
        "integer": "Long",
        "list": "range",
        "dataframe": "range",
        "series": "range",
        "Null": "Null",
        "False": "False"
    }

    generatedDocsCode = ""

    signature = []
    rangeValidator = ""
    UDFArguments = []
    argumentDescriptions = []

    if not 'UDFDescription' in y:
        y['UDFDescription'] = ''
    if y['UDFDescription'] is None:
        y['UDFDescription'] = ''

    intellisenseXML += '   <Function Name="'+y['pythonClassName']+'" Description='+xml.sax.saxutils.quoteattr(y['UDFDescription'])+' HelpTopic="">' + "\n"

    processedUDFs.append(y['pythonClassName'])
    if (y['parameters'] is None):
        y['parameters'] = []

    for p in y['parameters']:
        if "UDFParameter" in p:
            if p["UDFParameter"]:
                if not 'description' in p:
                    p['description'] = ''
                if p['description'] is None:
                    p['description'] = ''
                argumentDescriptions.append(p['description'])

                intellisenseXML += '      <Argument Name="' + p["parameterName"] + '" Description=' + xml.sax.saxutils.quoteattr(p["description"]) + ' />' + "\n"

                UDFArguments.append('        Array("' + p["parameterName"] + '", '+p["parameterName"] + ")")
                s = ""
                # if "default" in p:
                s += "Optional "
                s += p["parameterName"]
                # if "default" in p:
                s += " = Null"
                signature.append(s)

                isRangeType = False
                for t in p["type"]:
                    if t is None:
                        t = 'None'
                    if isinstance(t, bool):
                        if t == False:
                            t = 'False'
                    if typeMap[t] == 'range':
                        isRangeType = True

                if isRangeType:
                    if p["type"][0] == 'parameters':
                        # parameters (range type)
                        namespace = 'parameters'
                        if p["parameterName"] != 'parameters':
                            namespace = 'parameters.' + p["parameterName"]
                        rangeValidator += """
    If TypeOf {parameterName} Is range Then
        'If HasRangeErrors({parameterName}) Then GoTo valueError
        If ProcessParameterRanges2(pb, {parameterName}, "{namespace}") = False Then GoTo valueError
        {parameterName} = "@AICELLS-RANGE@"
    End If
""".format(parameterName=p["parameterName"], namespace=namespace)
                    else:
                        # range type, but not parameters
                        rangeValidator += """
    If TypeOf {parameterName} Is range Then
        'If HasRangeErrors({parameterName}) Then GoTo valueError
        'pb.StoreRange "parameters.{parameterName}", {parameterName}
        If ProcessParameterRanges2(pb, {parameterName}, "{namespace}") = False Then GoTo valueError        
        {parameterName} = "@AICELLS-RANGE@"
    ElseIf IsArray({parameterName}) Then
        pb.StoreArray "parameters.{parameterName}", {parameterName}
        scorers = "@AICELLS-RANGE@"
    End If
""".format(parameterName=p["parameterName"], namespace='parameters'+'.'+p["parameterName"])
                else:
                    # not range type
                    rangeValidator += """
    If TypeOf {parameterName} Is range Then
        If {parameterName}.Count <> 1 Then: GoTo valueError
        {parameterName} = {parameterName}.Value
    End If
""".format(parameterName=p["parameterName"])

    generatedCode += "' " + y['pythonClassName']+"\n"

    # DOCS {
    argumentDescriptionsBas = ''
    argumentDescriptionsBas2 = ''
    if len(argumentDescriptions) > 0:
        argumentDescriptionsBas = "ReDim argumentDescriptions(1 To {argument_count})".format(argument_count=len(argumentDescriptions))
        argumentDescriptionsBas2 = ', argumentDescriptions:=argumentDescriptions'

    generatedDocsCode = """
Private Sub {UDFName}_MacroOptions()
    Dim description As String
    Dim argumentDescriptions() As String
    {argumentDescriptionsBas}
    
    description = "{description}"
""".format(UDFName=y['pythonClassName'], description=LimitString255(y['description'].replace('"', '""')), argumentDescriptionsBas=argumentDescriptionsBas)

    for i in range(0, len(argumentDescriptions)):
        generatedDocsCode += f"    argumentDescriptions({i + 1}) = \"" + LimitString255(argumentDescriptions[i].replace('"', '""')) + "\"\n"

    generatedDocsCode += """
    Application.MacroOptions Macro:="{UDFName}", description:=Description, Category:="AICells"{argumentDescriptionsBas2}
End Sub

""".format(UDFName=y['pythonClassName'], argumentDescriptionsBas2=argumentDescriptionsBas2)

    generatedCode += generatedDocsCode

    # } DOCS

    signature = ", ".join(signature)

    generatedCode += "Function " + y['pythonClassName'] + "(" + signature + "):"

    volatile = "\n"
    if 'volatile' in y:
        if y['volatile']:
            volatile = "\n    Application.Volatile\n"

    generatedCode += """
    Dim PyReturn
    Dim PyParameters
    Dim pb As New PyParameterBuilder
    {volatile}
    If (IsFXWindowOpen()) Then
        '{pythonClassName} = "#FX"
        Exit Function
    End If
    
    pb.Init "{pythonClassName}"
        
    If TypeOf Application.Caller Is range Then
        DeleteErrorMessage Application.Caller
        On Error GoTo failed
    Else
        GoTo valueError
    End If
    
    Call LogUDFCall("{pythonClassName}", Application.Caller)    
""".format(pythonClassName=y["pythonClassName"], volatile=volatile)

    generatedCode += rangeValidator + "\n"

    if len(UDFArguments) != 0:
        generatedCode += "    pb.SetUdfArguments (Array( _\n"
        generatedCode += "        Array(\"_workbook_path\", Application.Caller.Worksheet.Parent.FullName), _\n"
        generatedCode += ", _\n".join(UDFArguments)
        generatedCode += "))"

    generatedCode += """
    PyReturn = Py.CallUDF("aicells-server", "aicUDFRunner", pb.GetParameterArray(), ThisWorkbook, Application.Caller)
    
    If CheckIfError(PyReturn) Then
        Call ShowErrors(PyReturn, Application.Caller)
        {UDFName} = CVErr(xlErrCalc)
    Else
        {UDFName} = PyReturn
    End If
    
    Exit Function
failed:
    {UDFName} = Err.Description
    Exit Function
valueError:
    {UDFName} = CVErr(xlErrValue)
    Exit Function
End Function

""".format(pythonClassName= y["pythonClassName"], UDFName = y['pythonClassName'])

    intellisenseXML += '    </Function>' + "\n"
# } for

#workbookOpen = "Private Sub Workbook_Open()\n"
workbookOpen = "Public Sub SetMacroOptions()\n"

for udf_name in processedUDFs:
    workbookOpen += f"    Call {udf_name}_MacroOptions\n"

workbookOpen += "End Sub\n\n"

generatedCode += workbookOpen

# print(generatedCode)

f = open("../aicells-excel-add-in/bas/aicells_udfs.bas", "w")
f.write(generatedCode)
f.close()

intellisenseXML += "  </FunctionInfo>\n</IntelliSense>\n"

f = open("../aicells-excel-add-in/aicells.intellisense.xml", "w")
f.write(intellisenseXML)
f.close()

print('Done.')