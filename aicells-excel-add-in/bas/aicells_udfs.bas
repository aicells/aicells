' aicCountEmptyCells

Private Sub aicCountEmptyCells_MacroOptions()
    Dim description As String
    Dim argumentDescriptions() As String
    ReDim argumentDescriptions(1 To 2)
    
    description = "Counts the empty cells in a range."
    argumentDescriptions(1) = "is a 2 dimensional list of parameter(s). The list contains key-value pairs."
    argumentDescriptions(2) = "is a table or range with or without header."

    Application.MacroOptions Macro:="aicCountEmptyCells", description:=Description, Category:="AICells", argumentDescriptions:=argumentDescriptions
End Sub

Function aicCountEmptyCells(Optional parameters = Null, Optional input_data = Null):
    Dim PyReturn
    Dim PyParameters
    Dim pb As New PyParameterBuilder
    

    If (IsFXWindowOpen()) Then
        'aicCountEmptyCells = "#FX"
        Exit Function
    End If
    
    pb.Init "aicCountEmptyCells"
        
    If TypeOf Application.Caller Is range Then
        DeleteErrorMessage Application.Caller
        On Error GoTo failed
    Else
        GoTo valueError
    End If
    
    Call LogUDFCall("aicCountEmptyCells", Application.Caller)    

    If TypeOf parameters Is range Then
        'If HasRangeErrors(parameters) Then GoTo valueError
        If ProcessParameterRanges2(pb, parameters, "parameters") = False Then GoTo valueError
        parameters = "@AICELLS-RANGE@"
    End If

    If TypeOf input_data Is range Then
        'If HasRangeErrors(input_data) Then GoTo valueError
        'pb.StoreRange "parameters.input_data", input_data
        If ProcessParameterRanges2(pb, input_data, "parameters.input_data") = False Then GoTo valueError        
        input_data = "@AICELLS-RANGE@"
    ElseIf IsArray(input_data) Then
        pb.StoreArray "parameters.input_data", input_data
        scorers = "@AICELLS-RANGE@"
    End If

    pb.SetUdfArguments (Array( _
        Array("_workbook_path", Application.Caller.Worksheet.Parent.FullName), _
        Array("parameters", parameters), _
        Array("input_data", input_data)))
    PyReturn = Py.CallUDF("aicells-server", "aicUDFRunner", pb.GetParameterArray(), ThisWorkbook, Application.Caller)
    
    If CheckIfError(PyReturn) Then
        Call ShowErrors(PyReturn, Application.Caller)
        aicCountEmptyCells = CVErr(xlErrCalc)
    Else
        aicCountEmptyCells = PyReturn
    End If
    
    Exit Function
failed:
    aicCountEmptyCells = Err.Description
    Exit Function
valueError:
    aicCountEmptyCells = CVErr(xlErrValue)
    Exit Function
End Function

' aicDataCopy

Private Sub aicDataCopy_MacroOptions()
    Dim description As String
    Dim argumentDescriptions() As String
    ReDim argumentDescriptions(1 To 3)
    
    description = ""
    argumentDescriptions(1) = "is a 2 dimensional list of parameter(s). The list contains key-value pairs."
    argumentDescriptions(2) = "is a table or range with header."
    argumentDescriptions(3) = "is a logical value: to transpose the results =TRUE"

    Application.MacroOptions Macro:="aicDataCopy", description:=Description, Category:="AICells", argumentDescriptions:=argumentDescriptions
End Sub

Function aicDataCopy(Optional parameters = Null, Optional input_data = Null, Optional transpose = Null):
    Dim PyReturn
    Dim PyParameters
    Dim pb As New PyParameterBuilder
    

    If (IsFXWindowOpen()) Then
        'aicDataCopy = "#FX"
        Exit Function
    End If
    
    pb.Init "aicDataCopy"
        
    If TypeOf Application.Caller Is range Then
        DeleteErrorMessage Application.Caller
        On Error GoTo failed
    Else
        GoTo valueError
    End If
    
    Call LogUDFCall("aicDataCopy", Application.Caller)    

    If TypeOf parameters Is range Then
        'If HasRangeErrors(parameters) Then GoTo valueError
        If ProcessParameterRanges2(pb, parameters, "parameters") = False Then GoTo valueError
        parameters = "@AICELLS-RANGE@"
    End If

    If TypeOf input_data Is range Then
        'If HasRangeErrors(input_data) Then GoTo valueError
        'pb.StoreRange "parameters.input_data", input_data
        If ProcessParameterRanges2(pb, input_data, "parameters.input_data") = False Then GoTo valueError        
        input_data = "@AICELLS-RANGE@"
    ElseIf IsArray(input_data) Then
        pb.StoreArray "parameters.input_data", input_data
        scorers = "@AICELLS-RANGE@"
    End If

    If TypeOf transpose Is range Then
        If transpose.Count <> 1 Then: GoTo valueError
        transpose = transpose.Value
    End If

    pb.SetUdfArguments (Array( _
        Array("_workbook_path", Application.Caller.Worksheet.Parent.FullName), _
        Array("parameters", parameters), _
        Array("input_data", input_data), _
        Array("transpose", transpose)))
    PyReturn = Py.CallUDF("aicells-server", "aicUDFRunner", pb.GetParameterArray(), ThisWorkbook, Application.Caller)
    
    If CheckIfError(PyReturn) Then
        Call ShowErrors(PyReturn, Application.Caller)
        aicDataCopy = CVErr(xlErrCalc)
    Else
        aicDataCopy = PyReturn
    End If
    
    Exit Function
failed:
    aicDataCopy = Err.Description
    Exit Function
valueError:
    aicDataCopy = CVErr(xlErrValue)
    Exit Function
End Function

' aicDescribe

Private Sub aicDescribe_MacroOptions()
    Dim description As String
    Dim argumentDescriptions() As String
    ReDim argumentDescriptions(1 To 10)
    
    description = "Returns a statistics for the selected columns."
    argumentDescriptions(1) = "is a 2 dimensional list of parameter(s). The list contains key-value pairs."
    argumentDescriptions(2) = "is a table or range with header."
    argumentDescriptions(3) = "is a list of selected column header names or empty. For select all columns, use empty. You can select adjacent cells only."
    argumentDescriptions(4) = "To display all statistics set this to empty, or range/array of selected statistics"
    argumentDescriptions(5) = "is a logical value: to display column headers =TRUE or empty cell; to hide column headers = FALSE."
    argumentDescriptions(6) = "is a logical value: to display row headers =TRUE or empty cell; to hide row headers = FALSE."
    argumentDescriptions(7) = "is a logical value: to transpose the results =TRUE"
    argumentDescriptions(8) = "is a number to use as the first percentile or empty cell for default value = 0.25."
    argumentDescriptions(9) = "is a number to use as the secnd percentile or empty cell for default value = 0.5."
    argumentDescriptions(10) = "is a number to use as the third percentile or empty cell for default value = 0.75."

    Application.MacroOptions Macro:="aicDescribe", description:=Description, Category:="AICells", argumentDescriptions:=argumentDescriptions
End Sub

Function aicDescribe(Optional parameters = Null, Optional input_data = Null, Optional selected_columns = Null, Optional selected_statistics = Null, Optional display_column_headers = Null, Optional display_row_headers = Null, Optional transpose = Null, Optional percentile1 = Null, Optional percentile2 = Null, Optional percentile3 = Null):
    Dim PyReturn
    Dim PyParameters
    Dim pb As New PyParameterBuilder
    

    If (IsFXWindowOpen()) Then
        'aicDescribe = "#FX"
        Exit Function
    End If
    
    pb.Init "aicDescribe"
        
    If TypeOf Application.Caller Is range Then
        DeleteErrorMessage Application.Caller
        On Error GoTo failed
    Else
        GoTo valueError
    End If
    
    Call LogUDFCall("aicDescribe", Application.Caller)    

    If TypeOf parameters Is range Then
        'If HasRangeErrors(parameters) Then GoTo valueError
        If ProcessParameterRanges2(pb, parameters, "parameters") = False Then GoTo valueError
        parameters = "@AICELLS-RANGE@"
    End If

    If TypeOf input_data Is range Then
        'If HasRangeErrors(input_data) Then GoTo valueError
        'pb.StoreRange "parameters.input_data", input_data
        If ProcessParameterRanges2(pb, input_data, "parameters.input_data") = False Then GoTo valueError        
        input_data = "@AICELLS-RANGE@"
    ElseIf IsArray(input_data) Then
        pb.StoreArray "parameters.input_data", input_data
        scorers = "@AICELLS-RANGE@"
    End If

    If TypeOf selected_columns Is range Then
        'If HasRangeErrors(selected_columns) Then GoTo valueError
        'pb.StoreRange "parameters.selected_columns", selected_columns
        If ProcessParameterRanges2(pb, selected_columns, "parameters.selected_columns") = False Then GoTo valueError        
        selected_columns = "@AICELLS-RANGE@"
    ElseIf IsArray(selected_columns) Then
        pb.StoreArray "parameters.selected_columns", selected_columns
        scorers = "@AICELLS-RANGE@"
    End If

    If TypeOf selected_statistics Is range Then
        'If HasRangeErrors(selected_statistics) Then GoTo valueError
        'pb.StoreRange "parameters.selected_statistics", selected_statistics
        If ProcessParameterRanges2(pb, selected_statistics, "parameters.selected_statistics") = False Then GoTo valueError        
        selected_statistics = "@AICELLS-RANGE@"
    ElseIf IsArray(selected_statistics) Then
        pb.StoreArray "parameters.selected_statistics", selected_statistics
        scorers = "@AICELLS-RANGE@"
    End If

    If TypeOf display_column_headers Is range Then
        If display_column_headers.Count <> 1 Then: GoTo valueError
        display_column_headers = display_column_headers.Value
    End If

    If TypeOf display_row_headers Is range Then
        If display_row_headers.Count <> 1 Then: GoTo valueError
        display_row_headers = display_row_headers.Value
    End If

    If TypeOf transpose Is range Then
        If transpose.Count <> 1 Then: GoTo valueError
        transpose = transpose.Value
    End If

    If TypeOf percentile1 Is range Then
        If percentile1.Count <> 1 Then: GoTo valueError
        percentile1 = percentile1.Value
    End If

    If TypeOf percentile2 Is range Then
        If percentile2.Count <> 1 Then: GoTo valueError
        percentile2 = percentile2.Value
    End If

    If TypeOf percentile3 Is range Then
        If percentile3.Count <> 1 Then: GoTo valueError
        percentile3 = percentile3.Value
    End If

    pb.SetUdfArguments (Array( _
        Array("_workbook_path", Application.Caller.Worksheet.Parent.FullName), _
        Array("parameters", parameters), _
        Array("input_data", input_data), _
        Array("selected_columns", selected_columns), _
        Array("selected_statistics", selected_statistics), _
        Array("display_column_headers", display_column_headers), _
        Array("display_row_headers", display_row_headers), _
        Array("transpose", transpose), _
        Array("percentile1", percentile1), _
        Array("percentile2", percentile2), _
        Array("percentile3", percentile3)))
    PyReturn = Py.CallUDF("aicells-server", "aicUDFRunner", pb.GetParameterArray(), ThisWorkbook, Application.Caller)
    
    If CheckIfError(PyReturn) Then
        Call ShowErrors(PyReturn, Application.Caller)
        aicDescribe = CVErr(xlErrCalc)
    Else
        aicDescribe = PyReturn
    End If
    
    Exit Function
failed:
    aicDescribe = Err.Description
    Exit Function
valueError:
    aicDescribe = CVErr(xlErrValue)
    Exit Function
End Function

' aicFillEmptyCells

Private Sub aicFillEmptyCells_MacroOptions()
    Dim description As String
    Dim argumentDescriptions() As String
    ReDim argumentDescriptions(1 To 8)
    
    description = "Fills the empty cells with value according the selected method."
    argumentDescriptions(1) = "is a 2 dimensional list of parameter(s). The list contains key-value pairs."
    argumentDescriptions(2) = "is a table or range with header."
    argumentDescriptions(3) = "is a list of selected column header names. For select all columns, leave it empty. You can select adjacent cells only."
    argumentDescriptions(4) = "is a logical value: to transpose the results =TRUE"
    argumentDescriptions(5) = "is a number or string: to fill empty cells (e.g. 0) when 'method' is Null."
    argumentDescriptions(6) = "Method to use for filling empty cells: 'value', 'pad', 'ffill', 'backfill', 'bfill', 'mean', 'median'."
    argumentDescriptions(7) = "when TRUE, function returns all columns from input_data, when FALSE, function returns only the filled columns"
    argumentDescriptions(8) = "if method is specified, this is the maximum number of consecutive NaN values to forward/backward fill. If method is not specified, this is the maximum number of entries along the entire axis where NaNs will be filled. Must be greater than 0 if not None."

    Application.MacroOptions Macro:="aicFillEmptyCells", description:=Description, Category:="AICells", argumentDescriptions:=argumentDescriptions
End Sub

Function aicFillEmptyCells(Optional parameters = Null, Optional input_data = Null, Optional selected_columns = Null, Optional transpose = Null, Optional value = Null, Optional method = Null, Optional return_all_columns = Null, Optional limit = Null):
    Dim PyReturn
    Dim PyParameters
    Dim pb As New PyParameterBuilder
    

    If (IsFXWindowOpen()) Then
        'aicFillEmptyCells = "#FX"
        Exit Function
    End If
    
    pb.Init "aicFillEmptyCells"
        
    If TypeOf Application.Caller Is range Then
        DeleteErrorMessage Application.Caller
        On Error GoTo failed
    Else
        GoTo valueError
    End If
    
    Call LogUDFCall("aicFillEmptyCells", Application.Caller)    

    If TypeOf parameters Is range Then
        'If HasRangeErrors(parameters) Then GoTo valueError
        If ProcessParameterRanges2(pb, parameters, "parameters") = False Then GoTo valueError
        parameters = "@AICELLS-RANGE@"
    End If

    If TypeOf input_data Is range Then
        'If HasRangeErrors(input_data) Then GoTo valueError
        'pb.StoreRange "parameters.input_data", input_data
        If ProcessParameterRanges2(pb, input_data, "parameters.input_data") = False Then GoTo valueError        
        input_data = "@AICELLS-RANGE@"
    ElseIf IsArray(input_data) Then
        pb.StoreArray "parameters.input_data", input_data
        scorers = "@AICELLS-RANGE@"
    End If

    If TypeOf selected_columns Is range Then
        'If HasRangeErrors(selected_columns) Then GoTo valueError
        'pb.StoreRange "parameters.selected_columns", selected_columns
        If ProcessParameterRanges2(pb, selected_columns, "parameters.selected_columns") = False Then GoTo valueError        
        selected_columns = "@AICELLS-RANGE@"
    ElseIf IsArray(selected_columns) Then
        pb.StoreArray "parameters.selected_columns", selected_columns
        scorers = "@AICELLS-RANGE@"
    End If

    If TypeOf transpose Is range Then
        If transpose.Count <> 1 Then: GoTo valueError
        transpose = transpose.Value
    End If

    If TypeOf value Is range Then
        If value.Count <> 1 Then: GoTo valueError
        value = value.Value
    End If

    If TypeOf method Is range Then
        If method.Count <> 1 Then: GoTo valueError
        method = method.Value
    End If

    If TypeOf return_all_columns Is range Then
        If return_all_columns.Count <> 1 Then: GoTo valueError
        return_all_columns = return_all_columns.Value
    End If

    If TypeOf limit Is range Then
        If limit.Count <> 1 Then: GoTo valueError
        limit = limit.Value
    End If

    pb.SetUdfArguments (Array( _
        Array("_workbook_path", Application.Caller.Worksheet.Parent.FullName), _
        Array("parameters", parameters), _
        Array("input_data", input_data), _
        Array("selected_columns", selected_columns), _
        Array("transpose", transpose), _
        Array("value", value), _
        Array("method", method), _
        Array("return_all_columns", return_all_columns), _
        Array("limit", limit)))
    PyReturn = Py.CallUDF("aicells-server", "aicUDFRunner", pb.GetParameterArray(), ThisWorkbook, Application.Caller)
    
    If CheckIfError(PyReturn) Then
        Call ShowErrors(PyReturn, Application.Caller)
        aicFillEmptyCells = CVErr(xlErrCalc)
    Else
        aicFillEmptyCells = PyReturn
    End If
    
    Exit Function
failed:
    aicFillEmptyCells = Err.Description
    Exit Function
valueError:
    aicFillEmptyCells = CVErr(xlErrValue)
    Exit Function
End Function

' aicFunctionDescription

Private Sub aicFunctionDescription_MacroOptions()
    Dim description As String
    Dim argumentDescriptions() As String
    ReDim argumentDescriptions(1 To 1)
    
    description = "Returns the description for a specific AIcells Excel function (UDF)."
    argumentDescriptions(1) = "is the name of an AIcells Excel function (UDF). You can get the list of AIcells function names with the aicUDF() function."

    Application.MacroOptions Macro:="aicFunctionDescription", description:=Description, Category:="AICells", argumentDescriptions:=argumentDescriptions
End Sub

Function aicFunctionDescription(Optional AIcells_UDF_name = Null):
    Dim PyReturn
    Dim PyParameters
    Dim pb As New PyParameterBuilder
    

    If (IsFXWindowOpen()) Then
        'aicFunctionDescription = "#FX"
        Exit Function
    End If
    
    pb.Init "aicFunctionDescription"
        
    If TypeOf Application.Caller Is range Then
        DeleteErrorMessage Application.Caller
        On Error GoTo failed
    Else
        GoTo valueError
    End If
    
    Call LogUDFCall("aicFunctionDescription", Application.Caller)    

    If TypeOf AIcells_UDF_name Is range Then
        If AIcells_UDF_name.Count <> 1 Then: GoTo valueError
        AIcells_UDF_name = AIcells_UDF_name.Value
    End If

    pb.SetUdfArguments (Array( _
        Array("_workbook_path", Application.Caller.Worksheet.Parent.FullName), _
        Array("AIcells_UDF_name", AIcells_UDF_name)))
    PyReturn = Py.CallUDF("aicells-server", "aicUDFRunner", pb.GetParameterArray(), ThisWorkbook, Application.Caller)
    
    If CheckIfError(PyReturn) Then
        Call ShowErrors(PyReturn, Application.Caller)
        aicFunctionDescription = CVErr(xlErrCalc)
    Else
        aicFunctionDescription = PyReturn
    End If
    
    Exit Function
failed:
    aicFunctionDescription = Err.Description
    Exit Function
valueError:
    aicFunctionDescription = CVErr(xlErrValue)
    Exit Function
End Function

' aicFunctionList

Private Sub aicFunctionList_MacroOptions()
    Dim description As String
    Dim argumentDescriptions() As String
    
    
    description = "Lists the available AIcells Excel functions (UDF). Enter the following formula in a cell: ""=aic.UDF()""."

    Application.MacroOptions Macro:="aicFunctionList", description:=Description, Category:="AICells"
End Sub

Function aicFunctionList():
    Dim PyReturn
    Dim PyParameters
    Dim pb As New PyParameterBuilder
    

    If (IsFXWindowOpen()) Then
        'aicFunctionList = "#FX"
        Exit Function
    End If
    
    pb.Init "aicFunctionList"
        
    If TypeOf Application.Caller Is range Then
        DeleteErrorMessage Application.Caller
        On Error GoTo failed
    Else
        GoTo valueError
    End If
    
    Call LogUDFCall("aicFunctionList", Application.Caller)    


    PyReturn = Py.CallUDF("aicells-server", "aicUDFRunner", pb.GetParameterArray(), ThisWorkbook, Application.Caller)
    
    If CheckIfError(PyReturn) Then
        Call ShowErrors(PyReturn, Application.Caller)
        aicFunctionList = CVErr(xlErrCalc)
    Else
        aicFunctionList = PyReturn
    End If
    
    Exit Function
failed:
    aicFunctionList = Err.Description
    Exit Function
valueError:
    aicFunctionList = CVErr(xlErrValue)
    Exit Function
End Function

' aicFunctionParameterDataTypes

Private Sub aicFunctionParameterDataTypes_MacroOptions()
    Dim description As String
    Dim argumentDescriptions() As String
    ReDim argumentDescriptions(1 To 3)
    
    description = "Lists the parameter data types for a specific AIcells Excel function (UDF)."
    argumentDescriptions(1) = "is the name of an AIcells Excel function (UDF). You can get the list of AIcells function names with the aicUDF() function."
    argumentDescriptions(2) = "is a logical value: to show the 'function' and 'output' fields =TRUE"
    argumentDescriptions(3) = "is a logical value: to transpose the results =TRUE"

    Application.MacroOptions Macro:="aicFunctionParameterDataTypes", description:=Description, Category:="AICells", argumentDescriptions:=argumentDescriptions
End Sub

Function aicFunctionParameterDataTypes(Optional AIcells_UDF_name = Null, Optional show_function_and_output = Null, Optional transpose = Null):
    Dim PyReturn
    Dim PyParameters
    Dim pb As New PyParameterBuilder
    

    If (IsFXWindowOpen()) Then
        'aicFunctionParameterDataTypes = "#FX"
        Exit Function
    End If
    
    pb.Init "aicFunctionParameterDataTypes"
        
    If TypeOf Application.Caller Is range Then
        DeleteErrorMessage Application.Caller
        On Error GoTo failed
    Else
        GoTo valueError
    End If
    
    Call LogUDFCall("aicFunctionParameterDataTypes", Application.Caller)    

    If TypeOf AIcells_UDF_name Is range Then
        If AIcells_UDF_name.Count <> 1 Then: GoTo valueError
        AIcells_UDF_name = AIcells_UDF_name.Value
    End If

    If TypeOf show_function_and_output Is range Then
        If show_function_and_output.Count <> 1 Then: GoTo valueError
        show_function_and_output = show_function_and_output.Value
    End If

    If TypeOf transpose Is range Then
        If transpose.Count <> 1 Then: GoTo valueError
        transpose = transpose.Value
    End If

    pb.SetUdfArguments (Array( _
        Array("_workbook_path", Application.Caller.Worksheet.Parent.FullName), _
        Array("AIcells_UDF_name", AIcells_UDF_name), _
        Array("show_function_and_output", show_function_and_output), _
        Array("transpose", transpose)))
    PyReturn = Py.CallUDF("aicells-server", "aicUDFRunner", pb.GetParameterArray(), ThisWorkbook, Application.Caller)
    
    If CheckIfError(PyReturn) Then
        Call ShowErrors(PyReturn, Application.Caller)
        aicFunctionParameterDataTypes = CVErr(xlErrCalc)
    Else
        aicFunctionParameterDataTypes = PyReturn
    End If
    
    Exit Function
failed:
    aicFunctionParameterDataTypes = Err.Description
    Exit Function
valueError:
    aicFunctionParameterDataTypes = CVErr(xlErrValue)
    Exit Function
End Function

' aicFunctionParameterDefaultValues

Private Sub aicFunctionParameterDefaultValues_MacroOptions()
    Dim description As String
    Dim argumentDescriptions() As String
    ReDim argumentDescriptions(1 To 3)
    
    description = "Lists the parameter default values for a specific AIcells Excel function (UDF)."
    argumentDescriptions(1) = "is the name of an AIcells Excel function (UDF). You can get the list of AIcells function names with the aicUDF() function."
    argumentDescriptions(2) = "is a logical value: to show the 'function' and 'output' fields =TRUE"
    argumentDescriptions(3) = "is a logical value: to transpose the results =TRUE"

    Application.MacroOptions Macro:="aicFunctionParameterDefaultValues", description:=Description, Category:="AICells", argumentDescriptions:=argumentDescriptions
End Sub

Function aicFunctionParameterDefaultValues(Optional AIcells_UDF_name = Null, Optional show_function_and_output = Null, Optional transpose = Null):
    Dim PyReturn
    Dim PyParameters
    Dim pb As New PyParameterBuilder
    

    If (IsFXWindowOpen()) Then
        'aicFunctionParameterDefaultValues = "#FX"
        Exit Function
    End If
    
    pb.Init "aicFunctionParameterDefaultValues"
        
    If TypeOf Application.Caller Is range Then
        DeleteErrorMessage Application.Caller
        On Error GoTo failed
    Else
        GoTo valueError
    End If
    
    Call LogUDFCall("aicFunctionParameterDefaultValues", Application.Caller)    

    If TypeOf AIcells_UDF_name Is range Then
        If AIcells_UDF_name.Count <> 1 Then: GoTo valueError
        AIcells_UDF_name = AIcells_UDF_name.Value
    End If

    If TypeOf show_function_and_output Is range Then
        If show_function_and_output.Count <> 1 Then: GoTo valueError
        show_function_and_output = show_function_and_output.Value
    End If

    If TypeOf transpose Is range Then
        If transpose.Count <> 1 Then: GoTo valueError
        transpose = transpose.Value
    End If

    pb.SetUdfArguments (Array( _
        Array("_workbook_path", Application.Caller.Worksheet.Parent.FullName), _
        Array("AIcells_UDF_name", AIcells_UDF_name), _
        Array("show_function_and_output", show_function_and_output), _
        Array("transpose", transpose)))
    PyReturn = Py.CallUDF("aicells-server", "aicUDFRunner", pb.GetParameterArray(), ThisWorkbook, Application.Caller)
    
    If CheckIfError(PyReturn) Then
        Call ShowErrors(PyReturn, Application.Caller)
        aicFunctionParameterDefaultValues = CVErr(xlErrCalc)
    Else
        aicFunctionParameterDefaultValues = PyReturn
    End If
    
    Exit Function
failed:
    aicFunctionParameterDefaultValues = Err.Description
    Exit Function
valueError:
    aicFunctionParameterDefaultValues = CVErr(xlErrValue)
    Exit Function
End Function

' aicFunctionParameterDescriptions

Private Sub aicFunctionParameterDescriptions_MacroOptions()
    Dim description As String
    Dim argumentDescriptions() As String
    ReDim argumentDescriptions(1 To 3)
    
    description = "Lists the parameter descriptions for a specific AIcells Excel function (UDF)."
    argumentDescriptions(1) = "is the name of an AIcells Excel function (UDF). You can get the list of AIcells function names with the aicUDF() function."
    argumentDescriptions(2) = "is a logical value: to show the 'function' and 'output' fields =TRUE"
    argumentDescriptions(3) = "is a logical value: to transpose the results =TRUE"

    Application.MacroOptions Macro:="aicFunctionParameterDescriptions", description:=Description, Category:="AICells", argumentDescriptions:=argumentDescriptions
End Sub

Function aicFunctionParameterDescriptions(Optional AIcells_UDF_name = Null, Optional show_function_and_output = Null, Optional transpose = Null):
    Dim PyReturn
    Dim PyParameters
    Dim pb As New PyParameterBuilder
    

    If (IsFXWindowOpen()) Then
        'aicFunctionParameterDescriptions = "#FX"
        Exit Function
    End If
    
    pb.Init "aicFunctionParameterDescriptions"
        
    If TypeOf Application.Caller Is range Then
        DeleteErrorMessage Application.Caller
        On Error GoTo failed
    Else
        GoTo valueError
    End If
    
    Call LogUDFCall("aicFunctionParameterDescriptions", Application.Caller)    

    If TypeOf AIcells_UDF_name Is range Then
        If AIcells_UDF_name.Count <> 1 Then: GoTo valueError
        AIcells_UDF_name = AIcells_UDF_name.Value
    End If

    If TypeOf show_function_and_output Is range Then
        If show_function_and_output.Count <> 1 Then: GoTo valueError
        show_function_and_output = show_function_and_output.Value
    End If

    If TypeOf transpose Is range Then
        If transpose.Count <> 1 Then: GoTo valueError
        transpose = transpose.Value
    End If

    pb.SetUdfArguments (Array( _
        Array("_workbook_path", Application.Caller.Worksheet.Parent.FullName), _
        Array("AIcells_UDF_name", AIcells_UDF_name), _
        Array("show_function_and_output", show_function_and_output), _
        Array("transpose", transpose)))
    PyReturn = Py.CallUDF("aicells-server", "aicUDFRunner", pb.GetParameterArray(), ThisWorkbook, Application.Caller)
    
    If CheckIfError(PyReturn) Then
        Call ShowErrors(PyReturn, Application.Caller)
        aicFunctionParameterDescriptions = CVErr(xlErrCalc)
    Else
        aicFunctionParameterDescriptions = PyReturn
    End If
    
    Exit Function
failed:
    aicFunctionParameterDescriptions = Err.Description
    Exit Function
valueError:
    aicFunctionParameterDescriptions = CVErr(xlErrValue)
    Exit Function
End Function

' aicFunctionParameters

Private Sub aicFunctionParameters_MacroOptions()
    Dim description As String
    Dim argumentDescriptions() As String
    ReDim argumentDescriptions(1 To 3)
    
    description = "Lists the parameters for a specific AIcells Excel function (UDF)."
    argumentDescriptions(1) = "is the name of an AIcells Excel function (UDF). You can get the list of AIcells function names with the aicUDF() function."
    argumentDescriptions(2) = "is a logical value: to show the 'function' and 'output' fields =TRUE"
    argumentDescriptions(3) = "is a logical value: to transpose the results =TRUE"

    Application.MacroOptions Macro:="aicFunctionParameters", description:=Description, Category:="AICells", argumentDescriptions:=argumentDescriptions
End Sub

Function aicFunctionParameters(Optional AIcells_UDF_name = Null, Optional show_function_and_output = Null, Optional transpose = Null):
    Dim PyReturn
    Dim PyParameters
    Dim pb As New PyParameterBuilder
    

    If (IsFXWindowOpen()) Then
        'aicFunctionParameters = "#FX"
        Exit Function
    End If
    
    pb.Init "aicFunctionParameters"
        
    If TypeOf Application.Caller Is range Then
        DeleteErrorMessage Application.Caller
        On Error GoTo failed
    Else
        GoTo valueError
    End If
    
    Call LogUDFCall("aicFunctionParameters", Application.Caller)    

    If TypeOf AIcells_UDF_name Is range Then
        If AIcells_UDF_name.Count <> 1 Then: GoTo valueError
        AIcells_UDF_name = AIcells_UDF_name.Value
    End If

    If TypeOf show_function_and_output Is range Then
        If show_function_and_output.Count <> 1 Then: GoTo valueError
        show_function_and_output = show_function_and_output.Value
    End If

    If TypeOf transpose Is range Then
        If transpose.Count <> 1 Then: GoTo valueError
        transpose = transpose.Value
    End If

    pb.SetUdfArguments (Array( _
        Array("_workbook_path", Application.Caller.Worksheet.Parent.FullName), _
        Array("AIcells_UDF_name", AIcells_UDF_name), _
        Array("show_function_and_output", show_function_and_output), _
        Array("transpose", transpose)))
    PyReturn = Py.CallUDF("aicells-server", "aicUDFRunner", pb.GetParameterArray(), ThisWorkbook, Application.Caller)
    
    If CheckIfError(PyReturn) Then
        Call ShowErrors(PyReturn, Application.Caller)
        aicFunctionParameters = CVErr(xlErrCalc)
    Else
        aicFunctionParameters = PyReturn
    End If
    
    Exit Function
failed:
    aicFunctionParameters = Err.Description
    Exit Function
valueError:
    aicFunctionParameters = CVErr(xlErrValue)
    Exit Function
End Function

' aicGetDummies

Private Sub aicGetDummies_MacroOptions()
    Dim description As String
    Dim argumentDescriptions() As String
    ReDim argumentDescriptions(1 To 6)
    
    description = "Converts categorical columns into dummy/indicator columns."
    argumentDescriptions(1) = "is a 2 dimensional list of parameter(s). The list contains key value pairs."
    argumentDescriptions(2) = "is a table or range with header."
    argumentDescriptions(3) = "is a list of selected column header names or empty. To use all columns = FALSE."
    argumentDescriptions(4) = "is a logical value: when TRUE, the function returns rows are not in ""selected_columns"" and the dummies of the selected columns"
    argumentDescriptions(5) = "is a logical value: to display column headers =TRUE or empty cell; to hide column headers = FALSE."
    argumentDescriptions(6) = "is a logical value: to transpose the results =TRUE"

    Application.MacroOptions Macro:="aicGetDummies", description:=Description, Category:="AICells", argumentDescriptions:=argumentDescriptions
End Sub

Function aicGetDummies(Optional parameters = Null, Optional input_data = Null, Optional selected_columns = Null, Optional full_table = Null, Optional column_header = Null, Optional transpose = Null):
    Dim PyReturn
    Dim PyParameters
    Dim pb As New PyParameterBuilder
    

    If (IsFXWindowOpen()) Then
        'aicGetDummies = "#FX"
        Exit Function
    End If
    
    pb.Init "aicGetDummies"
        
    If TypeOf Application.Caller Is range Then
        DeleteErrorMessage Application.Caller
        On Error GoTo failed
    Else
        GoTo valueError
    End If
    
    Call LogUDFCall("aicGetDummies", Application.Caller)    

    If TypeOf parameters Is range Then
        'If HasRangeErrors(parameters) Then GoTo valueError
        If ProcessParameterRanges2(pb, parameters, "parameters") = False Then GoTo valueError
        parameters = "@AICELLS-RANGE@"
    End If

    If TypeOf input_data Is range Then
        'If HasRangeErrors(input_data) Then GoTo valueError
        'pb.StoreRange "parameters.input_data", input_data
        If ProcessParameterRanges2(pb, input_data, "parameters.input_data") = False Then GoTo valueError        
        input_data = "@AICELLS-RANGE@"
    ElseIf IsArray(input_data) Then
        pb.StoreArray "parameters.input_data", input_data
        scorers = "@AICELLS-RANGE@"
    End If

    If TypeOf selected_columns Is range Then
        'If HasRangeErrors(selected_columns) Then GoTo valueError
        'pb.StoreRange "parameters.selected_columns", selected_columns
        If ProcessParameterRanges2(pb, selected_columns, "parameters.selected_columns") = False Then GoTo valueError        
        selected_columns = "@AICELLS-RANGE@"
    ElseIf IsArray(selected_columns) Then
        pb.StoreArray "parameters.selected_columns", selected_columns
        scorers = "@AICELLS-RANGE@"
    End If

    If TypeOf full_table Is range Then
        If full_table.Count <> 1 Then: GoTo valueError
        full_table = full_table.Value
    End If

    If TypeOf column_header Is range Then
        If column_header.Count <> 1 Then: GoTo valueError
        column_header = column_header.Value
    End If

    If TypeOf transpose Is range Then
        If transpose.Count <> 1 Then: GoTo valueError
        transpose = transpose.Value
    End If

    pb.SetUdfArguments (Array( _
        Array("_workbook_path", Application.Caller.Worksheet.Parent.FullName), _
        Array("parameters", parameters), _
        Array("input_data", input_data), _
        Array("selected_columns", selected_columns), _
        Array("full_table", full_table), _
        Array("column_header", column_header), _
        Array("transpose", transpose)))
    PyReturn = Py.CallUDF("aicells-server", "aicUDFRunner", pb.GetParameterArray(), ThisWorkbook, Application.Caller)
    
    If CheckIfError(PyReturn) Then
        Call ShowErrors(PyReturn, Application.Caller)
        aicGetDummies = CVErr(xlErrCalc)
    Else
        aicGetDummies = PyReturn
    End If
    
    Exit Function
failed:
    aicGetDummies = Err.Description
    Exit Function
valueError:
    aicGetDummies = CVErr(xlErrValue)
    Exit Function
End Function

' aicHelloWorld

Private Sub aicHelloWorld_MacroOptions()
    Dim description As String
    Dim argumentDescriptions() As String
    
    
    description = "Is a test function of Aicells Excel add-in. It returns a simple Hello World message. Enter the following formula in a cell: ""=aic.HelloWorld()""."

    Application.MacroOptions Macro:="aicHelloWorld", description:=Description, Category:="AICells"
End Sub

Function aicHelloWorld():
    Dim PyReturn
    Dim PyParameters
    Dim pb As New PyParameterBuilder
    

    If (IsFXWindowOpen()) Then
        'aicHelloWorld = "#FX"
        Exit Function
    End If
    
    pb.Init "aicHelloWorld"
        
    If TypeOf Application.Caller Is range Then
        DeleteErrorMessage Application.Caller
        On Error GoTo failed
    Else
        GoTo valueError
    End If
    
    Call LogUDFCall("aicHelloWorld", Application.Caller)    


    PyReturn = Py.CallUDF("aicells-server", "aicUDFRunner", pb.GetParameterArray(), ThisWorkbook, Application.Caller)
    
    If CheckIfError(PyReturn) Then
        Call ShowErrors(PyReturn, Application.Caller)
        aicHelloWorld = CVErr(xlErrCalc)
    Else
        aicHelloWorld = PyReturn
    End If
    
    Exit Function
failed:
    aicHelloWorld = Err.Description
    Exit Function
valueError:
    aicHelloWorld = CVErr(xlErrValue)
    Exit Function
End Function

' aicIsEmptyCell

Private Sub aicIsEmptyCell_MacroOptions()
    Dim description As String
    Dim argumentDescriptions() As String
    ReDim argumentDescriptions(1 To 2)
    
    description = "Check wetaher a cell or an a renge is empty, and returns TRUE orFALSE."
    argumentDescriptions(1) = "is a 2 dimensional list of parameter(s). The list contains key value pairs."
    argumentDescriptions(2) = "is a table or a range or a single cell."

    Application.MacroOptions Macro:="aicIsEmptyCell", description:=Description, Category:="AICells", argumentDescriptions:=argumentDescriptions
End Sub

Function aicIsEmptyCell(Optional parameters = Null, Optional input_data = Null):
    Dim PyReturn
    Dim PyParameters
    Dim pb As New PyParameterBuilder
    

    If (IsFXWindowOpen()) Then
        'aicIsEmptyCell = "#FX"
        Exit Function
    End If
    
    pb.Init "aicIsEmptyCell"
        
    If TypeOf Application.Caller Is range Then
        DeleteErrorMessage Application.Caller
        On Error GoTo failed
    Else
        GoTo valueError
    End If
    
    Call LogUDFCall("aicIsEmptyCell", Application.Caller)    

    If TypeOf parameters Is range Then
        'If HasRangeErrors(parameters) Then GoTo valueError
        If ProcessParameterRanges2(pb, parameters, "parameters") = False Then GoTo valueError
        parameters = "@AICELLS-RANGE@"
    End If

    If TypeOf input_data Is range Then
        'If HasRangeErrors(input_data) Then GoTo valueError
        'pb.StoreRange "parameters.input_data", input_data
        If ProcessParameterRanges2(pb, input_data, "parameters.input_data") = False Then GoTo valueError        
        input_data = "@AICELLS-RANGE@"
    ElseIf IsArray(input_data) Then
        pb.StoreArray "parameters.input_data", input_data
        scorers = "@AICELLS-RANGE@"
    End If

    pb.SetUdfArguments (Array( _
        Array("_workbook_path", Application.Caller.Worksheet.Parent.FullName), _
        Array("parameters", parameters), _
        Array("input_data", input_data)))
    PyReturn = Py.CallUDF("aicells-server", "aicUDFRunner", pb.GetParameterArray(), ThisWorkbook, Application.Caller)
    
    If CheckIfError(PyReturn) Then
        Call ShowErrors(PyReturn, Application.Caller)
        aicIsEmptyCell = CVErr(xlErrCalc)
    Else
        aicIsEmptyCell = PyReturn
    End If
    
    Exit Function
failed:
    aicIsEmptyCell = Err.Description
    Exit Function
valueError:
    aicIsEmptyCell = CVErr(xlErrValue)
    Exit Function
End Function

' aicLoadFromDataSource

Private Sub aicLoadFromDataSource_MacroOptions()
    Dim description As String
    Dim argumentDescriptions() As String
    ReDim argumentDescriptions(1 To 5)
    
    description = ""
    argumentDescriptions(1) = "is a 2 dimensional list of parameter(s). The list contains key-value pairs."
    argumentDescriptions(2) = ""
    argumentDescriptions(3) = "is a logical value: to display column headers =TRUE or empty cell; to hide column headers = FALSE."
    argumentDescriptions(4) = "is a logical value: to display row headers =TRUE or empty cell; to hide row headers = FALSE."
    argumentDescriptions(5) = "is a logical value: to transpose the results =TRUE"

    Application.MacroOptions Macro:="aicLoadFromDataSource", description:=Description, Category:="AICells", argumentDescriptions:=argumentDescriptions
End Sub

Function aicLoadFromDataSource(Optional parameters = Null, Optional data_source = Null, Optional display_column_headers = Null, Optional display_row_headers = Null, Optional transpose = Null):
    Dim PyReturn
    Dim PyParameters
    Dim pb As New PyParameterBuilder
    

    If (IsFXWindowOpen()) Then
        'aicLoadFromDataSource = "#FX"
        Exit Function
    End If
    
    pb.Init "aicLoadFromDataSource"
        
    If TypeOf Application.Caller Is range Then
        DeleteErrorMessage Application.Caller
        On Error GoTo failed
    Else
        GoTo valueError
    End If
    
    Call LogUDFCall("aicLoadFromDataSource", Application.Caller)    

    If TypeOf parameters Is range Then
        'If HasRangeErrors(parameters) Then GoTo valueError
        If ProcessParameterRanges2(pb, parameters, "parameters") = False Then GoTo valueError
        parameters = "@AICELLS-RANGE@"
    End If

    If TypeOf data_source Is range Then
        'If HasRangeErrors(data_source) Then GoTo valueError
        'pb.StoreRange "parameters.data_source", data_source
        If ProcessParameterRanges2(pb, data_source, "parameters.data_source") = False Then GoTo valueError        
        data_source = "@AICELLS-RANGE@"
    ElseIf IsArray(data_source) Then
        pb.StoreArray "parameters.data_source", data_source
        scorers = "@AICELLS-RANGE@"
    End If

    If TypeOf display_column_headers Is range Then
        If display_column_headers.Count <> 1 Then: GoTo valueError
        display_column_headers = display_column_headers.Value
    End If

    If TypeOf display_row_headers Is range Then
        If display_row_headers.Count <> 1 Then: GoTo valueError
        display_row_headers = display_row_headers.Value
    End If

    If TypeOf transpose Is range Then
        If transpose.Count <> 1 Then: GoTo valueError
        transpose = transpose.Value
    End If

    pb.SetUdfArguments (Array( _
        Array("_workbook_path", Application.Caller.Worksheet.Parent.FullName), _
        Array("parameters", parameters), _
        Array("data_source", data_source), _
        Array("display_column_headers", display_column_headers), _
        Array("display_row_headers", display_row_headers), _
        Array("transpose", transpose)))
    PyReturn = Py.CallUDF("aicells-server", "aicUDFRunner", pb.GetParameterArray(), ThisWorkbook, Application.Caller)
    
    If CheckIfError(PyReturn) Then
        Call ShowErrors(PyReturn, Application.Caller)
        aicLoadFromDataSource = CVErr(xlErrCalc)
    Else
        aicLoadFromDataSource = PyReturn
    End If
    
    Exit Function
failed:
    aicLoadFromDataSource = Err.Description
    Exit Function
valueError:
    aicLoadFromDataSource = CVErr(xlErrValue)
    Exit Function
End Function

' aicSaveToDataSource

Private Sub aicSaveToDataSource_MacroOptions()
    Dim description As String
    Dim argumentDescriptions() As String
    ReDim argumentDescriptions(1 To 3)
    
    description = ""
    argumentDescriptions(1) = "is a 2 dimensional list of parameter(s). The list contains key-value pairs."
    argumentDescriptions(2) = ""
    argumentDescriptions(3) = "is a table or range with header."

    Application.MacroOptions Macro:="aicSaveToDataSource", description:=Description, Category:="AICells", argumentDescriptions:=argumentDescriptions
End Sub

Function aicSaveToDataSource(Optional parameters = Null, Optional data_source = Null, Optional input_data = Null):
    Dim PyReturn
    Dim PyParameters
    Dim pb As New PyParameterBuilder
    

    If (IsFXWindowOpen()) Then
        'aicSaveToDataSource = "#FX"
        Exit Function
    End If
    
    pb.Init "aicSaveToDataSource"
        
    If TypeOf Application.Caller Is range Then
        DeleteErrorMessage Application.Caller
        On Error GoTo failed
    Else
        GoTo valueError
    End If
    
    Call LogUDFCall("aicSaveToDataSource", Application.Caller)    

    If TypeOf parameters Is range Then
        'If HasRangeErrors(parameters) Then GoTo valueError
        If ProcessParameterRanges2(pb, parameters, "parameters") = False Then GoTo valueError
        parameters = "@AICELLS-RANGE@"
    End If

    If TypeOf data_source Is range Then
        'If HasRangeErrors(data_source) Then GoTo valueError
        'pb.StoreRange "parameters.data_source", data_source
        If ProcessParameterRanges2(pb, data_source, "parameters.data_source") = False Then GoTo valueError        
        data_source = "@AICELLS-RANGE@"
    ElseIf IsArray(data_source) Then
        pb.StoreArray "parameters.data_source", data_source
        scorers = "@AICELLS-RANGE@"
    End If

    If TypeOf input_data Is range Then
        'If HasRangeErrors(input_data) Then GoTo valueError
        'pb.StoreRange "parameters.input_data", input_data
        If ProcessParameterRanges2(pb, input_data, "parameters.input_data") = False Then GoTo valueError        
        input_data = "@AICELLS-RANGE@"
    ElseIf IsArray(input_data) Then
        pb.StoreArray "parameters.input_data", input_data
        scorers = "@AICELLS-RANGE@"
    End If

    pb.SetUdfArguments (Array( _
        Array("_workbook_path", Application.Caller.Worksheet.Parent.FullName), _
        Array("parameters", parameters), _
        Array("data_source", data_source), _
        Array("input_data", input_data)))
    PyReturn = Py.CallUDF("aicells-server", "aicUDFRunner", pb.GetParameterArray(), ThisWorkbook, Application.Caller)
    
    If CheckIfError(PyReturn) Then
        Call ShowErrors(PyReturn, Application.Caller)
        aicSaveToDataSource = CVErr(xlErrCalc)
    Else
        aicSaveToDataSource = PyReturn
    End If
    
    Exit Function
failed:
    aicSaveToDataSource = Err.Description
    Exit Function
valueError:
    aicSaveToDataSource = CVErr(xlErrValue)
    Exit Function
End Function

' aicToolDescription

Private Sub aicToolDescription_MacroOptions()
    Dim description As String
    Dim argumentDescriptions() As String
    ReDim argumentDescriptions(1 To 1)
    
    description = "Returns the description for a specific AIcells tool."
    argumentDescriptions(1) = "is the name of the AIcells tool. You can get the aic function list with the aicTool() function."

    Application.MacroOptions Macro:="aicToolDescription", description:=Description, Category:="AICells", argumentDescriptions:=argumentDescriptions
End Sub

Function aicToolDescription(Optional AIcells_tool_name = Null):
    Dim PyReturn
    Dim PyParameters
    Dim pb As New PyParameterBuilder
    

    If (IsFXWindowOpen()) Then
        'aicToolDescription = "#FX"
        Exit Function
    End If
    
    pb.Init "aicToolDescription"
        
    If TypeOf Application.Caller Is range Then
        DeleteErrorMessage Application.Caller
        On Error GoTo failed
    Else
        GoTo valueError
    End If
    
    Call LogUDFCall("aicToolDescription", Application.Caller)    

    If TypeOf AIcells_tool_name Is range Then
        If AIcells_tool_name.Count <> 1 Then: GoTo valueError
        AIcells_tool_name = AIcells_tool_name.Value
    End If

    pb.SetUdfArguments (Array( _
        Array("_workbook_path", Application.Caller.Worksheet.Parent.FullName), _
        Array("AIcells_tool_name", AIcells_tool_name)))
    PyReturn = Py.CallUDF("aicells-server", "aicUDFRunner", pb.GetParameterArray(), ThisWorkbook, Application.Caller)
    
    If CheckIfError(PyReturn) Then
        Call ShowErrors(PyReturn, Application.Caller)
        aicToolDescription = CVErr(xlErrCalc)
    Else
        aicToolDescription = PyReturn
    End If
    
    Exit Function
failed:
    aicToolDescription = Err.Description
    Exit Function
valueError:
    aicToolDescription = CVErr(xlErrValue)
    Exit Function
End Function

' aicToolList

Private Sub aicToolList_MacroOptions()
    Dim description As String
    Dim argumentDescriptions() As String
    
    
    description = "Lists the available AIcells tools. Enter the following formula in a cell: ""=aic.Tool()""."

    Application.MacroOptions Macro:="aicToolList", description:=Description, Category:="AICells"
End Sub

Function aicToolList():
    Dim PyReturn
    Dim PyParameters
    Dim pb As New PyParameterBuilder
    

    If (IsFXWindowOpen()) Then
        'aicToolList = "#FX"
        Exit Function
    End If
    
    pb.Init "aicToolList"
        
    If TypeOf Application.Caller Is range Then
        DeleteErrorMessage Application.Caller
        On Error GoTo failed
    Else
        GoTo valueError
    End If
    
    Call LogUDFCall("aicToolList", Application.Caller)    


    PyReturn = Py.CallUDF("aicells-server", "aicUDFRunner", pb.GetParameterArray(), ThisWorkbook, Application.Caller)
    
    If CheckIfError(PyReturn) Then
        Call ShowErrors(PyReturn, Application.Caller)
        aicToolList = CVErr(xlErrCalc)
    Else
        aicToolList = PyReturn
    End If
    
    Exit Function
failed:
    aicToolList = Err.Description
    Exit Function
valueError:
    aicToolList = CVErr(xlErrValue)
    Exit Function
End Function

' aicToolParameterDataTypes

Private Sub aicToolParameterDataTypes_MacroOptions()
    Dim description As String
    Dim argumentDescriptions() As String
    ReDim argumentDescriptions(1 To 2)
    
    description = "Lists the parameter data types for a specific AIcells tool."
    argumentDescriptions(1) = "is the name of the AIcells tool. You can get the aic function list with the aicTool() function."
    argumentDescriptions(2) = "is a logical value: to transpose the results =TRUE"

    Application.MacroOptions Macro:="aicToolParameterDataTypes", description:=Description, Category:="AICells", argumentDescriptions:=argumentDescriptions
End Sub

Function aicToolParameterDataTypes(Optional AIcells_tool_name = Null, Optional transpose = Null):
    Dim PyReturn
    Dim PyParameters
    Dim pb As New PyParameterBuilder
    

    If (IsFXWindowOpen()) Then
        'aicToolParameterDataTypes = "#FX"
        Exit Function
    End If
    
    pb.Init "aicToolParameterDataTypes"
        
    If TypeOf Application.Caller Is range Then
        DeleteErrorMessage Application.Caller
        On Error GoTo failed
    Else
        GoTo valueError
    End If
    
    Call LogUDFCall("aicToolParameterDataTypes", Application.Caller)    

    If TypeOf AIcells_tool_name Is range Then
        If AIcells_tool_name.Count <> 1 Then: GoTo valueError
        AIcells_tool_name = AIcells_tool_name.Value
    End If

    If TypeOf transpose Is range Then
        If transpose.Count <> 1 Then: GoTo valueError
        transpose = transpose.Value
    End If

    pb.SetUdfArguments (Array( _
        Array("_workbook_path", Application.Caller.Worksheet.Parent.FullName), _
        Array("AIcells_tool_name", AIcells_tool_name), _
        Array("transpose", transpose)))
    PyReturn = Py.CallUDF("aicells-server", "aicUDFRunner", pb.GetParameterArray(), ThisWorkbook, Application.Caller)
    
    If CheckIfError(PyReturn) Then
        Call ShowErrors(PyReturn, Application.Caller)
        aicToolParameterDataTypes = CVErr(xlErrCalc)
    Else
        aicToolParameterDataTypes = PyReturn
    End If
    
    Exit Function
failed:
    aicToolParameterDataTypes = Err.Description
    Exit Function
valueError:
    aicToolParameterDataTypes = CVErr(xlErrValue)
    Exit Function
End Function

' aicToolParameterDefaultValues

Private Sub aicToolParameterDefaultValues_MacroOptions()
    Dim description As String
    Dim argumentDescriptions() As String
    ReDim argumentDescriptions(1 To 2)
    
    description = "Lists the parameter default values for a specific AIcells tool."
    argumentDescriptions(1) = "is the name of the AIcells tool. You can get the aic function list with the aicTool() function."
    argumentDescriptions(2) = "is a logical value: to transpose the results =TRUE"

    Application.MacroOptions Macro:="aicToolParameterDefaultValues", description:=Description, Category:="AICells", argumentDescriptions:=argumentDescriptions
End Sub

Function aicToolParameterDefaultValues(Optional AIcells_tool_name = Null, Optional transpose = Null):
    Dim PyReturn
    Dim PyParameters
    Dim pb As New PyParameterBuilder
    

    If (IsFXWindowOpen()) Then
        'aicToolParameterDefaultValues = "#FX"
        Exit Function
    End If
    
    pb.Init "aicToolParameterDefaultValues"
        
    If TypeOf Application.Caller Is range Then
        DeleteErrorMessage Application.Caller
        On Error GoTo failed
    Else
        GoTo valueError
    End If
    
    Call LogUDFCall("aicToolParameterDefaultValues", Application.Caller)    

    If TypeOf AIcells_tool_name Is range Then
        If AIcells_tool_name.Count <> 1 Then: GoTo valueError
        AIcells_tool_name = AIcells_tool_name.Value
    End If

    If TypeOf transpose Is range Then
        If transpose.Count <> 1 Then: GoTo valueError
        transpose = transpose.Value
    End If

    pb.SetUdfArguments (Array( _
        Array("_workbook_path", Application.Caller.Worksheet.Parent.FullName), _
        Array("AIcells_tool_name", AIcells_tool_name), _
        Array("transpose", transpose)))
    PyReturn = Py.CallUDF("aicells-server", "aicUDFRunner", pb.GetParameterArray(), ThisWorkbook, Application.Caller)
    
    If CheckIfError(PyReturn) Then
        Call ShowErrors(PyReturn, Application.Caller)
        aicToolParameterDefaultValues = CVErr(xlErrCalc)
    Else
        aicToolParameterDefaultValues = PyReturn
    End If
    
    Exit Function
failed:
    aicToolParameterDefaultValues = Err.Description
    Exit Function
valueError:
    aicToolParameterDefaultValues = CVErr(xlErrValue)
    Exit Function
End Function

' aicToolParameterDescriptions

Private Sub aicToolParameterDescriptions_MacroOptions()
    Dim description As String
    Dim argumentDescriptions() As String
    ReDim argumentDescriptions(1 To 2)
    
    description = "Lists the parameter descriptions for a specific AIcells tool."
    argumentDescriptions(1) = "is the name of the AIcells tool. You can get the aic function list with the aicTool() function."
    argumentDescriptions(2) = "is a logical value: to transpose the results =TRUE"

    Application.MacroOptions Macro:="aicToolParameterDescriptions", description:=Description, Category:="AICells", argumentDescriptions:=argumentDescriptions
End Sub

Function aicToolParameterDescriptions(Optional AIcells_tool_name = Null, Optional transpose = Null):
    Dim PyReturn
    Dim PyParameters
    Dim pb As New PyParameterBuilder
    

    If (IsFXWindowOpen()) Then
        'aicToolParameterDescriptions = "#FX"
        Exit Function
    End If
    
    pb.Init "aicToolParameterDescriptions"
        
    If TypeOf Application.Caller Is range Then
        DeleteErrorMessage Application.Caller
        On Error GoTo failed
    Else
        GoTo valueError
    End If
    
    Call LogUDFCall("aicToolParameterDescriptions", Application.Caller)    

    If TypeOf AIcells_tool_name Is range Then
        If AIcells_tool_name.Count <> 1 Then: GoTo valueError
        AIcells_tool_name = AIcells_tool_name.Value
    End If

    If TypeOf transpose Is range Then
        If transpose.Count <> 1 Then: GoTo valueError
        transpose = transpose.Value
    End If

    pb.SetUdfArguments (Array( _
        Array("_workbook_path", Application.Caller.Worksheet.Parent.FullName), _
        Array("AIcells_tool_name", AIcells_tool_name), _
        Array("transpose", transpose)))
    PyReturn = Py.CallUDF("aicells-server", "aicUDFRunner", pb.GetParameterArray(), ThisWorkbook, Application.Caller)
    
    If CheckIfError(PyReturn) Then
        Call ShowErrors(PyReturn, Application.Caller)
        aicToolParameterDescriptions = CVErr(xlErrCalc)
    Else
        aicToolParameterDescriptions = PyReturn
    End If
    
    Exit Function
failed:
    aicToolParameterDescriptions = Err.Description
    Exit Function
valueError:
    aicToolParameterDescriptions = CVErr(xlErrValue)
    Exit Function
End Function

' aicToolParameters

Private Sub aicToolParameters_MacroOptions()
    Dim description As String
    Dim argumentDescriptions() As String
    ReDim argumentDescriptions(1 To 2)
    
    description = "Lists the parameters for a specific AIcells tool."
    argumentDescriptions(1) = "is the name of the AIcells tool. You can get the aic function list with the aicTool() function."
    argumentDescriptions(2) = "is a logical value: to transpose the results =TRUE"

    Application.MacroOptions Macro:="aicToolParameters", description:=Description, Category:="AICells", argumentDescriptions:=argumentDescriptions
End Sub

Function aicToolParameters(Optional AIcells_tool_name = Null, Optional transpose = Null):
    Dim PyReturn
    Dim PyParameters
    Dim pb As New PyParameterBuilder
    

    If (IsFXWindowOpen()) Then
        'aicToolParameters = "#FX"
        Exit Function
    End If
    
    pb.Init "aicToolParameters"
        
    If TypeOf Application.Caller Is range Then
        DeleteErrorMessage Application.Caller
        On Error GoTo failed
    Else
        GoTo valueError
    End If
    
    Call LogUDFCall("aicToolParameters", Application.Caller)    

    If TypeOf AIcells_tool_name Is range Then
        If AIcells_tool_name.Count <> 1 Then: GoTo valueError
        AIcells_tool_name = AIcells_tool_name.Value
    End If

    If TypeOf transpose Is range Then
        If transpose.Count <> 1 Then: GoTo valueError
        transpose = transpose.Value
    End If

    pb.SetUdfArguments (Array( _
        Array("_workbook_path", Application.Caller.Worksheet.Parent.FullName), _
        Array("AIcells_tool_name", AIcells_tool_name), _
        Array("transpose", transpose)))
    PyReturn = Py.CallUDF("aicells-server", "aicUDFRunner", pb.GetParameterArray(), ThisWorkbook, Application.Caller)
    
    If CheckIfError(PyReturn) Then
        Call ShowErrors(PyReturn, Application.Caller)
        aicToolParameters = CVErr(xlErrCalc)
    Else
        aicToolParameters = PyReturn
    End If
    
    Exit Function
failed:
    aicToolParameters = Err.Description
    Exit Function
valueError:
    aicToolParameters = CVErr(xlErrValue)
    Exit Function
End Function

' aicCorrelationMatrix

Private Sub aicCorrelationMatrix_MacroOptions()
    Dim description As String
    Dim argumentDescriptions() As String
    ReDim argumentDescriptions(1 To 8)
    
    description = "Returns the correlation matrix for the selected columns. Compute pairwise correlation of columns, excluding NA/null values."
    argumentDescriptions(1) = "is a 2 dimensional list of parameter(s). The list contains key-value pairs."
    argumentDescriptions(2) = "is a table or range with header."
    argumentDescriptions(3) = "is a list of selected column header names. For select all columns, leave it empty. These are the columns of the matrix."
    argumentDescriptions(4) = "is a list of selected column header names. For select all columns, leave it empty. These are the rows of the matrix."
    argumentDescriptions(5) = "is a logical value: to return original correlation coefficients leave it empty; to return the absolute values of correlation coefficients = TRUE."
    argumentDescriptions(6) = "is a logical value: set it TRUE to display column headers"
    argumentDescriptions(7) = "is a logical value: set it TRUE to display row headers"
    argumentDescriptions(8) = "is a logical value: to transpose the results =TRUE"

    Application.MacroOptions Macro:="aicCorrelationMatrix", description:=Description, Category:="AICells", argumentDescriptions:=argumentDescriptions
End Sub

Function aicCorrelationMatrix(Optional parameters = Null, Optional input_data = Null, Optional selected_columns_1 = Null, Optional selected_columns_2 = Null, Optional absolute_values = Null, Optional display_column_headers = Null, Optional display_row_headers = Null, Optional transpose = Null):
    Dim PyReturn
    Dim PyParameters
    Dim pb As New PyParameterBuilder
    

    If (IsFXWindowOpen()) Then
        'aicCorrelationMatrix = "#FX"
        Exit Function
    End If
    
    pb.Init "aicCorrelationMatrix"
        
    If TypeOf Application.Caller Is range Then
        DeleteErrorMessage Application.Caller
        On Error GoTo failed
    Else
        GoTo valueError
    End If
    
    Call LogUDFCall("aicCorrelationMatrix", Application.Caller)    

    If TypeOf parameters Is range Then
        'If HasRangeErrors(parameters) Then GoTo valueError
        If ProcessParameterRanges2(pb, parameters, "parameters") = False Then GoTo valueError
        parameters = "@AICELLS-RANGE@"
    End If

    If TypeOf input_data Is range Then
        'If HasRangeErrors(input_data) Then GoTo valueError
        'pb.StoreRange "parameters.input_data", input_data
        If ProcessParameterRanges2(pb, input_data, "parameters.input_data") = False Then GoTo valueError        
        input_data = "@AICELLS-RANGE@"
    ElseIf IsArray(input_data) Then
        pb.StoreArray "parameters.input_data", input_data
        scorers = "@AICELLS-RANGE@"
    End If

    If TypeOf selected_columns_1 Is range Then
        'If HasRangeErrors(selected_columns_1) Then GoTo valueError
        'pb.StoreRange "parameters.selected_columns_1", selected_columns_1
        If ProcessParameterRanges2(pb, selected_columns_1, "parameters.selected_columns_1") = False Then GoTo valueError        
        selected_columns_1 = "@AICELLS-RANGE@"
    ElseIf IsArray(selected_columns_1) Then
        pb.StoreArray "parameters.selected_columns_1", selected_columns_1
        scorers = "@AICELLS-RANGE@"
    End If

    If TypeOf selected_columns_2 Is range Then
        'If HasRangeErrors(selected_columns_2) Then GoTo valueError
        'pb.StoreRange "parameters.selected_columns_2", selected_columns_2
        If ProcessParameterRanges2(pb, selected_columns_2, "parameters.selected_columns_2") = False Then GoTo valueError        
        selected_columns_2 = "@AICELLS-RANGE@"
    ElseIf IsArray(selected_columns_2) Then
        pb.StoreArray "parameters.selected_columns_2", selected_columns_2
        scorers = "@AICELLS-RANGE@"
    End If

    If TypeOf absolute_values Is range Then
        If absolute_values.Count <> 1 Then: GoTo valueError
        absolute_values = absolute_values.Value
    End If

    If TypeOf display_column_headers Is range Then
        If display_column_headers.Count <> 1 Then: GoTo valueError
        display_column_headers = display_column_headers.Value
    End If

    If TypeOf display_row_headers Is range Then
        If display_row_headers.Count <> 1 Then: GoTo valueError
        display_row_headers = display_row_headers.Value
    End If

    If TypeOf transpose Is range Then
        If transpose.Count <> 1 Then: GoTo valueError
        transpose = transpose.Value
    End If

    pb.SetUdfArguments (Array( _
        Array("_workbook_path", Application.Caller.Worksheet.Parent.FullName), _
        Array("parameters", parameters), _
        Array("input_data", input_data), _
        Array("selected_columns_1", selected_columns_1), _
        Array("selected_columns_2", selected_columns_2), _
        Array("absolute_values", absolute_values), _
        Array("display_column_headers", display_column_headers), _
        Array("display_row_headers", display_row_headers), _
        Array("transpose", transpose)))
    PyReturn = Py.CallUDF("aicells-server", "aicUDFRunner", pb.GetParameterArray(), ThisWorkbook, Application.Caller)
    
    If CheckIfError(PyReturn) Then
        Call ShowErrors(PyReturn, Application.Caller)
        aicCorrelationMatrix = CVErr(xlErrCalc)
    Else
        aicCorrelationMatrix = PyReturn
    End If
    
    Exit Function
failed:
    aicCorrelationMatrix = Err.Description
    Exit Function
valueError:
    aicCorrelationMatrix = CVErr(xlErrValue)
    Exit Function
End Function

' aicRandom

Private Sub aicRandom_MacroOptions()
    Dim description As String
    Dim argumentDescriptions() As String
    ReDim argumentDescriptions(1 To 3)
    
    description = "Random number generator"
    argumentDescriptions(1) = "is a 2 dimensional list of parameter(s). The list contains key value pairs."
    argumentDescriptions(2) = ""
    argumentDescriptions(3) = "is a 2 dimensional list of the model's parameter(s). The list contains key value pairs."

    Application.MacroOptions Macro:="aicRandom", description:=Description, Category:="AICells", argumentDescriptions:=argumentDescriptions
End Sub

Function aicRandom(Optional parameters = Null, Optional tool_name = Null, Optional tool_parameters = Null):
    Dim PyReturn
    Dim PyParameters
    Dim pb As New PyParameterBuilder
    
    Application.Volatile

    If (IsFXWindowOpen()) Then
        'aicRandom = "#FX"
        Exit Function
    End If
    
    pb.Init "aicRandom"
        
    If TypeOf Application.Caller Is range Then
        DeleteErrorMessage Application.Caller
        On Error GoTo failed
    Else
        GoTo valueError
    End If
    
    Call LogUDFCall("aicRandom", Application.Caller)    

    If TypeOf parameters Is range Then
        'If HasRangeErrors(parameters) Then GoTo valueError
        If ProcessParameterRanges2(pb, parameters, "parameters") = False Then GoTo valueError
        parameters = "@AICELLS-RANGE@"
    End If

    If TypeOf tool_name Is range Then
        If tool_name.Count <> 1 Then: GoTo valueError
        tool_name = tool_name.Value
    End If

    If TypeOf tool_parameters Is range Then
        'If HasRangeErrors(tool_parameters) Then GoTo valueError
        If ProcessParameterRanges2(pb, tool_parameters, "parameters.tool_parameters") = False Then GoTo valueError
        tool_parameters = "@AICELLS-RANGE@"
    End If

    pb.SetUdfArguments (Array( _
        Array("_workbook_path", Application.Caller.Worksheet.Parent.FullName), _
        Array("parameters", parameters), _
        Array("tool_name", tool_name), _
        Array("tool_parameters", tool_parameters)))
    PyReturn = Py.CallUDF("aicells-server", "aicUDFRunner", pb.GetParameterArray(), ThisWorkbook, Application.Caller)
    
    If CheckIfError(PyReturn) Then
        Call ShowErrors(PyReturn, Application.Caller)
        aicRandom = CVErr(xlErrCalc)
    Else
        aicRandom = PyReturn
    End If
    
    Exit Function
failed:
    aicRandom = Err.Description
    Exit Function
valueError:
    aicRandom = CVErr(xlErrValue)
    Exit Function
End Function

' aicSLModelMetrics

Private Sub aicSLModelMetrics_MacroOptions()
    Dim description As String
    Dim argumentDescriptions() As String
    ReDim argumentDescriptions(1 To 12)
    
    description = "Returns score and metrics for evaluating the quality of a model’s predictions."
    argumentDescriptions(1) = "is a 2 dimensional list of parameter(s). The list contains key-value pairs."
    argumentDescriptions(2) = "is the name of the AIcells tool. You can get the aic function list with the aicTool() function."
    argumentDescriptions(3) = "is a 2 dimensional list of tool parameter(s). The list contains key-value pairs."
    argumentDescriptions(4) = "is a table or range with header."
    argumentDescriptions(5) = "is a selected column header name."
    argumentDescriptions(6) = "is a list of selected column header names for model features. When it's Null or not defined, the model uses all columns except selected_target"
    argumentDescriptions(7) = "is a float value between 0.0 and 1.0: toset the proportion of the dataset to include in the test split."
    argumentDescriptions(8) = "list of metrics."
    argumentDescriptions(9) = "TODO"
    argumentDescriptions(10) = "is a logical value: to display column headers =TRUE or empty cell; to hide column headers = FALSE."
    argumentDescriptions(11) = "is a logical value: to transpose the results =TRUE"
    argumentDescriptions(12) = "Random seed"

    Application.MacroOptions Macro:="aicSLModelMetrics", description:=Description, Category:="AICells", argumentDescriptions:=argumentDescriptions
End Sub

Function aicSLModelMetrics(Optional parameters = Null, Optional AIcells_tool_name = Null, Optional tool_parameters = Null, Optional input_data = Null, Optional selected_target = Null, Optional selected_features = Null, Optional test_size = Null, Optional selected_metrics = Null, Optional greater_is_better = Null, Optional display_column_headers = Null, Optional transpose = Null, Optional seed = Null):
    Dim PyReturn
    Dim PyParameters
    Dim pb As New PyParameterBuilder
    

    If (IsFXWindowOpen()) Then
        'aicSLModelMetrics = "#FX"
        Exit Function
    End If
    
    pb.Init "aicSLModelMetrics"
        
    If TypeOf Application.Caller Is range Then
        DeleteErrorMessage Application.Caller
        On Error GoTo failed
    Else
        GoTo valueError
    End If
    
    Call LogUDFCall("aicSLModelMetrics", Application.Caller)    

    If TypeOf parameters Is range Then
        'If HasRangeErrors(parameters) Then GoTo valueError
        If ProcessParameterRanges2(pb, parameters, "parameters") = False Then GoTo valueError
        parameters = "@AICELLS-RANGE@"
    End If

    If TypeOf AIcells_tool_name Is range Then
        If AIcells_tool_name.Count <> 1 Then: GoTo valueError
        AIcells_tool_name = AIcells_tool_name.Value
    End If

    If TypeOf tool_parameters Is range Then
        'If HasRangeErrors(tool_parameters) Then GoTo valueError
        If ProcessParameterRanges2(pb, tool_parameters, "parameters.tool_parameters") = False Then GoTo valueError
        tool_parameters = "@AICELLS-RANGE@"
    End If

    If TypeOf input_data Is range Then
        'If HasRangeErrors(input_data) Then GoTo valueError
        'pb.StoreRange "parameters.input_data", input_data
        If ProcessParameterRanges2(pb, input_data, "parameters.input_data") = False Then GoTo valueError        
        input_data = "@AICELLS-RANGE@"
    ElseIf IsArray(input_data) Then
        pb.StoreArray "parameters.input_data", input_data
        scorers = "@AICELLS-RANGE@"
    End If

    If TypeOf selected_target Is range Then
        If selected_target.Count <> 1 Then: GoTo valueError
        selected_target = selected_target.Value
    End If

    If TypeOf selected_features Is range Then
        'If HasRangeErrors(selected_features) Then GoTo valueError
        'pb.StoreRange "parameters.selected_features", selected_features
        If ProcessParameterRanges2(pb, selected_features, "parameters.selected_features") = False Then GoTo valueError        
        selected_features = "@AICELLS-RANGE@"
    ElseIf IsArray(selected_features) Then
        pb.StoreArray "parameters.selected_features", selected_features
        scorers = "@AICELLS-RANGE@"
    End If

    If TypeOf test_size Is range Then
        If test_size.Count <> 1 Then: GoTo valueError
        test_size = test_size.Value
    End If

    If TypeOf selected_metrics Is range Then
        'If HasRangeErrors(selected_metrics) Then GoTo valueError
        'pb.StoreRange "parameters.selected_metrics", selected_metrics
        If ProcessParameterRanges2(pb, selected_metrics, "parameters.selected_metrics") = False Then GoTo valueError        
        selected_metrics = "@AICELLS-RANGE@"
    ElseIf IsArray(selected_metrics) Then
        pb.StoreArray "parameters.selected_metrics", selected_metrics
        scorers = "@AICELLS-RANGE@"
    End If

    If TypeOf greater_is_better Is range Then
        If greater_is_better.Count <> 1 Then: GoTo valueError
        greater_is_better = greater_is_better.Value
    End If

    If TypeOf display_column_headers Is range Then
        If display_column_headers.Count <> 1 Then: GoTo valueError
        display_column_headers = display_column_headers.Value
    End If

    If TypeOf transpose Is range Then
        If transpose.Count <> 1 Then: GoTo valueError
        transpose = transpose.Value
    End If

    If TypeOf seed Is range Then
        If seed.Count <> 1 Then: GoTo valueError
        seed = seed.Value
    End If

    pb.SetUdfArguments (Array( _
        Array("_workbook_path", Application.Caller.Worksheet.Parent.FullName), _
        Array("parameters", parameters), _
        Array("AIcells_tool_name", AIcells_tool_name), _
        Array("tool_parameters", tool_parameters), _
        Array("input_data", input_data), _
        Array("selected_target", selected_target), _
        Array("selected_features", selected_features), _
        Array("test_size", test_size), _
        Array("selected_metrics", selected_metrics), _
        Array("greater_is_better", greater_is_better), _
        Array("display_column_headers", display_column_headers), _
        Array("transpose", transpose), _
        Array("seed", seed)))
    PyReturn = Py.CallUDF("aicells-server", "aicUDFRunner", pb.GetParameterArray(), ThisWorkbook, Application.Caller)
    
    If CheckIfError(PyReturn) Then
        Call ShowErrors(PyReturn, Application.Caller)
        aicSLModelMetrics = CVErr(xlErrCalc)
    Else
        aicSLModelMetrics = PyReturn
    End If
    
    Exit Function
failed:
    aicSLModelMetrics = Err.Description
    Exit Function
valueError:
    aicSLModelMetrics = CVErr(xlErrValue)
    Exit Function
End Function

' aicSLModelMetricsCV

Private Sub aicSLModelMetricsCV_MacroOptions()
    Dim description As String
    Dim argumentDescriptions() As String
    ReDim argumentDescriptions(1 To 12)
    
    description = "Returns score and metrics for evaluating the quality of a model’s predictions."
    argumentDescriptions(1) = "is a 2 dimensional list of parameter(s). The list contains key-value pairs."
    argumentDescriptions(2) = "is the name of the AIcells tool. You can get the aic function list with the aicTool() function."
    argumentDescriptions(3) = "is a 2 dimensional list of tool parameter(s). The list contains key-value pairs."
    argumentDescriptions(4) = "is a table or range with header."
    argumentDescriptions(5) = "is a selected column header name."
    argumentDescriptions(6) = "is a list of selected column header names for model features. When it's Null or not defined, the model uses all columns except selected_target"
    argumentDescriptions(7) = "Number of folds. Must be at least 2."
    argumentDescriptions(8) = "list of metrics."
    argumentDescriptions(9) = "TODO"
    argumentDescriptions(10) = "is a logical value: to display column headers =TRUE or empty cell; to hide column headers = FALSE."
    argumentDescriptions(11) = "is a logical value: to transpose the results =TRUE"
    argumentDescriptions(12) = "Random seed"

    Application.MacroOptions Macro:="aicSLModelMetricsCV", description:=Description, Category:="AICells", argumentDescriptions:=argumentDescriptions
End Sub

Function aicSLModelMetricsCV(Optional parameters = Null, Optional AIcells_tool_name = Null, Optional tool_parameters = Null, Optional input_data = Null, Optional selected_target = Null, Optional selected_features = Null, Optional n_splits = Null, Optional selected_metrics = Null, Optional greater_is_better = Null, Optional display_column_headers = Null, Optional transpose = Null, Optional seed = Null):
    Dim PyReturn
    Dim PyParameters
    Dim pb As New PyParameterBuilder
    

    If (IsFXWindowOpen()) Then
        'aicSLModelMetricsCV = "#FX"
        Exit Function
    End If
    
    pb.Init "aicSLModelMetricsCV"
        
    If TypeOf Application.Caller Is range Then
        DeleteErrorMessage Application.Caller
        On Error GoTo failed
    Else
        GoTo valueError
    End If
    
    Call LogUDFCall("aicSLModelMetricsCV", Application.Caller)    

    If TypeOf parameters Is range Then
        'If HasRangeErrors(parameters) Then GoTo valueError
        If ProcessParameterRanges2(pb, parameters, "parameters") = False Then GoTo valueError
        parameters = "@AICELLS-RANGE@"
    End If

    If TypeOf AIcells_tool_name Is range Then
        If AIcells_tool_name.Count <> 1 Then: GoTo valueError
        AIcells_tool_name = AIcells_tool_name.Value
    End If

    If TypeOf tool_parameters Is range Then
        'If HasRangeErrors(tool_parameters) Then GoTo valueError
        If ProcessParameterRanges2(pb, tool_parameters, "parameters.tool_parameters") = False Then GoTo valueError
        tool_parameters = "@AICELLS-RANGE@"
    End If

    If TypeOf input_data Is range Then
        'If HasRangeErrors(input_data) Then GoTo valueError
        'pb.StoreRange "parameters.input_data", input_data
        If ProcessParameterRanges2(pb, input_data, "parameters.input_data") = False Then GoTo valueError        
        input_data = "@AICELLS-RANGE@"
    ElseIf IsArray(input_data) Then
        pb.StoreArray "parameters.input_data", input_data
        scorers = "@AICELLS-RANGE@"
    End If

    If TypeOf selected_target Is range Then
        If selected_target.Count <> 1 Then: GoTo valueError
        selected_target = selected_target.Value
    End If

    If TypeOf selected_features Is range Then
        'If HasRangeErrors(selected_features) Then GoTo valueError
        'pb.StoreRange "parameters.selected_features", selected_features
        If ProcessParameterRanges2(pb, selected_features, "parameters.selected_features") = False Then GoTo valueError        
        selected_features = "@AICELLS-RANGE@"
    ElseIf IsArray(selected_features) Then
        pb.StoreArray "parameters.selected_features", selected_features
        scorers = "@AICELLS-RANGE@"
    End If

    If TypeOf n_splits Is range Then
        If n_splits.Count <> 1 Then: GoTo valueError
        n_splits = n_splits.Value
    End If

    If TypeOf selected_metrics Is range Then
        'If HasRangeErrors(selected_metrics) Then GoTo valueError
        'pb.StoreRange "parameters.selected_metrics", selected_metrics
        If ProcessParameterRanges2(pb, selected_metrics, "parameters.selected_metrics") = False Then GoTo valueError        
        selected_metrics = "@AICELLS-RANGE@"
    ElseIf IsArray(selected_metrics) Then
        pb.StoreArray "parameters.selected_metrics", selected_metrics
        scorers = "@AICELLS-RANGE@"
    End If

    If TypeOf greater_is_better Is range Then
        If greater_is_better.Count <> 1 Then: GoTo valueError
        greater_is_better = greater_is_better.Value
    End If

    If TypeOf display_column_headers Is range Then
        If display_column_headers.Count <> 1 Then: GoTo valueError
        display_column_headers = display_column_headers.Value
    End If

    If TypeOf transpose Is range Then
        If transpose.Count <> 1 Then: GoTo valueError
        transpose = transpose.Value
    End If

    If TypeOf seed Is range Then
        If seed.Count <> 1 Then: GoTo valueError
        seed = seed.Value
    End If

    pb.SetUdfArguments (Array( _
        Array("_workbook_path", Application.Caller.Worksheet.Parent.FullName), _
        Array("parameters", parameters), _
        Array("AIcells_tool_name", AIcells_tool_name), _
        Array("tool_parameters", tool_parameters), _
        Array("input_data", input_data), _
        Array("selected_target", selected_target), _
        Array("selected_features", selected_features), _
        Array("n_splits", n_splits), _
        Array("selected_metrics", selected_metrics), _
        Array("greater_is_better", greater_is_better), _
        Array("display_column_headers", display_column_headers), _
        Array("transpose", transpose), _
        Array("seed", seed)))
    PyReturn = Py.CallUDF("aicells-server", "aicUDFRunner", pb.GetParameterArray(), ThisWorkbook, Application.Caller)
    
    If CheckIfError(PyReturn) Then
        Call ShowErrors(PyReturn, Application.Caller)
        aicSLModelMetricsCV = CVErr(xlErrCalc)
    Else
        aicSLModelMetricsCV = PyReturn
    End If
    
    Exit Function
failed:
    aicSLModelMetricsCV = Err.Description
    Exit Function
valueError:
    aicSLModelMetricsCV = CVErr(xlErrValue)
    Exit Function
End Function

' aicSLModelPredict

Private Sub aicSLModelPredict_MacroOptions()
    Dim description As String
    Dim argumentDescriptions() As String
    ReDim argumentDescriptions(1 To 9)
    
    description = "Calcultes model metrics"
    argumentDescriptions(1) = "is a 2 dimensional list of parameter(s). The list contains key value pairs."
    argumentDescriptions(2) = ""
    argumentDescriptions(3) = "is a 2 dimensional list of the model's parameter(s). The list contains key value pairs."
    argumentDescriptions(4) = "is a table or range with header."
    argumentDescriptions(5) = "is a selected column header name."
    argumentDescriptions(6) = "is a list of selected column header names. When it's Null or not defined, the model uses all columns except selected_target."
    argumentDescriptions(7) = "is a table or range with header."
    argumentDescriptions(8) = "is a logical value: to transpose the results =TRUE"
    argumentDescriptions(9) = "Random seed"

    Application.MacroOptions Macro:="aicSLModelPredict", description:=Description, Category:="AICells", argumentDescriptions:=argumentDescriptions
End Sub

Function aicSLModelPredict(Optional parameters = Null, Optional tool_name = Null, Optional tool_parameters = Null, Optional train_data = Null, Optional selected_target = Null, Optional selected_features = Null, Optional predict_data = Null, Optional transpose = Null, Optional seed = Null):
    Dim PyReturn
    Dim PyParameters
    Dim pb As New PyParameterBuilder
    

    If (IsFXWindowOpen()) Then
        'aicSLModelPredict = "#FX"
        Exit Function
    End If
    
    pb.Init "aicSLModelPredict"
        
    If TypeOf Application.Caller Is range Then
        DeleteErrorMessage Application.Caller
        On Error GoTo failed
    Else
        GoTo valueError
    End If
    
    Call LogUDFCall("aicSLModelPredict", Application.Caller)    

    If TypeOf parameters Is range Then
        'If HasRangeErrors(parameters) Then GoTo valueError
        If ProcessParameterRanges2(pb, parameters, "parameters") = False Then GoTo valueError
        parameters = "@AICELLS-RANGE@"
    End If

    If TypeOf tool_name Is range Then
        If tool_name.Count <> 1 Then: GoTo valueError
        tool_name = tool_name.Value
    End If

    If TypeOf tool_parameters Is range Then
        'If HasRangeErrors(tool_parameters) Then GoTo valueError
        If ProcessParameterRanges2(pb, tool_parameters, "parameters.tool_parameters") = False Then GoTo valueError
        tool_parameters = "@AICELLS-RANGE@"
    End If

    If TypeOf train_data Is range Then
        'If HasRangeErrors(train_data) Then GoTo valueError
        'pb.StoreRange "parameters.train_data", train_data
        If ProcessParameterRanges2(pb, train_data, "parameters.train_data") = False Then GoTo valueError        
        train_data = "@AICELLS-RANGE@"
    ElseIf IsArray(train_data) Then
        pb.StoreArray "parameters.train_data", train_data
        scorers = "@AICELLS-RANGE@"
    End If

    If TypeOf selected_target Is range Then
        If selected_target.Count <> 1 Then: GoTo valueError
        selected_target = selected_target.Value
    End If

    If TypeOf selected_features Is range Then
        'If HasRangeErrors(selected_features) Then GoTo valueError
        'pb.StoreRange "parameters.selected_features", selected_features
        If ProcessParameterRanges2(pb, selected_features, "parameters.selected_features") = False Then GoTo valueError        
        selected_features = "@AICELLS-RANGE@"
    ElseIf IsArray(selected_features) Then
        pb.StoreArray "parameters.selected_features", selected_features
        scorers = "@AICELLS-RANGE@"
    End If

    If TypeOf predict_data Is range Then
        'If HasRangeErrors(predict_data) Then GoTo valueError
        'pb.StoreRange "parameters.predict_data", predict_data
        If ProcessParameterRanges2(pb, predict_data, "parameters.predict_data") = False Then GoTo valueError        
        predict_data = "@AICELLS-RANGE@"
    ElseIf IsArray(predict_data) Then
        pb.StoreArray "parameters.predict_data", predict_data
        scorers = "@AICELLS-RANGE@"
    End If

    If TypeOf transpose Is range Then
        If transpose.Count <> 1 Then: GoTo valueError
        transpose = transpose.Value
    End If

    If TypeOf seed Is range Then
        If seed.Count <> 1 Then: GoTo valueError
        seed = seed.Value
    End If

    pb.SetUdfArguments (Array( _
        Array("_workbook_path", Application.Caller.Worksheet.Parent.FullName), _
        Array("parameters", parameters), _
        Array("tool_name", tool_name), _
        Array("tool_parameters", tool_parameters), _
        Array("train_data", train_data), _
        Array("selected_target", selected_target), _
        Array("selected_features", selected_features), _
        Array("predict_data", predict_data), _
        Array("transpose", transpose), _
        Array("seed", seed)))
    PyReturn = Py.CallUDF("aicells-server", "aicUDFRunner", pb.GetParameterArray(), ThisWorkbook, Application.Caller)
    
    If CheckIfError(PyReturn) Then
        Call ShowErrors(PyReturn, Application.Caller)
        aicSLModelPredict = CVErr(xlErrCalc)
    Else
        aicSLModelPredict = PyReturn
    End If
    
    Exit Function
failed:
    aicSLModelPredict = Err.Description
    Exit Function
valueError:
    aicSLModelPredict = CVErr(xlErrValue)
    Exit Function
End Function

' aicSeabornGetExampleData

Private Sub aicSeabornGetExampleData_MacroOptions()
    Dim description As String
    Dim argumentDescriptions() As String
    ReDim argumentDescriptions(1 To 2)
    
    description = "Load an example dataset from the Seaborn online repository (requires internet)."
    argumentDescriptions(1) = "is a 2 dimensional list of parameter(s). The list contains key-value pairs."
    argumentDescriptions(2) = "Name of the dataset"

    Application.MacroOptions Macro:="aicSeabornGetExampleData", description:=Description, Category:="AICells", argumentDescriptions:=argumentDescriptions
End Sub

Function aicSeabornGetExampleData(Optional parameters = Null, Optional name = Null):
    Dim PyReturn
    Dim PyParameters
    Dim pb As New PyParameterBuilder
    

    If (IsFXWindowOpen()) Then
        'aicSeabornGetExampleData = "#FX"
        Exit Function
    End If
    
    pb.Init "aicSeabornGetExampleData"
        
    If TypeOf Application.Caller Is range Then
        DeleteErrorMessage Application.Caller
        On Error GoTo failed
    Else
        GoTo valueError
    End If
    
    Call LogUDFCall("aicSeabornGetExampleData", Application.Caller)    

    If TypeOf parameters Is range Then
        'If HasRangeErrors(parameters) Then GoTo valueError
        If ProcessParameterRanges2(pb, parameters, "parameters") = False Then GoTo valueError
        parameters = "@AICELLS-RANGE@"
    End If

    If TypeOf name Is range Then
        If name.Count <> 1 Then: GoTo valueError
        name = name.Value
    End If

    pb.SetUdfArguments (Array( _
        Array("_workbook_path", Application.Caller.Worksheet.Parent.FullName), _
        Array("parameters", parameters), _
        Array("name", name)))
    PyReturn = Py.CallUDF("aicells-server", "aicUDFRunner", pb.GetParameterArray(), ThisWorkbook, Application.Caller)
    
    If CheckIfError(PyReturn) Then
        Call ShowErrors(PyReturn, Application.Caller)
        aicSeabornGetExampleData = CVErr(xlErrCalc)
    Else
        aicSeabornGetExampleData = PyReturn
    End If
    
    Exit Function
failed:
    aicSeabornGetExampleData = Err.Description
    Exit Function
valueError:
    aicSeabornGetExampleData = CVErr(xlErrValue)
    Exit Function
End Function

Public Sub SetMacroOptions()
    Call aicCountEmptyCells_MacroOptions
    Call aicDataCopy_MacroOptions
    Call aicDescribe_MacroOptions
    Call aicFillEmptyCells_MacroOptions
    Call aicFunctionDescription_MacroOptions
    Call aicFunctionList_MacroOptions
    Call aicFunctionParameterDataTypes_MacroOptions
    Call aicFunctionParameterDefaultValues_MacroOptions
    Call aicFunctionParameterDescriptions_MacroOptions
    Call aicFunctionParameters_MacroOptions
    Call aicGetDummies_MacroOptions
    Call aicHelloWorld_MacroOptions
    Call aicIsEmptyCell_MacroOptions
    Call aicLoadFromDataSource_MacroOptions
    Call aicSaveToDataSource_MacroOptions
    Call aicToolDescription_MacroOptions
    Call aicToolList_MacroOptions
    Call aicToolParameterDataTypes_MacroOptions
    Call aicToolParameterDefaultValues_MacroOptions
    Call aicToolParameterDescriptions_MacroOptions
    Call aicToolParameters_MacroOptions
    Call aicCorrelationMatrix_MacroOptions
    Call aicRandom_MacroOptions
    Call aicSLModelMetrics_MacroOptions
    Call aicSLModelMetricsCV_MacroOptions
    Call aicSLModelPredict_MacroOptions
    Call aicSeabornGetExampleData_MacroOptions
End Sub

