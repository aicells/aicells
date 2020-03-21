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
        pb.StoreRange "parameters.input_data", input_data
        input_data = "@AICELLS-RANGE@"
    ElseIf IsArray(input_data) Then
        pb.StoreArray "parameters.input_data", input_data
        scorers = "@AICELLS-RANGE@"
    End If

    If TypeOf selected_columns_1 Is range Then
        'If HasRangeErrors(selected_columns_1) Then GoTo valueError
        pb.StoreRange "parameters.selected_columns_1", selected_columns_1
        selected_columns_1 = "@AICELLS-RANGE@"
    ElseIf IsArray(selected_columns_1) Then
        pb.StoreArray "parameters.selected_columns_1", selected_columns_1
        scorers = "@AICELLS-RANGE@"
    End If

    If TypeOf selected_columns_2 Is range Then
        'If HasRangeErrors(selected_columns_2) Then GoTo valueError
        pb.StoreRange "parameters.selected_columns_2", selected_columns_2
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
        Array("parameters", parameters), _
        Array("input_data", input_data), _
        Array("selected_columns_1", selected_columns_1), _
        Array("selected_columns_2", selected_columns_2), _
        Array("absolute_values", absolute_values), _
        Array("display_column_headers", display_column_headers), _
        Array("display_row_headers", display_row_headers), _
        Array("transpose", transpose)))
    PyReturn = Py.CallUDF("udf-server", "aicRaw", pb.GetParameterArray(), ThisWorkbook, Application.Caller)
    
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
        pb.StoreRange "parameters.input_data", input_data
        input_data = "@AICELLS-RANGE@"
    ElseIf IsArray(input_data) Then
        pb.StoreArray "parameters.input_data", input_data
        scorers = "@AICELLS-RANGE@"
    End If

    pb.SetUdfArguments (Array( _
        Array("parameters", parameters), _
        Array("input_data", input_data)))
    PyReturn = Py.CallUDF("udf-server", "aicRaw", pb.GetParameterArray(), ThisWorkbook, Application.Caller)
    
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
        pb.StoreRange "parameters.input_data", input_data
        input_data = "@AICELLS-RANGE@"
    ElseIf IsArray(input_data) Then
        pb.StoreArray "parameters.input_data", input_data
        scorers = "@AICELLS-RANGE@"
    End If

    If TypeOf selected_columns Is range Then
        'If HasRangeErrors(selected_columns) Then GoTo valueError
        pb.StoreRange "parameters.selected_columns", selected_columns
        selected_columns = "@AICELLS-RANGE@"
    ElseIf IsArray(selected_columns) Then
        pb.StoreArray "parameters.selected_columns", selected_columns
        scorers = "@AICELLS-RANGE@"
    End If

    If TypeOf selected_statistics Is range Then
        'If HasRangeErrors(selected_statistics) Then GoTo valueError
        pb.StoreRange "parameters.selected_statistics", selected_statistics
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
    PyReturn = Py.CallUDF("udf-server", "aicRaw", pb.GetParameterArray(), ThisWorkbook, Application.Caller)
    
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
        pb.StoreRange "parameters.input_data", input_data
        input_data = "@AICELLS-RANGE@"
    ElseIf IsArray(input_data) Then
        pb.StoreArray "parameters.input_data", input_data
        scorers = "@AICELLS-RANGE@"
    End If

    If TypeOf selected_columns Is range Then
        'If HasRangeErrors(selected_columns) Then GoTo valueError
        pb.StoreRange "parameters.selected_columns", selected_columns
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
        Array("parameters", parameters), _
        Array("input_data", input_data), _
        Array("selected_columns", selected_columns), _
        Array("transpose", transpose), _
        Array("value", value), _
        Array("method", method), _
        Array("return_all_columns", return_all_columns), _
        Array("limit", limit)))
    PyReturn = Py.CallUDF("udf-server", "aicRaw", pb.GetParameterArray(), ThisWorkbook, Application.Caller)
    
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
        pb.StoreRange "parameters.input_data", input_data
        input_data = "@AICELLS-RANGE@"
    ElseIf IsArray(input_data) Then
        pb.StoreArray "parameters.input_data", input_data
        scorers = "@AICELLS-RANGE@"
    End If

    If TypeOf selected_columns Is range Then
        'If HasRangeErrors(selected_columns) Then GoTo valueError
        pb.StoreRange "parameters.selected_columns", selected_columns
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
        Array("parameters", parameters), _
        Array("input_data", input_data), _
        Array("selected_columns", selected_columns), _
        Array("full_table", full_table), _
        Array("column_header", column_header), _
        Array("transpose", transpose)))
    PyReturn = Py.CallUDF("udf-server", "aicRaw", pb.GetParameterArray(), ThisWorkbook, Application.Caller)
    
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


    PyReturn = Py.CallUDF("udf-server", "aicRaw", pb.GetParameterArray(), ThisWorkbook, Application.Caller)
    
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
        pb.StoreRange "parameters.input_data", input_data
        input_data = "@AICELLS-RANGE@"
    ElseIf IsArray(input_data) Then
        pb.StoreArray "parameters.input_data", input_data
        scorers = "@AICELLS-RANGE@"
    End If

    pb.SetUdfArguments (Array( _
        Array("parameters", parameters), _
        Array("input_data", input_data)))
    PyReturn = Py.CallUDF("udf-server", "aicRaw", pb.GetParameterArray(), ThisWorkbook, Application.Caller)
    
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
        pb.StoreRange "parameters.input_data", input_data
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
        pb.StoreRange "parameters.selected_features", selected_features
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
        pb.StoreRange "parameters.selected_metrics", selected_metrics
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
    PyReturn = Py.CallUDF("udf-server", "aicRaw", pb.GetParameterArray(), ThisWorkbook, Application.Caller)
    
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
        pb.StoreRange "parameters.train_data", train_data
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
        pb.StoreRange "parameters.selected_features", selected_features
        selected_features = "@AICELLS-RANGE@"
    ElseIf IsArray(selected_features) Then
        pb.StoreArray "parameters.selected_features", selected_features
        scorers = "@AICELLS-RANGE@"
    End If

    If TypeOf predict_data Is range Then
        'If HasRangeErrors(predict_data) Then GoTo valueError
        pb.StoreRange "parameters.predict_data", predict_data
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
        Array("parameters", parameters), _
        Array("tool_name", tool_name), _
        Array("tool_parameters", tool_parameters), _
        Array("train_data", train_data), _
        Array("selected_target", selected_target), _
        Array("selected_features", selected_features), _
        Array("predict_data", predict_data), _
        Array("transpose", transpose), _
        Array("seed", seed)))
    PyReturn = Py.CallUDF("udf-server", "aicRaw", pb.GetParameterArray(), ThisWorkbook, Application.Caller)
    
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

' aicTool

Private Sub aicTool_MacroOptions()
    Dim description As String
    Dim argumentDescriptions() As String
    
    
    description = "Lists the available AIcells tools. Enter the following formula in a cell: ""=aic.Tool()""."

    Application.MacroOptions Macro:="aicTool", description:=Description, Category:="AICells"
End Sub

Function aicTool():
    Dim PyReturn
    Dim PyParameters
    Dim pb As New PyParameterBuilder
    
    If (IsFXWindowOpen()) Then
        'aicTool = "#FX"
        Exit Function
    End If
    
    pb.Init "aicTool"
        
    If TypeOf Application.Caller Is range Then
        DeleteErrorMessage Application.Caller
        On Error GoTo failed
    Else
        GoTo valueError
    End If
    
    Call LogUDFCall("aicTool", Application.Caller)    


    PyReturn = Py.CallUDF("udf-server", "aicRaw", pb.GetParameterArray(), ThisWorkbook, Application.Caller)
    
    If CheckIfError(PyReturn) Then
        Call ShowErrors(PyReturn, Application.Caller)
        aicTool = CVErr(xlErrCalc)
    Else
        aicTool = PyReturn
    End If
    
    Exit Function
failed:
    aicTool = Err.Description
    Exit Function
valueError:
    aicTool = CVErr(xlErrValue)
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
        Array("AIcells_tool_name", AIcells_tool_name)))
    PyReturn = Py.CallUDF("udf-server", "aicRaw", pb.GetParameterArray(), ThisWorkbook, Application.Caller)
    
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

' aicToolParameterDataTypes

Private Sub aicToolParameterDataTypes_MacroOptions()
    Dim description As String
    Dim argumentDescriptions() As String
    ReDim argumentDescriptions(1 To 1)
    
    description = "Lists the parameter data types for a specific AIcells tool."
    argumentDescriptions(1) = "is the name of the AIcells tool. You can get the aic function list with the aicTool() function."

    Application.MacroOptions Macro:="aicToolParameterDataTypes", description:=Description, Category:="AICells", argumentDescriptions:=argumentDescriptions
End Sub

Function aicToolParameterDataTypes(Optional AIcells_tool_name = Null):
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

    pb.SetUdfArguments (Array( _
        Array("AIcells_tool_name", AIcells_tool_name)))
    PyReturn = Py.CallUDF("udf-server", "aicRaw", pb.GetParameterArray(), ThisWorkbook, Application.Caller)
    
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
    ReDim argumentDescriptions(1 To 1)
    
    description = "Lists the parameter default values for a specific AIcells tool."
    argumentDescriptions(1) = "is the name of the AIcells tool. You can get the aic function list with the aicTool() function."

    Application.MacroOptions Macro:="aicToolParameterDefaultValues", description:=Description, Category:="AICells", argumentDescriptions:=argumentDescriptions
End Sub

Function aicToolParameterDefaultValues(Optional AIcells_tool_name = Null):
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

    pb.SetUdfArguments (Array( _
        Array("AIcells_tool_name", AIcells_tool_name)))
    PyReturn = Py.CallUDF("udf-server", "aicRaw", pb.GetParameterArray(), ThisWorkbook, Application.Caller)
    
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
    ReDim argumentDescriptions(1 To 1)
    
    description = "Lists the parameter descriptions for a specific AIcells tool."
    argumentDescriptions(1) = "is the name of the AIcells tool. You can get the aic function list with the aicTool() function."

    Application.MacroOptions Macro:="aicToolParameterDescriptions", description:=Description, Category:="AICells", argumentDescriptions:=argumentDescriptions
End Sub

Function aicToolParameterDescriptions(Optional AIcells_tool_name = Null):
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

    pb.SetUdfArguments (Array( _
        Array("AIcells_tool_name", AIcells_tool_name)))
    PyReturn = Py.CallUDF("udf-server", "aicRaw", pb.GetParameterArray(), ThisWorkbook, Application.Caller)
    
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
    ReDim argumentDescriptions(1 To 1)
    
    description = "Lists the parameters for a specific AIcells tool."
    argumentDescriptions(1) = "is the name of the AIcells tool. You can get the aic function list with the aicTool() function."

    Application.MacroOptions Macro:="aicToolParameters", description:=Description, Category:="AICells", argumentDescriptions:=argumentDescriptions
End Sub

Function aicToolParameters(Optional AIcells_tool_name = Null):
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

    pb.SetUdfArguments (Array( _
        Array("AIcells_tool_name", AIcells_tool_name)))
    PyReturn = Py.CallUDF("udf-server", "aicRaw", pb.GetParameterArray(), ThisWorkbook, Application.Caller)
    
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

' aicUDF

Private Sub aicUDF_MacroOptions()
    Dim description As String
    Dim argumentDescriptions() As String
    
    
    description = "Lists the available AIcells Excel functions (UDF). Enter the following formula in a cell: ""=aic.UDF()""."

    Application.MacroOptions Macro:="aicUDF", description:=Description, Category:="AICells"
End Sub

Function aicUDF():
    Dim PyReturn
    Dim PyParameters
    Dim pb As New PyParameterBuilder
    
    If (IsFXWindowOpen()) Then
        'aicUDF = "#FX"
        Exit Function
    End If
    
    pb.Init "aicUDF"
        
    If TypeOf Application.Caller Is range Then
        DeleteErrorMessage Application.Caller
        On Error GoTo failed
    Else
        GoTo valueError
    End If
    
    Call LogUDFCall("aicUDF", Application.Caller)    


    PyReturn = Py.CallUDF("udf-server", "aicRaw", pb.GetParameterArray(), ThisWorkbook, Application.Caller)
    
    If CheckIfError(PyReturn) Then
        Call ShowErrors(PyReturn, Application.Caller)
        aicUDF = CVErr(xlErrCalc)
    Else
        aicUDF = PyReturn
    End If
    
    Exit Function
failed:
    aicUDF = Err.Description
    Exit Function
valueError:
    aicUDF = CVErr(xlErrValue)
    Exit Function
End Function

' aicUDFDescription

Private Sub aicUDFDescription_MacroOptions()
    Dim description As String
    Dim argumentDescriptions() As String
    ReDim argumentDescriptions(1 To 1)
    
    description = "Returns the description for a specific AIcells Excel function (UDF)."
    argumentDescriptions(1) = "is the name of an AIcells Excel function (UDF). You can get the list of AIcells function names with the aicUDF() function."

    Application.MacroOptions Macro:="aicUDFDescription", description:=Description, Category:="AICells", argumentDescriptions:=argumentDescriptions
End Sub

Function aicUDFDescription(Optional AIcells_UDF_name = Null):
    Dim PyReturn
    Dim PyParameters
    Dim pb As New PyParameterBuilder
    
    If (IsFXWindowOpen()) Then
        'aicUDFDescription = "#FX"
        Exit Function
    End If
    
    pb.Init "aicUDFDescription"
        
    If TypeOf Application.Caller Is range Then
        DeleteErrorMessage Application.Caller
        On Error GoTo failed
    Else
        GoTo valueError
    End If
    
    Call LogUDFCall("aicUDFDescription", Application.Caller)    

    If TypeOf AIcells_UDF_name Is range Then
        If AIcells_UDF_name.Count <> 1 Then: GoTo valueError
        AIcells_UDF_name = AIcells_UDF_name.Value
    End If

    pb.SetUdfArguments (Array( _
        Array("AIcells_UDF_name", AIcells_UDF_name)))
    PyReturn = Py.CallUDF("udf-server", "aicRaw", pb.GetParameterArray(), ThisWorkbook, Application.Caller)
    
    If CheckIfError(PyReturn) Then
        Call ShowErrors(PyReturn, Application.Caller)
        aicUDFDescription = CVErr(xlErrCalc)
    Else
        aicUDFDescription = PyReturn
    End If
    
    Exit Function
failed:
    aicUDFDescription = Err.Description
    Exit Function
valueError:
    aicUDFDescription = CVErr(xlErrValue)
    Exit Function
End Function

' aicUDFParameterDataTypes

Private Sub aicUDFParameterDataTypes_MacroOptions()
    Dim description As String
    Dim argumentDescriptions() As String
    ReDim argumentDescriptions(1 To 1)
    
    description = "Lists the parameter data types for a specific AIcells Excel function (UDF)."
    argumentDescriptions(1) = "is the name of an AIcells Excel function (UDF). You can get the list of AIcells function names with the aicUDF() function."

    Application.MacroOptions Macro:="aicUDFParameterDataTypes", description:=Description, Category:="AICells", argumentDescriptions:=argumentDescriptions
End Sub

Function aicUDFParameterDataTypes(Optional AIcells_UDF_name = Null):
    Dim PyReturn
    Dim PyParameters
    Dim pb As New PyParameterBuilder
    
    If (IsFXWindowOpen()) Then
        'aicUDFParameterDataTypes = "#FX"
        Exit Function
    End If
    
    pb.Init "aicUDFParameterDataTypes"
        
    If TypeOf Application.Caller Is range Then
        DeleteErrorMessage Application.Caller
        On Error GoTo failed
    Else
        GoTo valueError
    End If
    
    Call LogUDFCall("aicUDFParameterDataTypes", Application.Caller)    

    If TypeOf AIcells_UDF_name Is range Then
        If AIcells_UDF_name.Count <> 1 Then: GoTo valueError
        AIcells_UDF_name = AIcells_UDF_name.Value
    End If

    pb.SetUdfArguments (Array( _
        Array("AIcells_UDF_name", AIcells_UDF_name)))
    PyReturn = Py.CallUDF("udf-server", "aicRaw", pb.GetParameterArray(), ThisWorkbook, Application.Caller)
    
    If CheckIfError(PyReturn) Then
        Call ShowErrors(PyReturn, Application.Caller)
        aicUDFParameterDataTypes = CVErr(xlErrCalc)
    Else
        aicUDFParameterDataTypes = PyReturn
    End If
    
    Exit Function
failed:
    aicUDFParameterDataTypes = Err.Description
    Exit Function
valueError:
    aicUDFParameterDataTypes = CVErr(xlErrValue)
    Exit Function
End Function

' aicUDFParameterDefaultValues

Private Sub aicUDFParameterDefaultValues_MacroOptions()
    Dim description As String
    Dim argumentDescriptions() As String
    ReDim argumentDescriptions(1 To 1)
    
    description = "Lists the parameter default values for a specific AIcells Excel function (UDF)."
    argumentDescriptions(1) = "is the name of an AIcells Excel function (UDF). You can get the list of AIcells function names with the aicUDF() function."

    Application.MacroOptions Macro:="aicUDFParameterDefaultValues", description:=Description, Category:="AICells", argumentDescriptions:=argumentDescriptions
End Sub

Function aicUDFParameterDefaultValues(Optional AIcells_UDF_name = Null):
    Dim PyReturn
    Dim PyParameters
    Dim pb As New PyParameterBuilder
    
    If (IsFXWindowOpen()) Then
        'aicUDFParameterDefaultValues = "#FX"
        Exit Function
    End If
    
    pb.Init "aicUDFParameterDefaultValues"
        
    If TypeOf Application.Caller Is range Then
        DeleteErrorMessage Application.Caller
        On Error GoTo failed
    Else
        GoTo valueError
    End If
    
    Call LogUDFCall("aicUDFParameterDefaultValues", Application.Caller)    

    If TypeOf AIcells_UDF_name Is range Then
        If AIcells_UDF_name.Count <> 1 Then: GoTo valueError
        AIcells_UDF_name = AIcells_UDF_name.Value
    End If

    pb.SetUdfArguments (Array( _
        Array("AIcells_UDF_name", AIcells_UDF_name)))
    PyReturn = Py.CallUDF("udf-server", "aicRaw", pb.GetParameterArray(), ThisWorkbook, Application.Caller)
    
    If CheckIfError(PyReturn) Then
        Call ShowErrors(PyReturn, Application.Caller)
        aicUDFParameterDefaultValues = CVErr(xlErrCalc)
    Else
        aicUDFParameterDefaultValues = PyReturn
    End If
    
    Exit Function
failed:
    aicUDFParameterDefaultValues = Err.Description
    Exit Function
valueError:
    aicUDFParameterDefaultValues = CVErr(xlErrValue)
    Exit Function
End Function

' aicUDFParameterDescriptions

Private Sub aicUDFParameterDescriptions_MacroOptions()
    Dim description As String
    Dim argumentDescriptions() As String
    ReDim argumentDescriptions(1 To 1)
    
    description = "Lists the parameter descriptions for a specific AIcells Excel function (UDF)."
    argumentDescriptions(1) = "is the name of an AIcells Excel function (UDF). You can get the list of AIcells function names with the aicUDF() function."

    Application.MacroOptions Macro:="aicUDFParameterDescriptions", description:=Description, Category:="AICells", argumentDescriptions:=argumentDescriptions
End Sub

Function aicUDFParameterDescriptions(Optional AIcells_UDF_name = Null):
    Dim PyReturn
    Dim PyParameters
    Dim pb As New PyParameterBuilder
    
    If (IsFXWindowOpen()) Then
        'aicUDFParameterDescriptions = "#FX"
        Exit Function
    End If
    
    pb.Init "aicUDFParameterDescriptions"
        
    If TypeOf Application.Caller Is range Then
        DeleteErrorMessage Application.Caller
        On Error GoTo failed
    Else
        GoTo valueError
    End If
    
    Call LogUDFCall("aicUDFParameterDescriptions", Application.Caller)    

    If TypeOf AIcells_UDF_name Is range Then
        If AIcells_UDF_name.Count <> 1 Then: GoTo valueError
        AIcells_UDF_name = AIcells_UDF_name.Value
    End If

    pb.SetUdfArguments (Array( _
        Array("AIcells_UDF_name", AIcells_UDF_name)))
    PyReturn = Py.CallUDF("udf-server", "aicRaw", pb.GetParameterArray(), ThisWorkbook, Application.Caller)
    
    If CheckIfError(PyReturn) Then
        Call ShowErrors(PyReturn, Application.Caller)
        aicUDFParameterDescriptions = CVErr(xlErrCalc)
    Else
        aicUDFParameterDescriptions = PyReturn
    End If
    
    Exit Function
failed:
    aicUDFParameterDescriptions = Err.Description
    Exit Function
valueError:
    aicUDFParameterDescriptions = CVErr(xlErrValue)
    Exit Function
End Function

' aicUDFParameters

Private Sub aicUDFParameters_MacroOptions()
    Dim description As String
    Dim argumentDescriptions() As String
    ReDim argumentDescriptions(1 To 1)
    
    description = "Lists the parameters for a specific AIcells Excel function (UDF)."
    argumentDescriptions(1) = "is the name of an AIcells Excel function (UDF). You can get the list of AIcells function names with the aicUDF() function."

    Application.MacroOptions Macro:="aicUDFParameters", description:=Description, Category:="AICells", argumentDescriptions:=argumentDescriptions
End Sub

Function aicUDFParameters(Optional AIcells_UDF_name = Null):
    Dim PyReturn
    Dim PyParameters
    Dim pb As New PyParameterBuilder
    
    If (IsFXWindowOpen()) Then
        'aicUDFParameters = "#FX"
        Exit Function
    End If
    
    pb.Init "aicUDFParameters"
        
    If TypeOf Application.Caller Is range Then
        DeleteErrorMessage Application.Caller
        On Error GoTo failed
    Else
        GoTo valueError
    End If
    
    Call LogUDFCall("aicUDFParameters", Application.Caller)    

    If TypeOf AIcells_UDF_name Is range Then
        If AIcells_UDF_name.Count <> 1 Then: GoTo valueError
        AIcells_UDF_name = AIcells_UDF_name.Value
    End If

    pb.SetUdfArguments (Array( _
        Array("AIcells_UDF_name", AIcells_UDF_name)))
    PyReturn = Py.CallUDF("udf-server", "aicRaw", pb.GetParameterArray(), ThisWorkbook, Application.Caller)
    
    If CheckIfError(PyReturn) Then
        Call ShowErrors(PyReturn, Application.Caller)
        aicUDFParameters = CVErr(xlErrCalc)
    Else
        aicUDFParameters = PyReturn
    End If
    
    Exit Function
failed:
    aicUDFParameters = Err.Description
    Exit Function
valueError:
    aicUDFParameters = CVErr(xlErrValue)
    Exit Function
End Function

Public Sub SetMacroOptions()
    Call aicCorrelationMatrix_MacroOptions
    Call aicCountEmptyCells_MacroOptions
    Call aicDescribe_MacroOptions
    Call aicFillEmptyCells_MacroOptions
    Call aicGetDummies_MacroOptions
    Call aicHelloWorld_MacroOptions
    Call aicIsEmptyCell_MacroOptions
    Call aicSLModelMetrics_MacroOptions
    Call aicSLModelPredict_MacroOptions
    Call aicTool_MacroOptions
    Call aicToolDescription_MacroOptions
    Call aicToolParameterDataTypes_MacroOptions
    Call aicToolParameterDefaultValues_MacroOptions
    Call aicToolParameterDescriptions_MacroOptions
    Call aicToolParameters_MacroOptions
    Call aicUDF_MacroOptions
    Call aicUDFDescription_MacroOptions
    Call aicUDFParameterDataTypes_MacroOptions
    Call aicUDFParameterDefaultValues_MacroOptions
    Call aicUDFParameterDescriptions_MacroOptions
    Call aicUDFParameters_MacroOptions
End Sub

