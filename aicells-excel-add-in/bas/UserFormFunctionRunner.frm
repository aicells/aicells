VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserFormFunctionRunner 
   Caption         =   "AIcells function runner"
   ClientHeight    =   6735
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7935
   OleObjectBlob   =   "UserFormFunctionRunner.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserFormFunctionRunner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim abort As Boolean

Dim calledFromMacro As Boolean

'Dim calledFromMacroRange As range
'Dim calledFromMacroRecalculate As Boolean
'
'Public Function SetMacroMode(calledFromMacroRange, recalculate)
'    calledFromMacro = True
'    calledFromMacroRecalculate = recalculate
'End Function
'
'Public Function SetRibbonMode()
'    calledFromMacro = False
'End Function
'
'
Private Sub UserForm_Activate()
    Dim referencedRange As range
    If Selection.Columns.Count = 1 Then
        If Selection.Rows.Count = 1 And Selection.Columns.Count = 1 Then
            If Selection.cells(1, 1).HasFormula Then
                Set referencedRange = DecodeRangeReference(Selection.cells(1, 1))
                If Not (referencedRange Is Nothing) Then
                    Log ("Selected parameter range: " & Replace(referencedRange.address(External:=True), "$", ""))
                Else
                    Log ("ERROR: Invalid parameter range selected.")
                    Exit Sub
                End If
            End If
        Else
            ' single column, multiple rows
            Log ("Selected function list range (" & CStr(Selection.Rows.Count) & " functions): " & Replace(Selection.address(External:=True), "$", ""))
        End If
    Else
        Log ("Selected parameter range: " & Replace(Selection.address(External:=True), "$", ""))
    End If
End Sub

'    If calledFromMacro Then
'        Call UserFormFunctionRunner.Log("RunFunction macro called", True)
'        Call UserFormFunctionRunner.ExecuteAICellsFunctionList(calledFromMacroRange, calledFromMacroRecalculate)
'    End If
'End Sub

Private Sub Log(l As String, Optional clear As Boolean = False)
    If clear Then
        UserFormFunctionRunner.TextBoxLog.value = ""
    End If
    If l <> "" Then
        UserFormFunctionRunner.TextBoxLog.value = UserFormFunctionRunner.TextBoxLog.value & Format(Now, "hh:nn:ss: ") & l & vbCr
    End If
End Sub

Private Sub ShowErrors(e)
    Dim i As Long
    Debug.Print "---"
    
    Dim msg As String
    
    For i = 1 To UBound(e, 1)
        Debug.Print e(i, 0)
        Debug.Print e(i, 1)
        Debug.Print "---"
        msg = msg + e(i, 1) + " (" + e(i, 0) + ")" + vbLf
    Next i
    
    Log "ERROR: " + vbLf + msg
End Sub

Private Sub CommandButtonAbort_Click()
    Dim PyQueueReturn
    abort = True
    PyQueueReturn = Py.CallUDF("aicells-server", "aicAbortProcess", Array(), ThisWorkbook, Application.Caller)
End Sub

Private Sub CommandButtonRun_Click()
    Call ExecuteAICellsFunctionList(Selection, CheckBoxRecalculate.value)
End Sub

Public Sub ExecuteAICellsFunctionList(r As range, recalculate As Boolean)
    Dim referencedRange As range
    'Dim r As range
    Dim b As Boolean
    Dim y As Long
    Dim isError As Boolean
    
    UserFormFunctionRunner.CommandButtonRun.Enabled = False
    UserFormFunctionRunner.CommandButtonAbort.Enabled = True
    
    On Error GoTo valueError
    
    Call Log("", True)
    
    'Set r = Selection
        
    If r.Columns.Count = 1 Then
        isError = False
        For y = 1 To r.Rows.Count
            If r.cells(y, 1).HasFormula Then
                Set referencedRange = DecodeRangeReference(r.cells(y, 1))
                If (referencedRange Is Nothing) Then
                    Log ("ERROR: Invalid parameter range selected: " & Replace(r.cells(1, 1).address(External:=True), "$", ""))
                    isError = True
                End If
            End If
        Next y
        
        If isError Then
            UserFormFunctionRunner.CommandButtonRun.Enabled = True
            UserFormFunctionRunner.CommandButtonAbort.Enabled = False
            Exit Sub
        End If
        
        'Log ("Selected function list range (" & CStr(Selection.Rows.Count) & " functions): " & Replace(Selection.address(External:=True), "$", ""))
                
        For y = 1 To r.Rows.Count
            Set referencedRange = DecodeRangeReference(r.cells(y, 1))
            b = ExecuteAICellsFunction(referencedRange)
            If Not b Then
                Exit Sub
            End If
            
            If recalculate Then
                Log ("Recalculating all open workbook...")
                Calculate
            End If
            
            Log ("---")
        Next y
        Log ("Done.")
        UserFormFunctionRunner.CommandButtonRun.Enabled = True
        UserFormFunctionRunner.CommandButtonAbort.Enabled = False
        Exit Sub
    End If
    
    'Log ("Selected parameter range: " & Replace(r.address(External:=True), "$", ""))
    
    b = ExecuteAICellsFunction(r)
    
    If recalculate Then
        Log ("Recalculating all open workbook...")
        Calculate
    End If
    
    Log ("Done.")
    
    UserFormFunctionRunner.CommandButtonRun.Enabled = True
    UserFormFunctionRunner.CommandButtonAbort.Enabled = False
    
    Exit Sub
    
valueError:
    Log ("Error")
    UserFormFunctionRunner.CommandButtonRun.Enabled = True
    UserFormFunctionRunner.CommandButtonAbort.Enabled = False
End Sub

Function ExecuteAICellsFunction(r As range) As Boolean
    Dim PyReturn
    Dim PyQueueReturn
    Dim PyParameters
    Dim pb As New PyParameterBuilder
    'Dim r As range
    Dim aicFunction As String
    Dim output As range
    Dim y As Long
    Dim tStart, tLast, tDiff As Double
    Dim isOutputDatasource As Boolean
    Dim isOutputNothing As Boolean
    Dim isOutputSVG As Boolean
    Dim results As Variant
    Dim isAICellsFunctionDefined As Boolean
    
    Dim shp, shpNew As Shape
    
    Dim shpLeft, shpTop, shpWidth, shpHeight As Long
    
    On Error GoTo valueError
    
    tStart = Timer
    
    abort = False
    
    ExecuteAICellsFunction = True
    
    isOutputDatasource = False
    isOutputNothing = True
    isOutputSVG = False
    isAICellsFunctionDefined = False
        
        
    For y = 1 To r.Rows.Count
        If r.cells(y, 1).value = "function" Then
            aicFunction = r.cells(y, 2).value
            isAICellsFunctionDefined = True
        End If
    Next y
    
    If Not isAICellsFunctionDefined Then
        Call Log("AIcells function is not defined")
        ExecuteAICellsFunction = False
        Exit Function
    End If
    
    Call Log(aicFunction & " running (" & Replace(r.address(External:=True), "$", "") & ")...")
    
    pb.Init aicFunction

    If ProcessParameterRanges2(pb, r, "parameters") = False Then GoTo valueError
    
    Set output = pb.GetRangeByName("parameters.output")
    
    If output Is Nothing Then
        Log ("WARNING: output range is not defined!")
        isOutputNothing = True
    Else
        If output.cells(1, 1).value = "data_source" Then
            Call Log("Selected output data source: " & Replace(output.address(External:=True), "$", ""))
            isOutputDatasource = True
        Else
            Call Log("Selected output range: " & Replace(output.address(External:=True), "$", ""))
            isOutputNothing = False
        End If
    End If
        
    pb.SetUdfArguments (Array( _
        Array("_workbook_path", ActiveWorkbook.FullName), _
        Array("parameters", "@AICELLS-RANGE@") _
    ))

    PyReturn = Py.CallUDF("aicells-server", "aicProcessRunner", pb.GetParameterArray(), ThisWorkbook, Application.Caller)
    
    If PyReturn <> "OK" Then
        UserFormFunctionRunner.CommandButtonRun.Enabled = True
        UserFormFunctionRunner.CommandButtonAbort.Enabled = False
        Log ("ERROR: Python server is busy.")
        Exit Function
    End If
        
    tLast = Timer
    
    Do While Not abort:
        DoEvents
        tDiff = Abs(Timer - tLast) ' we can just pass midnight
        If tDiff > 0.01 Then
            PyQueueReturn = Py.CallUDF("aicells-server", "aicQueueGet", Array(), ThisWorkbook, Application.Caller)
            If PyQueueReturn(0) = "result" Then
                abort = True
                results = PyQueueReturn(1)
                
                If CheckIfError(results) Then
                    Call ShowErrors(results)
                    UserFormFunctionRunner.CommandButtonRun.Enabled = True
                    UserFormFunctionRunner.CommandButtonAbort.Enabled = False
                    Exit Function
                Else
                
                
                    If TypeName(results) = "Variant()" Then
                        On Error Resume Next
                        If UBound(results, 1) = 0 And UBound(results, 2) = 2 Then
                            If results(0, 0) = "#AICELLS-SVG!" Then
                                isOutputSVG = True
                            End If
                        End If
                    End If
                
                    If isOutputSVG Then
                        ' output is an svg
                        If results(0, 1) <> "" Then
                            ' picture has a name
                            Set shp = Nothing
                            On Error Resume Next
                            Set shp = output.Parent.Shapes(results(0, 1))
                            On Error GoTo valueError
                            
                            If Not (shp Is Nothing) Then
                                ' picture exists with the same name
                                shpLeft = shp.left
                                shpTop = shp.top
                                shpWidth = shp.width
                                shpHeight = shp.height
                                shp.Delete
                            
                                Set shpNew = output.Parent.Shapes.AddPicture(results(0, 2), False, True, shpLeft, shpTop, shpWidth, shpHeight)
                                shpNew.name = results(0, 1)
                            Else
                                Set shpNew = output.Parent.Shapes.AddPicture(results(0, 2), False, True, output.left, output.top, -1, -1)
                                shpNew.name = results(0, 1)
                            End If
                        Else
                            Set shpNew = output.Parent.Shapes.AddPicture(results(0, 2), False, True, output.left, output.top, -1, -1)
                        End If
                        
                    ElseIf Not isOutputDatasource Then
                        If Not isOutputNothing Then
                            ' output to range
                            Log ("Writing results to range...")
                            If IsArray(results) Then
                                Set output = output.Resize(UBound(results, 1) + 1, UBound(results, 2) + 1)
                            Else
                                output.cells(1, 1).value = results
                            End If
                            output.value = results
                        Else
                            Log ("Output range is not defined.")
                        End If
                    Else
                        ' output to data source
                        Log (results)
                        
                        For y = 1 To output.Rows.Count
                            If output.cells(y, 1).value = "hash" Then
                                output.cells(y, 2).value = results
                            End If
                        Next y
                        
                    End If
                    
                    Log (aicFunction & " finished (" + CStr(Round(Timer - tStart, 2)) + "s).")
                    ExecuteAICellsFunction = True
                    Exit Function
                End If
            End If
            If PyQueueReturn(0) = "debug" Then
                Log (PyQueueReturn(1))
            End If
            If PyQueueReturn(0) = "progress" Then
                Log (PyQueueReturn(1))
            End If
        End If
    Loop
    
    Log ("Aborted.")
    ExecuteAICellsFunction = False
    Exit Function
valueError:
    Log ("ExecuteAICellsFunction Error")
    ExecuteAICellsFunction = False
End Function

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    Debug.Print ("UserForm_QueryClose")
    abort = True
End Sub


