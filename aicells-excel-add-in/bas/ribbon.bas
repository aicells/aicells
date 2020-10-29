Attribute VB_Name = "ribbon"
Option Explicit

'Callback for ButtonFixUDFPaths onAction
Sub onActionFixUDFPaths(control As IRibbonControl)
    Dim x
    If ActiveWorkbook Is Nothing Then
        MsgBox "Open a workbook with AIcells UDFs first."
    Else
        x = Application.Run("aicells.xlam!FixUDFPaths", ActiveWorkbook)
    End If
End Sub

'Callback for ButtonRunFunction onAction
Sub onActionRunFunction(control As IRibbonControl)
    Dim referencedRange As range
    
    UserFormFunctionRunner.CommandButtonRun.Enabled = True
    UserFormFunctionRunner.CommandButtonAbort.Enabled = False
    
    If Selection.Columns.Count = 1 Then
        If Selection.Rows.Count = 1 And Selection.Columns.Count = 1 Then
            If Selection.cells(1, 1).HasFormula Then
                Set referencedRange = DecodeRangeReference(Selection.cells(1, 1))
                If Not (referencedRange Is Nothing) Then
                    UserFormFunctionRunner.TextBoxLog.value = "Selected parameter range: " & Replace(referencedRange.address(External:=True), "$", "") + vbCr
                Else
                    UserFormFunctionRunner.TextBoxLog.value = "ERROR: Invalid parameter range selected."
                    Exit Sub
                End If
            End If
        Else
            ' single column, multiple rows
            UserFormFunctionRunner.TextBoxLog.value = "Selected function list range (" & CStr(Selection.Rows.Count) & " functions): " & Replace(Selection.address(External:=True), "$", "") + vbCr
        End If
        
        
    Else
        UserFormFunctionRunner.TextBoxLog.value = "Selected parameter range: " & Replace(Selection.address(External:=True), "$", "") + vbCr
    End If
    
    

    
    'UserFormFunctionRunner.Show(VBA.FormShowConstants.vbModeless)
    Call UserFormFunctionRunner.Show
End Sub

'Callback for ButtonShowErrors onAction
Sub onActionShowErrors(control As IRibbonControl)
End Sub

