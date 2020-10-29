Attribute VB_Name = "aicells"
Option Explicit

Private Const ShowErrorsAsComments = True

Dim processParameterRangesValueError As Boolean

Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr
Declare PtrSafe Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As LongPtr, lpdwProcessId As Long) As Long
Declare PtrSafe Function GetCurrentProcessId Lib "kernel32" () As Long

#If Win64 Then
    Declare PtrSafe Function GetWindowLongPtr Lib "user32" Alias "GetWindowLongPtrA" (ByVal hwnd As LongPtr, ByVal nIndex As Long) As LongPtr
#Else
    Declare PtrSafe Function GetWindowLongPtr Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As LongPtr, ByVal nIndex As Long) As LongPtr
#End If

' Window field offsets for GetWindowLong() and GetWindowWord()
Const GWL_WNDPROC = (-4)
Const GWL_HINSTANCE = (-6)
Const GWL_HWNDPARENT = (-8)
Const GWL_STYLE = (-16)
Const GWL_EXSTYLE = (-20)
Const GWL_USERDATA = (-21)
Const GWL_ID = (-12)

' Window Styles
Const WS_OVERLAPPED = &H0&
Const WS_POPUP = &H80000000
Const WS_CHILD = &H40000000
Const WS_MINIMIZE = &H20000000
Const WS_VISIBLE = &H10000000
Const WS_DISABLED = &H8000000
Const WS_CLIPSIBLINGS = &H4000000
Const WS_CLIPCHILDREN = &H2000000
Const WS_MAXIMIZE = &H1000000
Const WS_CAPTION = &HC00000                  '  WS_BORDER Or WS_DLGFRAME
Const WS_BORDER = &H800000
Const WS_DLGFRAME = &H400000
Const WS_VSCROLL = &H200000
Const WS_HSCROLL = &H100000
Const WS_SYSMENU = &H80000
Const WS_THICKFRAME = &H40000
Const WS_GROUP = &H20000
Const WS_TABSTOP = &H10000

Const WS_MINIMIZEBOX = &H20000
Const WS_MAXIMIZEBOX = &H10000

Const WS_TILED = WS_OVERLAPPED
Const WS_ICONIC = WS_MINIMIZE
Const WS_SIZEBOX = WS_THICKFRAME
Const WS_OVERLAPPEDWINDOW = (WS_OVERLAPPED Or WS_CAPTION Or WS_SYSMENU Or WS_THICKFRAME Or WS_MINIMIZEBOX Or WS_MAXIMIZEBOX)
Const WS_TILEDWINDOW = WS_OVERLAPPEDWINDOW

Public Function LogUDFCall(UDFName, callerRange)
    Debug.Print (Format(Now, "yyyy-mm-dd hh:mm:ss") & " " & UDFName & ": " & callerRange.address(External:=True))
End Function

Private Function IsWindowClassOpen(c) As Boolean
    Dim windowPID, currentPID As Long
    Dim hwnd As LongPtr
    Dim isResizable As Boolean
    
    IsWindowClassOpen = False
    
    currentPID = GetCurrentProcessId
    
    hwnd = FindWindow(c, vbNullString)
    If hwnd = 0 Then: Exit Function
    
    ' Formula Argunments dialog is not resizable
    isResizable = ((GetWindowLongPtr(hwnd, GWL_STYLE) And WS_SIZEBOX) <> 0)
    If isResizable Then: Exit Function
        
    Call GetWindowThreadProcessId(hwnd, windowPID)
    
    If currentPID = windowPID Then: IsWindowClassOpen = True
End Function

Public Function IsFXWindowOpen() As Boolean
    IsFXWindowOpen = IsWindowClassOpen("bosa_sdm_XL9")
    
    ' IsFXWindowOpen = False
    ' If (Not Application.CommandBars("Standard").Controls(1).Enabled) Then IsFXWindowOpen = True
    
End Function

Public Function RunFunction(r As range, recalculate As Boolean)
    'Call UserFormFunctionRunner.SetMacroMode(r, recalculate)
    r.Select
    Call UserFormFunctionRunner.Show
End Function


Public Function aicRangeReference(r As range, Optional add_hash = True)
    Dim formula, rangeStr, rangeStrSheet, rangeStrRC As String
    Dim callerRange
    Dim i, j As Long
    Dim rightCut As Long
    
    Set callerRange = Application.Caller
    formula = callerRange.cells(1, 1).formula
    If Mid$(formula, 1, 19) <> "=aicRangeReference(" Then GoTo valueError
    
    rightCut = 0
    
    If Right$(formula, 6) = "FALSE)" Then: rightCut = 7
    If Right$(formula, 5) = "TRUE)" Then: rightCut = 5
        
    rangeStr = Mid$(formula, 20, Len(formula) - 20 - rightCut)
        
    j = InStrRev(rangeStr, "[")
    If j = 0 Then
        ' not table header reference
        i = InStrRev(rangeStr, "!")
        If i <> 0 Then
            rangeStrSheet = Mid$(rangeStr, 1, i - 1)
            rangeStrRC = Mid$(rangeStr, i + 1, Len(rangeStr) - 3)
    
            If left$(rangeStrSheet, 1) = "'" Then
                rangeStrSheet = Mid$(rangeStr, 2, Len(rangeStrSheet) - 2)
            End If
    
            rangeStrSheet = Replace(rangeStrSheet, "''", "'")
    
            rangeStr = rangeStrSheet & "!" & rangeStrRC
        End If
    Else
        rangeStr = Replace(rangeStr, "''", "'")
    End If
    
        
'    Debug.Print formula
'    Debug.Print rangeStr

    If add_hash Then
        aicRangeReference = rangeStr & "  {" & Right("00000000" & Hex(Int(Rnd * 2147483647)), 8) & "}"
    Else
        aicRangeReference = rangeStr
    End If
        
    Exit Function
valueError:
    Set aicRangeReference = Nothing
    Exit Function
End Function

Public Function aicDebugRangeReference(r As range)
    Dim addressAsText As String
    
    If Not r.HasFormula Then GoTo valueError

    addressAsText = r.cells(1, 1).text
    
    If Mid$(r.cells(1, 1).formula, 1, 19) <> "=aicRangeReference(" Then GoTo valueError

    On Error GoTo valueError
    Set r = RangeFromAddress(addressAsText)
    
    aicDebugRangeReference = r.Parent.name & "!" & r.address(External:=False)
    
    Exit Function
valueError:
    Set aicDebugRangeReference = Nothing
    Exit Function
End Function


Private Function FixUDFPaths(wb As Workbook)
    Dim aicPath As String
    Dim links
    Dim xlamName As String
    Dim i As Long
    
    'Debug.Print IsFXWindowOpen()
    
    aicPath = Application.ThisWorkbook.Path + "\" + Application.ThisWorkbook.name
    
    links = wb.LinkSources(xlExcelLinks)
    If Not IsEmpty(links) Then
      For i = 1 To UBound(links)
        xlamName = UCase(Right(links(i), Len(links(i)) - InStrRev(links(i), "\")))
        
        If xlamName = "AICELLS.XLAM" Then wb.ChangeLink links(i), aicPath, xlExcelLinks
        
      Next
    End If
    
    MsgBox "The path of aicells.xlam was successfully updated." + vbLf + "Excel is now in manual calculation mode"
End Function

Public Function CountRangeErrors(cells As range) As Long
   Dim cell As range
   For Each cell In cells
      If Application.WorksheetFunction.isError(cell) Then CountRangeErrors = CountRangeErrors + 1
   Next
End Function

Public Function HasRangeErrors(cells) As Boolean
   Dim cell As range
   HasRangeErrors = False
   For Each cell In cells
      If Application.WorksheetFunction.isError(cell) Then
        HasRangeErrors = True
        Exit Function
      End If
   Next
End Function

Function DeleteErrorMessage(callerRange As range)
    If ShowErrorsAsComments Then: callerRange.ClearComments
End Function

Function CheckIfError(ret)
    
    CheckIfError = False
    If TypeName(ret) = "Variant()" Then
        On Error GoTo endFunction
        If UBound(ret, 1) >= 1 And UBound(ret, 2) = 1 Then
            If ret(0, 0) = "#AICELLS-ERROR!" And ret(0, 1) = "#AICELLS-ERROR@" Then
                CheckIfError = True
            End If
        End If
        On Error GoTo 0
    End If
endFunction:
End Function

Public Sub ShowErrors(e, callerRange As range)
    Dim i As Long
    Debug.Print callerRange.Parent.name & "!" & Replace(callerRange.address(External:=False), "$", "")
    Debug.Print "---"
    
    
    Dim msg As String
    
    For i = 1 To UBound(e, 1)
        Debug.Print e(i, 0)
        Debug.Print e(i, 1)
        Debug.Print "---"
        msg = msg + e(i, 1) + " (" + e(i, 0) + ")" + vbLf
    Next i
    
    If ShowErrorsAsComments Then
        callerRange.AddComment msg
        callerRange.Comment.Shape.width = 300
        callerRange.Comment.Shape.TextFrame.AutoSize = True
    End If
End Sub

Public Function DecodeRangeReference(r As range)
    Dim addressAsText As String
    
    If Not r.HasFormula Then GoTo valueError

    addressAsText = r.cells(1, 1).text
    
    If Mid$(r.cells(1, 1).formula, 1, 19) <> "=aicRangeReference(" Then GoTo valueError

    On Error GoTo valueError
    Set DecodeRangeReference = RangeFromAddress(addressAsText)
    
    Exit Function
valueError:
    Set DecodeRangeReference = Nothing
    Exit Function
End Function

Function ProcessParameterRanges2(pb As PyParameterBuilder, parameters, namespace As String) As Boolean
    ProcessParameterRanges2 = True
    
    processParameterRangesValueError = False
    Call ProcessParameterRanges(pb, parameters, 0, namespace)
    
    If processParameterRangesValueError Then
        ProcessParameterRanges2 = False
    End If
End Function

Private Sub ProcessParameterRanges(pb As PyParameterBuilder, r, level As Long, namespace As String)
    Dim x, y As Long
    Dim referencedRange As range
    
    ' scan the inpur range for errors
    If HasRangeErrors(r) Then processParameterRangesValueError = True
    If processParameterRangesValueError Then Exit Sub
    
    If level >= 5 Then Exit Sub
    If pb.GetProcessParameterRangesPointer() >= 10 + 3 Then Exit Sub
            
    pb.StoreRange namespace, r

    If r.Rows.Count > 100 Then Exit Sub
    If r.Columns.Count > 100 Then Exit Sub
    
    ' when the range has 2 columns and less than 100 rows, search for references
    If r.Columns.Count = 2 Then
        For y = 1 To r.Rows.Count
            If r.cells(y, 2).HasFormula Then
                Set referencedRange = DecodeRangeReference(r.cells(y, 2))
                If Not (referencedRange Is Nothing) Then
                    Call ProcessParameterRanges(pb, referencedRange, level + 1, namespace + "." + r.cells(y, 1).value)
                End If
            End If
        Next y
    End If
    
    ' when the range has 2 rows and less than 100 columns, search for references
    If r.Rows.Count = 2 Then
        For x = 1 To r.Columns.Count
            If r.cells(2, x).HasFormula Then
                Set referencedRange = DecodeRangeReference(r.cells(2, x))
                If Not (referencedRange Is Nothing) Then
                    Call ProcessParameterRanges(pb, referencedRange, level + 1, namespace + "." + r.cells(1, x).value)
                End If
            End If
        Next x
    End If
    
End Sub

Function RangeFromAddress(address As String) As range
    Dim callerRange As range
    Dim callerWorkbook As Workbook
    Dim callerSheet, sh As Worksheet
    Dim ws As Worksheet
    Dim n As name
    Dim lo As ListObject
    Dim i, j, k As Long
    Dim addressSheet As String
    Dim addressRange As String
    Dim listRange As range
    Dim addressTableName, addressHeaderName As String
    
    If TypeOf Application.Caller Is range Then
        Set callerRange = Application.Caller
    Else
        Set callerRange = Selection
    End If
    Set callerSheet = callerRange.Worksheet
    Set callerWorkbook = callerSheet.Parent
    
    k = InStrRev(address, "{")
    If k <> 0 Then: address = left(address, Len(address) - 12)
    
    'Debug.Print (address)
    
    i = InStrRev(address, "!")
    j = InStr(address, "[")
    
    If j <> 0 Then
        ' table column
        On Error Resume Next
        
        addressTableName = Mid$(address, 1, j - 1)
        addressHeaderName = Mid$(address, j + 1, Len(address) - j - 1)
        
        
        For Each ws In callerWorkbook.Sheets
            On Error Resume Next
            Set lo = ws.ListObjects(addressTableName)
            If Not (lo Is Nothing) Then
                ' body
                If addressHeaderName = "#Data" Then
                    Set RangeFromAddress = lo.DataBodyRange
                    Exit Function
                End If
                
                ' entire column header row
                If addressHeaderName = "#Headers" Then
                    Set RangeFromAddress = lo.HeaderRowRange
                    Exit Function
                End If
                
                ' entire column header row, body, total row
                If addressHeaderName = "#All" Then
                    Set RangeFromAddress = lo.range
                    Exit Function
                End If
            
                ' entire total row
                If addressHeaderName = "#Totals" Then
                    Set RangeFromAddress = lo.TotalsRowRange
                    Exit Function
                End If
            
                ' single column, only include the body
                Set RangeFromAddress = lo.ListColumns(addressHeaderName).DataBodyRange
                Exit Function
            End If
            On Error GoTo 0
        Next ws
    End If
    
    If i = 0 Then
    
        ' named range ?
        On Error Resume Next
        Set n = callerWorkbook.Names(address)
        If Not (n Is Nothing) Then
            Set RangeFromAddress = n.RefersToRange
            Exit Function
        End If
        On Error GoTo 0
        
        ' table ?
        ' body only
        For Each ws In callerWorkbook.Sheets
            On Error Resume Next
            Set lo = ws.ListObjects(address)
            If Not (lo Is Nothing) Then
                Set RangeFromAddress = lo.DataBodyRange
                Exit Function
            End If
            On Error GoTo 0
        Next ws
        
        ' unidentified name
        'Set RangeFromAddress = Nothing
        
        Set RangeFromAddress = callerSheet.range(address)
        Exit Function
    End If
    
    On Error GoTo returnNothing
    addressSheet = left$(address, i - 1)
    addressRange = Mid$(address, i + 1)

    Set sh = callerWorkbook.Sheets(addressSheet)
    If Not (sh Is Nothing) Then
        Set RangeFromAddress = sh.range(addressRange)
        Exit Function
    End If
    
returnNothing:
    Set RangeFromAddress = Nothing
    Exit Function
End Function





