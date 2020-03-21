VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
#If Win64 Then
    Private Const ExcelDnaXLLName As String = "ExcelDna.IntelliSense64-1.1.0.xll"
#Else
    Private Const ExcelDnaXLLName As String = "ExcelDna.IntelliSense-1.1.0.xll"
#End If

#If VBA7 Then
    #If Win64 Then
        Private Const ExcelWin64VBA7 As Boolean = True
    #Else
        Private Const ExcelWin64VBA7 As Boolean = False
    #End If
#Else
    Private Const ExcelWin64VBA7 As Boolean = False
#End If


Private Sub Workbook_Open()
    On Error Resume Next
    
    If Not ExcelWin64VBA7 Then
        MsgBox "AIcells is only supported on 64 bit Office 365 and Windows."
        Exit Sub
    End If
    
    Application.RegisterXLL (Application.ThisWorkbook.Path + "\" + ExcelDnaXLLName)
    
    Call SetMacroOptions
End Sub

