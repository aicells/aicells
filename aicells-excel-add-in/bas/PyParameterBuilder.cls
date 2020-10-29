VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PyParameterBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pValueArr(12) As Variant
Private pNamespaceArr(9) As Variant
Private pProcessParameterRangesPointer As Long
Private pRangeArr(9) As Variant

Public Sub Init(UDFName)
    pValueArr(0) = UDFName
    pProcessParameterRangesPointer = 0
End Sub

Public Sub SetUdfArguments(args)
    pValueArr(1) = args
End Sub

Public Sub StoreRange(name As String, range)
    pNamespaceArr(pProcessParameterRangesPointer) = name
    Set pRangeArr(pProcessParameterRangesPointer) = range
    pValueArr(pProcessParameterRangesPointer + 3) = range.value
    pProcessParameterRangesPointer = pProcessParameterRangesPointer + 1
End Sub

Public Sub StoreArray(name As String, arr)
    pNamespaceArr(pProcessParameterRangesPointer) = name
    pValueArr(pProcessParameterRangesPointer + 3) = arr
    pProcessParameterRangesPointer = pProcessParameterRangesPointer + 1
End Sub

Public Function GetParameterArray()
    pValueArr(2) = pNamespaceArr
    GetParameterArray = pValueArr
End Function

Public Function GetProcessParameterRangesPointer()
    GetProcessParameterRangesPointer = pProcessParameterRangesPointer
End Function

Public Function GetRangeByName(name As String) As Variant
    Dim i As Long
    
    Set GetRangeByName = Nothing
    
    For i = 0 To 9
        If pNamespaceArr(i) = name Then
            Set GetRangeByName = pRangeArr(i)
            Exit Function
        End If
    Next i

End Function

