VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CAptOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Dim mOptions() As String
Dim bNoOption As Boolean

Private Sub Class_Initialize()
   bNoOption = True
End Sub

Public Sub Add(sOpt As String)
Dim iCurInd As Long
   If bNoOption = True Then
      ReDim mOptions(0) As String
      iCurInd = 0
      bNoOption = False
   Else
      iCurInd = UBound(mOptions) + 1
      ReDim Preserve mOptions(iCurInd) As String
   End If
   mOptions(iCurInd) = sOpt
End Sub

Public Function Item(i As Long) As String
   If i >= 0 Then
      If bNoOption = False Then
         If i <= UBound(mOptions) Then
            Item = mOptions(i)
         End If
      End If
   End If
End Function

Public Property Get Length() As Long
   If bNoOption = True Then
      Length = 0
   Else
      Length = UBound(mOptions) + 1
   End If
End Property


