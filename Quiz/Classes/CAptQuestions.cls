VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CAptQuestions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"CAptQuestion"
Attribute VB_Ext_KEY = "Member0" ,"CAptQuestion"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mCol As Collection
Private iCIndex As Long

Public Function Add(AnswerIndex As Long, Options As CAptOptions, Question As String, Optional sKey As Long) As CAptQuestion
   
   Dim objNewMember As CAptQuestion
   Set objNewMember = New CAptQuestion
   
   objNewMember.AnswerIndex = AnswerIndex
   If IsObject(Options) Then
      Set objNewMember.Options = Options
   Else
      objNewMember.Options = Options
   End If
   objNewMember.Question = Question
   mCol.Add objNewMember, "Key" & sKey

   Set Add = objNewMember
   Set objNewMember = Nothing
End Function

Public Function AddQuestion(oQuestion As CAptQuestion) As CAptQuestion
   mCol.Add oQuestion, "Key" & iCIndex
   iCIndex = iCIndex + 1
   Set AddQuestion = oQuestion
End Function

Public Property Get Item(iCInd As Long) As CAptQuestion
Attribute Item.VB_UserMemId = 0
   Set Item = mCol("Key" & iCInd)
End Property

Public Property Get Count() As Long
   Count = mCol.Count
End Property

Public Sub Remove(sKey As Long)
   mCol.Remove "Key" & sKey
   'iCIndex = iCIndex - 1
End Sub

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
   Set NewEnum = mCol.[_NewEnum]
End Property

Private Sub Class_Initialize()
   Set mCol = New Collection
End Sub

Private Sub Class_Terminate()
   Set mCol = Nothing
End Sub

