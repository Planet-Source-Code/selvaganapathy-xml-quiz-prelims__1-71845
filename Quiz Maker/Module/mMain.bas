Attribute VB_Name = "mMain"
Option Explicit

Public mMainWnd      As frmQuizMaker

Public Sub Main()
   
On Error GoTo ErrMain
   
   AddXPStyle          '   Adding current OS Style
   Set mMainWnd = New frmQuizMaker
   
   mMainWnd.Show
   Exit Sub
ErrMain:
   MsgBox "Unable to proceed", vbInformation
End Sub
Public Function ShowAddQuestion(Dom As DOMDocument, oDomEle As IXMLDOMNode, frmOwner As Form) As Boolean
   
   Dim fAddQuestion As frmAddQuestion
   
   Set fAddQuestion = New frmAddQuestion
   Set fAddQuestion.xmlAptitudeElement = oDomEle
   Set fAddQuestion.xmlDOMDoc = Dom
   
   fAddQuestion.Show vbModal, frmOwner
   If fAddQuestion.iResult = vbCancel Then
      ShowAddQuestion = False
      Unload fAddQuestion
      Exit Function
   End If
   ShowAddQuestion = True
   Unload fAddQuestion
End Function

Public Function ShowEditQuestion(Dom As DOMDocument, oDomEditEle As IXMLDOMElement, oDomReplace As IXMLDOMElement, frmOwner As Form) As Boolean
   
   Dim fAddQuestion As frmAddQuestion
   
   Set fAddQuestion = New frmAddQuestion
   Set fAddQuestion.xmlAptShowElement = oDomEditEle
   Set fAddQuestion.xmlAptitudeElement = oDomReplace
   Set fAddQuestion.xmlDOMDoc = Dom
   fAddQuestion.ElementToControl
   fAddQuestion.Show vbModal, frmOwner
   If fAddQuestion.iResult = vbCancel Then
      ShowEditQuestion = False
      Unload fAddQuestion
      Exit Function
   End If
   ShowEditQuestion = True
   Unload fAddQuestion
End Function

