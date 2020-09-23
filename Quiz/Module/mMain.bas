Attribute VB_Name = "mMain"
Option Explicit


Public mMainWnd      As frmLoadWnd
Public DOM           As DOMDocument
Public mAptTest      As frmQuizWnd

Public Const MAX_MINUTE = 30

Public Sub Main()
   
On Error GoTo ErrMain
   
   AddXPStyle          '   Adding current OS Style
   Set mMainWnd = New frmLoadWnd
   frmLoadWnd.Show
   Exit Sub
ErrMain:
   MsgBox "Unable to proceed", vbInformation
End Sub

Public Function DomAptitudeToCQuestion(xmlAptElement As IXMLDOMElement) As CAptQuestion
   Dim AptQuestion   As CAptQuestion
   Dim qNode         As IXMLDOMElement
   Dim optNodeList   As IXMLDOMNodeList
   Dim AnsNode       As IXMLDOMElement
   
   If xmlAptElement Is Nothing Then
      Exit Function
   End If
   
   Set AptQuestion = New CAptQuestion
   
   Dim i As Integer
   Set qNode = xmlAptElement.getElementsByTagName("Question").Item(0)
   AptQuestion.Question = qNode.Text
   Set optNodeList = xmlAptElement.getElementsByTagName("Option")
   
   For i = 0 To optNodeList.Length - 1
      AptQuestion.Options.Add optNodeList.Item(i).Text
   Next
   
   If xmlAptElement.getElementsByTagName("Answer").Length > 0 Then
      Set AnsNode = xmlAptElement.getElementsByTagName("Answer").Item(0)
      AptQuestion.AnswerIndex = Val(AnsNode.Text)
   End If
   Set DomAptitudeToCQuestion = AptQuestion
End Function

Public Function DomRootToCQuestions(xmlRootElement As IXMLDOMElement) As CAptQuestions
   Dim AptQuestions  As CAptQuestions
   Dim qNode         As IXMLDOMElement
   Dim optNodeList   As IXMLDOMNodeList
   Dim AnsNode       As IXMLDOMElement
   Dim AptQuestion   As CAptQuestion
   
   If xmlRootElement Is Nothing Then
      Exit Function
   End If
   
   Set AptQuestions = New CAptQuestions
   
   Dim i As Integer
   
   Set optNodeList = xmlRootElement.getElementsByTagName("Aptitude")
   
   For i = 0 To optNodeList.Length - 1
      Set AptQuestion = DomAptitudeToCQuestion(optNodeList.Item(i))
      If Not AptQuestion Is Nothing Then
         AptQuestions.AddQuestion AptQuestion
      End If
   Next
   Set DomRootToCQuestions = AptQuestions
End Function

