VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmQuizMaker 
   Caption         =   "Form1"
   ClientHeight    =   6360
   ClientLeft      =   120
   ClientTop       =   750
   ClientWidth     =   8880
   Icon            =   "frmAptMaker.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6360
   ScaleWidth      =   8880
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picSplitter 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      FillColor       =   &H00808080&
      Height          =   4800
      Left            =   5400
      ScaleHeight     =   2090.126
      ScaleMode       =   0  'User
      ScaleWidth      =   780
      TabIndex        =   5
      Top             =   285
      Visible         =   0   'False
      Width           =   72
   End
   Begin VB.PictureBox picQuestion 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   4815
      Left            =   2580
      ScaleHeight     =   4785
      ScaleWidth      =   7005
      TabIndex        =   6
      Top             =   360
      Width           =   7035
      Begin VB.TextBox txtOptions 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   0
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   12
         Top             =   2040
         Width           =   1875
      End
      Begin VB.TextBox txtQuestion 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   0
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   11
         Top             =   300
         Width           =   1875
      End
      Begin VB.Label lblAnswerIndex 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FEFAF5&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   3900
         TabIndex        =   10
         Top             =   3480
         Width           =   2355
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblCapAnswerIndex 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0057D971&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "&Answer Index"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   0
         TabIndex        =   9
         Top             =   3480
         Width           =   3840
      End
      Begin VB.Label lblCapOptions 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0057D971&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "&Options"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   0
         TabIndex        =   8
         Top             =   1740
         Width           =   6285
      End
      Begin VB.Label lblCapQuestion 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0057D971&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "&Question"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   0
         TabIndex        =   7
         Top             =   0
         Width           =   6255
      End
   End
   Begin VB.PictureBox picTitles 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   0
      ScaleHeight     =   300
      ScaleWidth      =   8880
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   0
      Width           =   8880
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         BackColor       =   &H00E9AA4B&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Questions"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   0
         Left            =   0
         TabIndex        =   3
         Tag             =   " TreeView:"
         Top             =   12
         Width           =   2016
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         BackColor       =   &H00E9AA4B&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "View"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   1
         Left            =   2078
         TabIndex        =   2
         Tag             =   " ListView:"
         Top             =   12
         Width           =   3216
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   6540
      TabIndex        =   0
      Top             =   3600
      Width           =   1515
   End
   Begin MSComctlLib.ImageList imlToolbarIcons 
      Left            =   6960
      Top             =   1500
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAptMaker.frx":0442
            Key             =   "Back"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAptMaker.frx":0554
            Key             =   "Forward"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAptMaker.frx":0666
            Key             =   "Cut"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAptMaker.frx":0778
            Key             =   "Copy"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAptMaker.frx":088A
            Key             =   "Paste"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAptMaker.frx":099C
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAptMaker.frx":0AAE
            Key             =   "Properties"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAptMaker.frx":0BC0
            Key             =   "View Large Icons"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAptMaker.frx":0CD2
            Key             =   "View Small Icons"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAptMaker.frx":0DE4
            Key             =   "View List"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAptMaker.frx":0EF6
            Key             =   "View Details"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView tvTreeView 
      Height          =   4800
      Left            =   0
      TabIndex        =   4
      Top             =   300
      Width           =   2550
      _ExtentX        =   4498
      _ExtentY        =   8467
      _Version        =   393217
      Indentation     =   353
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComDlg.CommonDialog dlgCommonDialog 
      Left            =   900
      Top             =   5580
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Image imgSplitter 
      Height          =   4788
      Left            =   1965
      MousePointer    =   9  'Size W E
      Top             =   285
      Width           =   150
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileNew 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuFileSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuSave 
         Caption         =   "&Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFileSaveAs 
         Caption         =   "Save &As"
      End
      Begin VB.Menu mnuFileSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileClose 
         Caption         =   "&Close"
      End
   End
   Begin VB.Menu mnuQuestion 
      Caption         =   "&Question"
      Begin VB.Menu mnuAddQuestion 
         Caption         =   "&Add Question"
         Shortcut        =   ^D
      End
      Begin VB.Menu mnuEditQuestion 
         Caption         =   "&Edit Question"
         Shortcut        =   ^E
      End
      Begin VB.Menu mnuRemoveQuestion 
         Caption         =   "&Remove Question"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu mnuRemoveAll 
         Caption         =   "&Remove All"
         Shortcut        =   +{DEL}
      End
   End
End
Attribute VB_Name = "frmQuizMaker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mbMoving As Boolean
Const sglSplitLimit = 500
Dim FileName As String

Dim Dom As DOMDocument
Dim Root As IXMLDOMElement
   
Dim bDirty As Boolean
   
Private Sub Form_Load()
   NewFile
End Sub

Private Sub Form_Resize()
   On Error Resume Next
   If Me.Width < 8000 Then Me.Width = 8000
   If Me.Height < 7000 Then Me.Height = 7000
   SizeControls imgSplitter.Left
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If bDirty = True Then
      Select Case MsgBox("Save changes. Do you want to continue ?", vbInformation + vbYesNoCancel)
      Case vbYes
         mnuSave_Click
      Case vbNo
      Case vbCancel
         Cancel = True
         Exit Sub
      End Select
   End If
End Sub

Private Sub imgSplitter_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   With imgSplitter
      picSplitter.Move .Left, .Top, .Width \ 2, .Height - 20
   End With
   picSplitter.Visible = True
   mbMoving = True
End Sub

Private Sub imgSplitter_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   Dim sglPos As Single
   If mbMoving Then
      sglPos = x + imgSplitter.Left
      If sglPos < sglSplitLimit Then
         picSplitter.Left = sglSplitLimit
      ElseIf sglPos > Me.Width - sglSplitLimit Then
         picSplitter.Left = Me.Width - sglSplitLimit
      Else
         picSplitter.Left = sglPos
      End If
   End If
End Sub

Private Sub imgSplitter_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
   SizeControls picSplitter.Left
   picSplitter.Visible = False
   mbMoving = False
End Sub

Sub SizeControls(x As Single)
   On Error Resume Next
   Dim SIZE_T As Integer
   SIZE_T = 2500
   If x < SIZE_T Then x = SIZE_T
   If x > (Me.Width - SIZE_T) Then x = Me.Width - SIZE_T
   tvTreeView.Width = x
   imgSplitter.Left = x
   picQuestion.Left = x + 40
   picQuestion.Width = Me.Width - (tvTreeView.Width + 140)
   lblTitle(0).Width = tvTreeView.Width
   lblTitle(1).Left = picQuestion.Left
   lblTitle(1).Width = picQuestion.Width - 40
   picQuestion.Top = tvTreeView.Top
   tvTreeView.Height = Me.Height - picTitles.Height - 450
   picQuestion.Height = tvTreeView.Height
   imgSplitter.Top = tvTreeView.Top
   imgSplitter.Height = tvTreeView.Height
End Sub

Private Sub mnuAddQuestion_Click()
   Dim Nod As IXMLDOMNode
   
   Set Nod = Dom.createElement("Aptitude")
   If ShowAddQuestion(Dom, Nod, Me) = True Then
      Root.appendChild Nod
      LoadNodesFromRoot True
      bDirty = True
   End If
   
End Sub
Private Sub LoadNodesFromRoot(Optional bRootExpand As Boolean = False)
   Dim NodList As IXMLDOMNodeList
   Dim iNode As Node
   Dim i As Integer
   
   tvTreeView.Nodes.Clear
   Set NodList = Root.getElementsByTagName("Aptitude")
   Set iNode = tvTreeView.Nodes.Add(, , , Root.nodeName)
   For i = 0 To NodList.length - 1
      tvTreeView.Nodes.Add iNode, tvwChild, , NodList.Item(i).childNodes(0).nodeTypedValue
   Next
   iNode.Expanded = bRootExpand

End Sub
Private Sub mnuEditQuestion_Click()
   '
Dim xmlAptitudeNodeList As IXMLDOMNodeList
Dim xmlOptionsNodeList  As IXMLDOMNodeList
Dim xmlCurAptitudeNode  As IXMLDOMElement
Dim xmlQuestionNode     As IXMLDOMElement
Dim xmlOptionsNode      As IXMLDOMElement
Dim xmlAptitudeNode     As IXMLDOMElement
Dim sOption             As String
Dim sQuestion           As String
Dim sAnswer             As String
Dim i                   As Long
Dim j                   As Long
   
   If tvTreeView.SelectedItem Is Nothing Then Exit Sub
   
   Set xmlAptitudeNodeList = Root.selectNodes("//Aptitude")
   
   For i = 0 To xmlAptitudeNodeList.length - 1
      Set xmlCurAptitudeNode = xmlAptitudeNodeList.Item(i)
      Set xmlQuestionNode = xmlCurAptitudeNode.getElementsByTagName("Question").Item(0)
      
      If xmlQuestionNode.Text = tvTreeView.SelectedItem.Text Then
         Set xmlAptitudeNode = Dom.createElement("Aptitude")
         If ShowEditQuestion(Dom, xmlCurAptitudeNode, xmlAptitudeNode, Me) = True Then
            ClearPane
            Root.replaceChild xmlAptitudeNode, xmlCurAptitudeNode
            LoadNodesFromRoot True
            bDirty = True
         End If
         Exit For
      End If
   Next
End Sub

Private Sub mnuFileClose_Click()
   Unload Me
End Sub

Private Sub mnuFileNew_Click()
   NewFile
End Sub

Private Sub mnuFileOpen_Click()

   With dlgCommonDialog
      .DialogTitle = "Open"
      .CancelError = False
      .Filter = "XML File (*.qst)|*.qst"
      .Flags = cdlOFNFileMustExist Or cdlOFNPathMustExist Or cdlOFNExtensionDifferent
      .ShowOpen
      If Len(.FileName) = 0 Then
         Exit Sub
      End If
      FileName = .FileName
   End With
   Me.Caption = "[" & FileName & "]"
   If LoadFile(FileName) = False Then
      NewFile
   End If

End Sub

Private Sub mnuFileSaveAs_Click()

   With dlgCommonDialog
      .DialogTitle = "Save"
      .CancelError = False
      .Filter = "Question File (*.qst)|*.qst"
      .Flags = cdlOFNFileMustExist Or cdlOFNPathMustExist Or cdlOFNOverwritePrompt
      .ShowSave
      If Len(.FileName) = 0 Then
         Exit Sub
      End If
      FileName = .FileName
   End With

   Dom.save FileName
   Me.Caption = "[" & FileName & "]"
   bDirty = False

End Sub

Private Sub mnuRemoveAll_Click()
   
Dim xmlAptitudeNodeList As IXMLDOMNodeList
Dim i                   As Long
   
   If MsgBox("Are you sure want to remove this question?", vbCritical + vbYesNo, "Delete") = vbYes Then
      If tvTreeView.SelectedItem Is Nothing Then Exit Sub
      Set xmlAptitudeNodeList = Root.selectNodes("//Aptitude")
      For i = 0 To xmlAptitudeNodeList.length - 1
         Root.removeChild xmlAptitudeNodeList.Item(i)
         bDirty = True
      Next
      LoadNodesFromRoot True
      ClearPane
      
   End If
End Sub

Private Sub mnuRemoveQuestion_Click()

Dim xmlAptitudeNodeList As IXMLDOMNodeList
Dim xmlCurAptitudeNode  As IXMLDOMElement
Dim xmlQuestionNode     As IXMLDOMElement
Dim i                   As Long
   
   If MsgBox("Are you sure want to remove this question?", vbCritical + vbYesNo, "Delete") = vbYes Then
      If tvTreeView.SelectedItem Is Nothing Then Exit Sub
      
      Set xmlAptitudeNodeList = Root.selectNodes("//Aptitude")
      
      For i = 0 To xmlAptitudeNodeList.length - 1
         Set xmlCurAptitudeNode = xmlAptitudeNodeList.Item(i)
         Set xmlQuestionNode = xmlCurAptitudeNode.getElementsByTagName("Question").Item(0)
         If xmlQuestionNode.Text = tvTreeView.SelectedItem.Text Then
            Root.removeChild xmlCurAptitudeNode
            LoadNodesFromRoot True
            bDirty = True
            ClearPane
            Exit For
         End If
      Next
   End If
End Sub

Private Sub mnuSave_Click()
   If FileName = "" Then
      With dlgCommonDialog
         .DialogTitle = "Save"
         .CancelError = False
         .Filter = "Question File (*.qst)|*.qst"
         .Flags = cdlOFNFileMustExist Or cdlOFNPathMustExist Or cdlOFNOverwritePrompt
         .ShowSave
         If Len(.FileName) = 0 Then
            Exit Sub
         End If
         FileName = .FileName
      End With
   End If
   'If dlgCommonDialog.FilterIndex = 0 Then
      Dom.save FileName
   'Else
      'SaveASDB
   'End If
   Me.Caption = "[" & FileName & "]"
   bDirty = False

End Sub

Private Sub picQuestion_Resize()
   On Error Resume Next
   lblCapQuestion.Move 0, 0, picQuestion.Width - 100
   txtQuestion.Move 0, lblCapQuestion.Height + lblCapQuestion.Top, picQuestion.Width - 100, picQuestion.Height / 3
   lblCapOptions.Move 0, txtQuestion.Top + txtQuestion.Height, picQuestion.Width - 100 ', picQuestion.Height / 2
   txtOptions.Move 0, lblCapOptions.Height + lblCapOptions.Top, picQuestion.Width - 100, picQuestion.Height / 3
   lblCapAnswerIndex.Move 0, txtOptions.Height + txtOptions.Top + 20, 2 * picQuestion.Width / 3
   lblAnswerIndex.Move lblCapAnswerIndex.Left + lblCapAnswerIndex.Width + 10, lblCapAnswerIndex.Top + 10, picQuestion.Width - lblCapAnswerIndex.Width - 100
End Sub

Private Sub tvTreeView_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   If Button = vbRightButton Then
      PopupMenu mnuQuestion
   End If
End Sub

Private Function LoadFile(sFileName As String, Optional Filter As String = "qst") As Boolean
   Dim i As Integer
On Error GoTo LoadErr
   If Filter = "qst" Then
      Dom.Load (sFileName)
      tvTreeView.Nodes.Clear
      Set Root = Dom.documentElement
      If Root.nodeName = "Test" Then
         LoadNodes Root, True
         LoadFile = True
      Else
         LoadFile = False
         MsgBox "Unable to Load File...", vbCritical, "Loading Error"
         Exit Function
      End If
   
   ElseIf Filter = "mdb" Then
      'LoadDBToXML
   Else
      LoadFile = False
      MsgBox "Unable to Load File...", vbCritical, "Loading Error"
      Exit Function
   End If
   
   Exit Function
LoadErr:
   MsgBox "Error : " & Err.Description, vbCritical, " Error "
   LoadFile = False
End Function

Private Sub NewFile(Optional bIgnore As Boolean = False)

If bDirty = True And bIgnore = False Then
   Select Case MsgBox("Save changes. Do you want to continue ?", vbYesNo + vbInformation)
   Case vbYes
      Dom.save FileName
   Case vbNo
   End Select
End If
   bDirty = False
   FileName = ""
   Set Dom = New DOMDocument
   Set Root = Dom.createElement("Test")
   Dom.appendChild Root
   tvTreeView.Nodes.Clear
   tvTreeView.Nodes.Add , , , "Test"
   Me.Caption = "[Untitled]"
End Sub

Private Sub LoadNodes(Rt As IXMLDOMNode, Optional bExpand As Boolean = False)
Dim iNode As Node, i As Integer
   Set iNode = tvTreeView.Nodes.Add(, , , Rt.nodeName)
   For i = 0 To Root.childNodes.length - 1
      tvTreeView.Nodes.Add iNode, tvwChild, , Rt.childNodes(i).childNodes(0).Text
   Next
   iNode.Expanded = bExpand
End Sub

Private Sub tvTreeView_NodeClick(ByVal Node As MSComctlLib.Node)

Dim xmlAptitudeNodeList As IXMLDOMNodeList
Dim xmlOptionsNodeList  As IXMLDOMNodeList
Dim xmlCurAptitudeNode  As IXMLDOMElement
Dim xmlQuestionNode     As IXMLDOMElement
Dim xmlOptionsNode      As IXMLDOMElement
Dim xmlAnswerNode       As IXMLDOMElement
Dim sOption             As String
Dim sQuestion           As String
Dim sAnswer             As String
Dim i                   As Long
Dim j                   As Long
   
   If tvTreeView.SelectedItem Is Nothing Then Exit Sub
   
   Set xmlAptitudeNodeList = Root.selectNodes("//Aptitude")
   ClearPane
   For i = 0 To xmlAptitudeNodeList.length - 1
      Set xmlCurAptitudeNode = xmlAptitudeNodeList.Item(i)
      Set xmlQuestionNode = xmlCurAptitudeNode.getElementsByTagName("Question").Item(0)
      
      If xmlQuestionNode.Text = tvTreeView.SelectedItem.Text Then
         sQuestion = tvTreeView.SelectedItem.Text
         Set xmlOptionsNodeList = xmlCurAptitudeNode.getElementsByTagName("Option")
         For j = 0 To xmlOptionsNodeList.length - 1
            Set xmlOptionsNode = xmlOptionsNodeList.Item(j)
            sOption = sOption & " (" & (j + 1) & "). " & xmlOptionsNode.Text & "   "
         Next
         
         Set xmlAnswerNode = xmlCurAptitudeNode.getElementsByTagName("Answer").Item(0)
         sAnswer = xmlAnswerNode.Text
         ShowPane sQuestion, sOption, sAnswer
         Exit For
      End If
   Next
End Sub

Private Sub ShowPane(sQuestion As String, sOptions As String, sAnswer As String)
   Dim sQ As String
   sQ = Replace(sQuestion, Chr(10), vbCrLf)
   txtQuestion.Text = sQ
   txtOptions.Text = sOptions
   lblAnswerIndex = sAnswer
End Sub

Private Sub ClearPane()
   txtQuestion.Text = ""
   txtOptions.Text = ""
   lblAnswerIndex = ""
End Sub
