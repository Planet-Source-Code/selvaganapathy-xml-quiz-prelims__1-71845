VERSION 5.00
Begin VB.Form frmAddQuestion 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add Question"
   ClientHeight    =   6735
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7575
   Icon            =   "frmAddQuestion.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6735
   ScaleWidth      =   7575
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtAnswerIndex 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   2700
      TabIndex        =   8
      Top             =   6360
      Width           =   1815
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   6660
      TabIndex        =   5
      Top             =   3480
      Width           =   855
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   6660
      TabIndex        =   4
      Top             =   3000
      Width           =   855
   End
   Begin VB.TextBox txtWord 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   945
      Left            =   60
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   3000
      Width           =   6555
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6060
      TabIndex        =   10
      Top             =   6360
      Width           =   1455
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&OK"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4560
      TabIndex        =   9
      Top             =   6360
      Width           =   1455
   End
   Begin VB.ListBox lstOptions 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2370
      ItemData        =   "frmAddQuestion.frx":0442
      Left            =   60
      List            =   "frmAddQuestion.frx":0444
      TabIndex        =   6
      Top             =   3960
      Width           =   7455
   End
   Begin VB.TextBox txtQuestion 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   60
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   360
      Width           =   7455
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E9AA4B&
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
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   7215
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E9AA4B&
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
      Left            =   60
      TabIndex        =   2
      Top             =   2700
      Width           =   6285
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E9AA4B&
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
      Height          =   345
      Left            =   60
      TabIndex        =   7
      Top             =   6360
      Width           =   2580
   End
End
Attribute VB_Name = "frmAddQuestion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public xmlAptitudeElement  As IXMLDOMElement
Public xmlDOMDoc           As IXMLDOMDocument
Public iResult             As VbMsgBoxResult
Public xmlAptShowElement   As IXMLDOMElement

Public Sub ElementToControl()
   Dim qNode         As IXMLDOMElement
   Dim optNodeList   As IXMLDOMNodeList
   Dim AnsNode       As IXMLDOMElement
   
   Dim i As Integer
   Set qNode = xmlAptShowElement.getElementsByTagName("Question").Item(0)
   txtQuestion.Text = qNode.Text
   Set optNodeList = xmlAptShowElement.getElementsByTagName("Option")
   lstOptions.Clear
   For i = 0 To optNodeList.length - 1
      lstOptions.AddItem optNodeList.Item(i).Text
   Next
   
   If xmlAptShowElement.getElementsByTagName("Answer").length > 0 Then
      Set AnsNode = xmlAptShowElement.getElementsByTagName("Answer").Item(0)
      txtAnswerIndex.Text = AnsNode.Text
   End If
End Sub

Private Sub cmdCancel_Click()
   iResult = vbCancel
   Me.Hide
End Sub

Private Sub cmdOK_Click()
   Dim xmlElement As IXMLDOMElement
   Dim xmlText As IXMLDOMText
   Dim i As Integer
   
   If lstOptions.ListCount > 0 And Trim(txtAnswerIndex) = "" Then
      MsgBox "Please Choose any answer.", vbCritical
      Exit Sub
   End If
   
   If txtQuestion.Text = "" Then
      MsgBox "Please enter the Question.", vbInformation, "No Question"
      txtQuestion.SetFocus
      Exit Sub
   End If
      
   Set xmlElement = xmlDOMDoc.createElement("Question")
   Set xmlText = xmlDOMDoc.createTextNode(txtQuestion.Text)
   xmlElement.appendChild xmlText
   xmlAptitudeElement.appendChild xmlElement
   
   For i = 0 To lstOptions.ListCount - 1
      Set xmlElement = xmlDOMDoc.createElement("Option")
      Set xmlText = xmlDOMDoc.createTextNode(lstOptions.List(i))
      xmlElement.appendChild xmlText
      xmlAptitudeElement.appendChild xmlElement
   Next
   
   If lstOptions.ListCount = 0 And txtWord.Text <> "" Then
      Set xmlElement = xmlDOMDoc.createElement("Option")
      Set xmlText = xmlDOMDoc.createTextNode(txtWord)
      xmlElement.appendChild xmlText
      xmlAptitudeElement.appendChild xmlElement
      Set xmlElement = xmlDOMDoc.createElement("Answer")
      Set xmlText = xmlDOMDoc.createTextNode("0")
      xmlElement.appendChild xmlText
      xmlAptitudeElement.appendChild xmlElement
   Else
      Set xmlElement = xmlDOMDoc.createElement("Answer")
      Set xmlText = xmlDOMDoc.createTextNode(txtAnswerIndex)
      xmlElement.appendChild xmlText
      xmlAptitudeElement.appendChild xmlElement
   End If
   
   iResult = vbYes
   Me.Hide
End Sub

Private Sub Form_Load()
   iResult = vbCancel
   cmdAdd.Enabled = False
   cmdDelete.Enabled = False
   txtAnswerIndex.Locked = True
End Sub

Private Sub cmdAdd_Click()
    lstOptions.AddItem Trim(txtWord.Text)
    lstOptions.ListIndex = lstOptions.NewIndex
    txtWord.Text = ""
    txtWord.SetFocus
    cmdOk.Default = True
End Sub

Private Sub cmdDelete_Click()
    lstOptions.RemoveItem lstOptions.ListIndex
    txtWord.Text = ""
    txtWord.SetFocus
    lstOptions.ListIndex = -1
End Sub

Private Sub lstOptions_Click()
    txtWord.Text = lstOptions.Text
    txtAnswerIndex.Text = lstOptions.ListIndex
End Sub

Private Sub txtWord_Change()
Dim iInd As Integer
   iInd = CheckWordInList(txtWord.Text)
   If Len(Trim(txtWord.Text)) = 0 Then
      cmdAdd.Enabled = False
      cmdDelete.Enabled = False
   ElseIf iInd = -1 Then
      cmdDelete.Enabled = False
      cmdAdd.Enabled = True
      cmdAdd.Default = True
   Else
      lstOptions.ListIndex = iInd
      cmdDelete.Default = True
      cmdDelete.Enabled = True
      cmdAdd.Enabled = False
   End If
End Sub

Private Function CheckWordInList(sWord As String) As Integer
   Dim i As Integer
   For i = 0 To lstOptions.ListCount - 1
      If (Trim(sWord)) = (lstOptions.List(i)) Then
         CheckWordInList = i
         Exit Function
      End If
   Next i
   CheckWordInList = -1
End Function



