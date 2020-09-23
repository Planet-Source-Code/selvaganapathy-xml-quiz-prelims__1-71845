VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmLoadWnd 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Choose Question"
   ClientHeight    =   4770
   ClientLeft      =   270
   ClientTop       =   480
   ClientWidth     =   7860
   Icon            =   "frmLoadWnd.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4770
   ScaleWidth      =   7860
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrMotion 
      Left            =   1260
      Top             =   4260
   End
   Begin VB.CommandButton cmdFinish 
      Caption         =   "&Finish"
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
      Left            =   6240
      TabIndex        =   11
      Top             =   4260
      Width           =   1275
   End
   Begin VB.PictureBox picWizard 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   1395
      Index           =   1
      Left            =   240
      ScaleHeight     =   1365
      ScaleWidth      =   7245
      TabIndex        =   8
      Top             =   2580
      Width           =   7275
      Begin VB.TextBox txtNoOfQuestions 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   180
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   780
         Width           =   6855
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000D&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "&Number of Questions Available"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   180
         TabIndex        =   9
         Top             =   480
         Width           =   6855
      End
   End
   Begin VB.PictureBox picWizard 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   1395
      Index           =   0
      Left            =   240
      ScaleHeight     =   1365
      ScaleWidth      =   7245
      TabIndex        =   4
      Top             =   2580
      Width           =   7275
      Begin VB.CommandButton cmdLoad 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6540
         TabIndex        =   6
         Top             =   780
         Width           =   435
      End
      Begin VB.TextBox txtLoad 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   5
         Top             =   780
         Width           =   6375
      End
      Begin VB.Label lblQuestionFile 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000D&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "&Choose Question File"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   120
         TabIndex        =   7
         Top             =   480
         Width           =   6855
      End
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "&Back"
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
      Left            =   3600
      TabIndex        =   3
      Top             =   4260
      Width           =   1275
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
      Left            =   2280
      TabIndex        =   2
      Top             =   4260
      Width           =   1275
   End
   Begin VB.PictureBox picSofTech 
      Align           =   1  'Align Top
      BackColor       =   &H0057D971&
      BorderStyle     =   0  'None
      Height          =   2415
      Left            =   0
      ScaleHeight     =   2415
      ScaleWidth      =   7860
      TabIndex        =   1
      Top             =   0
      Width           =   7860
      Begin VB.Label lblSofTech 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Quiz"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   48
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   1125
         Left            =   2760
         TabIndex        =   12
         Top             =   660
         Width           =   2280
      End
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "&Next"
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
      Height          =   375
      Left            =   4920
      TabIndex        =   0
      Top             =   4260
      Width           =   1275
   End
   Begin MSComDlg.CommonDialog dlgCommonDialog 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   240
      X2              =   7480
      Y1              =   4170
      Y2              =   4170
   End
   Begin VB.Line Line1 
      X1              =   240
      X2              =   7480
      Y1              =   4140
      Y2              =   4140
   End
End
Attribute VB_Name = "frmLoadWnd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim FileName   As String
Dim iWizard    As Integer
Dim bDirty     As Boolean
Dim Motion As Integer
Dim xOffset As Integer, yOffset As Integer

Private Sub cmdBack_Click()
   If iWizard > 0 Then
      iWizard = iWizard - 1
   End If
   ShowWizard iWizard
End Sub

Private Sub cmdCancel_Click()
   Unload Me
End Sub

Private Sub cmdFinish_Click()
   If iWizard = 0 Then
      cmdNext_Click
   End If
   Set mAptTest = New frmQuizWnd
   mAptTest.Show
   Unload Me
End Sub

Private Sub cmdLoad_Click()
   With dlgCommonDialog
      .DialogTitle = "Open"
      .CancelError = False
      .Filter = "Aptitude Question File(*.qst)|*.qst|XML File (*.xml)|*.xml"
      .DefaultExt = "qst"
      .Flags = cdlOFNFileMustExist Or cdlOFNPathMustExist Or cdlOFNExtensionDifferent
      .ShowOpen
      If Len(.FileName) = 0 Then
         Exit Sub
      End If
      FileName = .FileName
   End With
   txtLoad = FileName
End Sub

Private Sub cmdNext_Click()
   Select Case iWizard
   Case 0
      If IsFileExists(txtLoad) = False Or Trim(txtLoad) = "" Then
         MsgBox "Invalid file name. Please choose the correct valid file", vbCritical, "Not found"
         txtLoad.SetFocus
         Exit Sub
      End If
      If bDirty = True Then
         LoadFile FileName
         txtNoOfQuestions = NoOfQuestions
         bDirty = False
      End If
   Case 1
      
   End Select
   If iWizard < picWizard.Count - 1 Then
      iWizard = iWizard + 1
      ShowWizard iWizard
   End If
End Sub

Private Sub ShowWizard(iIndex As Integer)
Dim i As Integer
   If iIndex < 0 Or iIndex > picWizard.Count - 1 Then
      Exit Sub
   End If
   For i = 0 To picWizard.Count - 1
      picWizard(i).Visible = False
   Next
   picWizard(iIndex).Visible = True
   If iIndex = 0 Then
      cmdBack.Enabled = False
   Else
      cmdBack.Enabled = True
   End If
   If iIndex = picWizard.Count - 1 Then
      cmdNext.Enabled = False
   Else
      cmdNext.Enabled = True
   End If
End Sub

Private Sub Form_Load()
   iWizard = 0
   ShowWizard iWizard
   Motion = 1
   xOffset = 20
   yOffset = 15
   cmdNext.Enabled = txtLoad.Text <> ""
   cmdFinish.Enabled = txtLoad.Text <> ""
End Sub

Private Sub tmrMotion_Timer()
    Select Case Motion
    Case 1
        lblSofTech.Move lblSofTech.Left - xOffset, lblSofTech.Top - yOffset
        If lblSofTech.Left <= 0 Then
            Motion = 2
        ElseIf lblSofTech.Top <= 0 Then
            Motion = 4
        End If
    Case 2
        lblSofTech.Move lblSofTech.Left + xOffset, lblSofTech.Top - yOffset
        If lblSofTech.Left >= (picSofTech.Width - lblSofTech.Width) Then
            Motion = 1
        ElseIf lblSofTech.Top <= 0 Then
            Motion = 3
        End If
    Case 3
        lblSofTech.Move lblSofTech.Left + xOffset, lblSofTech.Top + yOffset
        If lblSofTech.Left >= (picSofTech.Width - lblSofTech.Width) Then
            Motion = 4
        ElseIf lblSofTech.Top >= (picSofTech.Height - lblSofTech.Height) Then
            Motion = 2
        End If
    Case 4
        lblSofTech.Move lblSofTech.Left - xOffset, lblSofTech.Top + yOffset
        If lblSofTech.Left <= 0 Then
            Motion = 3
        ElseIf lblSofTech.Top >= (picSofTech.Height - lblSofTech.Height) Then
            Motion = 1
        End If
    End Select
End Sub

Private Sub txtLoad_Change()
   bDirty = True
   cmdNext.Enabled = txtLoad.Text <> ""
   cmdFinish.Enabled = txtLoad.Text <> ""
End Sub

Private Sub txtLoad_GotFocus()
   SelectText txtLoad
End Sub

Private Function LoadFile(sFileName As String) As Boolean
Dim Root As IXMLDOMElement

On Error GoTo LoadErr
   
   Set DOM = New DOMDocument
   DOM.Load (sFileName)
   Set Root = DOM.documentElement
   If Root.nodeName = "Test" Then
      LoadFile = True
   Else
      LoadFile = False
      MsgBox "Error : Unable to Load the file " & sFileName & vbCrLf & "This is not a valid question file.", vbCritical, " Error "
      Exit Function
   End If
   Exit Function
LoadErr:
   MsgBox "Error : Unable to Load the file " & sFileName & vbCrLf & "This is not a valid question file.", vbCritical, " Error "
   LoadFile = False
End Function

Private Function NoOfQuestions() As Long
   Dim NList As IXMLDOMNodeList
   Set NList = DOM.documentElement.getElementsByTagName("Aptitude")
   NoOfQuestions = NList.Length
End Function


