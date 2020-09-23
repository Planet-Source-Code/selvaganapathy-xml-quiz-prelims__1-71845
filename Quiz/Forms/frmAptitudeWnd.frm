VERSION 5.00
Begin VB.Form frmQuizWnd 
   Caption         =   "Test"
   ClientHeight    =   7170
   ClientLeft      =   180
   ClientTop       =   375
   ClientWidth     =   9810
   Icon            =   "frmAptitudeWnd.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7170
   ScaleWidth      =   9810
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrTime 
      Interval        =   1000
      Left            =   0
      Top             =   0
   End
   Begin VB.PictureBox picControlsContainer 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   510
      Left            =   0
      ScaleHeight     =   480
      ScaleWidth      =   9780
      TabIndex        =   0
      Top             =   6660
      Width           =   9810
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
         Left            =   5700
         TabIndex        =   13
         Top             =   60
         Width           =   1275
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
         Left            =   4380
         TabIndex        =   12
         Top             =   60
         Width           =   1275
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
         Left            =   7020
         TabIndex        =   11
         Top             =   60
         Width           =   1275
      End
      Begin VB.Label lblTimer 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H001DCD2F&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   2280
         TabIndex        =   15
         Top             =   120
         Width           =   1965
      End
      Begin VB.Label lblQuestNumber 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H001DCD2F&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   0
         TabIndex        =   14
         Top             =   120
         Width           =   2145
      End
   End
   Begin VB.PictureBox picQuestionContainer 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2115
      Left            =   0
      ScaleHeight     =   2115
      ScaleWidth      =   9810
      TabIndex        =   6
      Top             =   0
      Width           =   9810
      Begin VB.TextBox txtQuestion 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1815
         Left            =   480
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   7
         Top             =   420
         Width           =   6255
      End
      Begin VB.Label lblQuestion 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E9AA4B&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "&Question"
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
         Left            =   0
         TabIndex        =   8
         Top             =   0
         Width           =   6255
      End
   End
   Begin VB.PictureBox picOptionContainer 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   3375
      Left            =   0
      ScaleHeight     =   3345
      ScaleWidth      =   9780
      TabIndex        =   1
      Top             =   2115
      Width           =   9810
      Begin VB.HScrollBar hscOption 
         Height          =   255
         Left            =   1980
         TabIndex        =   10
         Top             =   3000
         Width           =   1275
      End
      Begin VB.VScrollBar vscOption 
         Height          =   1395
         Left            =   6420
         TabIndex        =   5
         Top             =   480
         Width           =   255
      End
      Begin VB.PictureBox picOutContainer 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   2595
         Left            =   120
         ScaleHeight     =   2565
         ScaleWidth      =   5205
         TabIndex        =   2
         Top             =   300
         Width           =   5235
         Begin VB.PictureBox picInContainer 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   735
            Left            =   0
            ScaleHeight     =   735
            ScaleWidth      =   1875
            TabIndex        =   3
            Top             =   0
            Width           =   1875
            Begin VB.OptionButton optOption 
               BackColor       =   &H00FFFFFF&
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
               Index           =   0
               Left            =   0
               TabIndex        =   4
               Top             =   0
               Width           =   2295
            End
         End
      End
      Begin VB.Label lblOption 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E9AA4B&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "&Options"
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
         Left            =   0
         TabIndex        =   9
         Top             =   0
         Width           =   6285
      End
   End
End
Attribute VB_Name = "frmQuizWnd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim bOptionExist  As Boolean
Dim mQuestions    As CAptQuestions
Dim mCQIndex      As Long
Dim mAnsIndex()   As Long
Dim mAttend       As Long
Dim mTotalQuestions As Long
Dim oTimer        As CTimer
   

Private Sub cmdBack_Click()
Dim mCC As Long
   If mCQIndex > 0 Then
      mCC = SelectedIndex
      If (mCC <> -1) Then
         mAnsIndex(mCQIndex) = mCC
         If mCQIndex = mAttend Then
            mAttend = mAttend + 1
         End If
      End If

      mCQIndex = mCQIndex - 1
      
      LoadQuestion mQuestions.Item(mCQIndex)
      optOption(mAnsIndex(mCQIndex)).Value = True
      lblQuestNumber = (mCQIndex + 1) & "/" & mTotalQuestions
   End If
   cmdBack.Enabled = Not (mCQIndex = 0)
   cmdNext.Enabled = Not (mCQIndex = mTotalQuestions - 1)
End Sub

Private Sub cmdFinish_Click()
   '
Dim mCC As Long
   mCC = SelectedIndex
   If (mCC <> -1) Then
      If mAnsIndex(mCQIndex) = -1 Then
         mAttend = mAttend + 1
      End If
      mAnsIndex(mCQIndex) = mCC
   End If
   
   If mAttend < mTotalQuestions Then
      If MsgBox("Yet some questions there. Do you want to finish?", vbInformation + vbYesNo, "Finish") = vbNo Then
         
         Exit Sub
      Else
         Unload Me
      End If
   Else
      Dim fResult As frmResult
      Set fResult = New frmResult
      
      fResult.SetResult mTotalQuestions, NoOfCorrect, mAttend
      fResult.Show
      Unload Me
   End If
End Sub

Private Sub cmdNext_Click()
Dim mCC As Long

   If mCQIndex < mTotalQuestions - 1 Then
      mCC = SelectedIndex
      If (mCC = -1) Then
         MsgBox "Please choose any one option", vbCritical, "Choose"
         Exit Sub
      End If
      mAnsIndex(mCQIndex) = mCC
      If mCQIndex = mAttend Then
         mAttend = mAttend + 1
      End If
      mCQIndex = mCQIndex + 1
      LoadQuestion mQuestions.Item(mCQIndex)
      If (mAnsIndex(mCQIndex) <> -1) Then
         optOption(mAnsIndex(mCQIndex)).Value = True
      End If
      lblQuestNumber = (mCQIndex + 1) & "/" & mTotalQuestions
   End If

   cmdNext.Enabled = Not (mCQIndex = mTotalQuestions - 1)
   cmdBack.Enabled = Not (mCQIndex = 0)
End Sub

Private Sub Form_Load()
   Dim i As Long
   
   Set mQuestions = DomRootToCQuestions(DOM.documentElement)
   Set oTimer = New CTimer
   
   If mQuestions.Count > 0 Then
   mTotalQuestions = mQuestions.Count
      ReDim mAnsIndex(mTotalQuestions - 1) As Long
      For i = 0 To mTotalQuestions - 1
         mAnsIndex(i) = -1
      Next
      LoadQuestion mQuestions.Item(0)
      mCQIndex = 0
      mAttend = 0
      lblQuestNumber = (mCQIndex + 1) & "/" & mTotalQuestions
      cmdBack.Enabled = Not (mCQIndex = 0)
      cmdNext.Enabled = Not (mCQIndex = mTotalQuestions - 1)
   End If
End Sub

Private Sub Form_Resize()
Dim i As Integer
   On Error Resume Next
   If Me.Width < 8000 Then Me.Width = 8000
   If Me.Height < 7000 Then Me.Height = 7000
   picQuestionContainer.Height = Me.Height / 3
   picOptionContainer.Height = Me.Height / 2
   picControlsContainer.Height = Me.Height - (picOptionContainer.Height + picOptionContainer.Top) - 500
End Sub

Private Sub SetScroll()
Dim lHeight As Long
Dim lWidth As Long
   lHeight = picInContainer.Height - picOutContainer.Height
   If lHeight > 0 Then
      vscOption.Max = lHeight
      vscOption.SmallChange = lHeight / 7
      vscOption.LargeChange = lHeight / 3
      vscOption.ZOrder 0
      vscOption.Visible = True
   Else
      vscOption.Visible = False
   End If

   lWidth = picInContainer.Width - picOutContainer.Width
   If lWidth > 0 Then
      hscOption.Max = lWidth
      hscOption.SmallChange = lWidth / 7
      hscOption.LargeChange = lWidth / 3
      hscOption.ZOrder 0
      hscOption.Visible = True
   Else
      hscOption.Visible = False
   End If

End Sub

Private Sub hscOption_Change()
   picInContainer.Left = -vscOption.Value
End Sub

Private Sub picOptionContainer_Resize()
Dim i As Integer
   lblOption.Move 0, 0, picOptionContainer.Width
   picOutContainer.Move 400, 400 + lblOption.Height, picOptionContainer.Width - 800, picOptionContainer.Height - 800 - lblOption.Height
   vscOption.Move picOutContainer.Width + picOutContainer.Left + 10, picOutContainer.Top, vscOption.Width, picOutContainer.Height
   picInContainer.Move 0, 0, picOutContainer.Width, optOption(optOption.Count - 1).Height + optOption(optOption.Count - 1).Top
   For i = 0 To optOption.Count - 1
      optOption(i).Move 0, optOption(i).Top, picInContainer.Width
   Next
   SetScroll

End Sub

Private Sub picQuestionContainer_Resize()
   On Error Resume Next
   lblQuestion.Move 0, 0, picQuestionContainer.Width
   txtQuestion.Move 0, lblQuestion.Height + lblQuestion.Top, picQuestionContainer.Width, picQuestionContainer.Height
End Sub

Private Sub tmrTime_Timer()
   oTimer.mSecond = oTimer.mSecond + 1
   lblTimer.Caption = oTimer.ToString
   If oTimer.IsEqual(0, MAX_MINUTE, 0) Then
       tmrTime.Enabled = False
       TimeOut
   End If
End Sub

Private Sub vscOption_Change()
   picInContainer.Top = -vscOption.Value
End Sub

Private Sub AddOptions(sOptions As CAptOptions)
Dim iIndex As Long
   For iIndex = 0 To sOptions.Length - 1
      AddOption sOptions.Item(iIndex)
   Next
End Sub

Private Sub AddOption(sOption As String)
Dim iIndex As Integer
Dim OFFSET As Integer
OFFSET = 300

   iIndex = optOption.Count
   If bOptionExist = True Then
      Load optOption(iIndex)
      optOption(iIndex).Move 0, optOption(iIndex - 1).Top _
      + optOption(iIndex - 1).Height, _
      optOption(iIndex - 1).Width, TextHeight(sOption) + OFFSET
   Else
      iIndex = 0
      bOptionExist = True
      optOption(iIndex).Height = TextHeight(sOption) + OFFSET      'optOption(iIndex).Height + OFFSET - 30
   End If
   optOption(iIndex).Value = False
   optOption(iIndex).Visible = True
   optOption(iIndex).Caption = sOption
   picInContainer.Move 0, 0, picOutContainer.Width, (optOption(optOption.Count - 1).Height + optOption(optOption.Count - 1).Top)
   SetScroll

End Sub

Private Sub ClearAllOption()
Dim i As Integer
   bOptionExist = False
   optOption(0).Caption = ""
   optOption(0).Value = False
   optOption(0).Visible = False
   For i = 1 To optOption.Count - 1
      Unload optOption(i)
   Next
End Sub

Private Sub LoadQuestion(oQuest As CAptQuestion)
   txtQuestion.Text = oQuest.Question
   ClearAllOption
   AddOptions oQuest.Options
End Sub

Private Function SelectedIndex() As Long
Dim i As Long
   If optOption(0).Visible = False Then
      SelectedIndex = -1
      Exit Function
   End If
   SelectedIndex = -1
   For i = 0 To optOption.Count - 1
      If optOption(i).Value = True Then
         SelectedIndex = i
         Exit Function
      End If
   Next
End Function


Private Function NoOfCorrect() As Long
   Dim i As Long
   NoOfCorrect = 0
   For i = 0 To mAttend - 1
      If mAnsIndex(i) = mQuestions.Item(i).AnswerIndex Then
         NoOfCorrect = NoOfCorrect + 1
      End If
   Next
   
End Function

Private Sub TimeOut()
   MsgBox "Your time out", vbInformation, "Time out"
   cmdFinish_Click
End Sub

