VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTimer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim mMin As Integer
Dim mSec As Integer
Dim mHr As Integer

Public Function IsEqual(mSnd As Integer, mMt As Integer, mHur As Integer) As Boolean
    If mSnd = mSec And mMt = mMin And mHur = mHr Then
        IsEqual = True
    End If
End Function

Public Property Get mMinute() As Integer
    mMinute = mMin
End Property

Public Property Let mMinute(ByVal vNewValue As Integer)
    If vNewValue >= 60 Then
        mHr = mHr + 1
        If mHr >= 12 Then
            mHr = 1
        End If
    End If
        mMin = vNewValue Mod 60
End Property

Public Property Get mSecond() As Integer
    mSecond = mSec
End Property

Public Property Let mSecond(ByVal vNewValue As Integer)
    If vNewValue >= 60 Then
        mMin = mMin + 1
        If mMin >= 60 Then
            mHr = mHr + 1
                If mHr >= 12 Then
                    mHr = 1
                End If
        End If
    End If
        mSec = vNewValue Mod 60
End Property

Public Property Get mHour() As Integer
    mHour = mHr
End Property

Public Property Let mHour(ByVal vNewValue As Integer)
    If vNewValue >= 12 Then
        mHr = 1
    Else
        mHr = vNewValue
    End If
End Property


Public Property Get ToString() As String
    ToString = Format(mHr, "00") & " : " & Format(mMin, "00") & " : " & Format(mSec, "00")
End Property
