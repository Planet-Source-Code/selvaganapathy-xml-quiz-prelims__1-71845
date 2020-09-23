Attribute VB_Name = "mUtility"
Option Explicit

Public Sub SelectText(oText As TextBox)
   oText.SelStart = 0
   oText.SelLength = Len(oText.Text)
End Sub

Public Sub ChangeFocus(KeyAscii As Integer)
   If KeyAscii = vbKeyReturn Then KeyAscii = 0: SendKeys "{TAB}"
End Sub

