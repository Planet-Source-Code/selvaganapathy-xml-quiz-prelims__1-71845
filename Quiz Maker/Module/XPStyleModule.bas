Attribute VB_Name = "mXPStyleModule"
Option Explicit

Private Type tagInitCommonControlsEx
   lngSize As Long
   lngICC As Long
End Type
Private Const ICC_USEREX_CLASSES = &H200

Private Declare Function InitCommonControlsEx Lib "comctl32.dll" _
   (iccex As tagInitCommonControlsEx) As Boolean
Declare Function PathFileExists Lib "shlwapi.dll" Alias "PathFileExistsA" (ByVal pszPath As String) As Long

Public Function IsFileExists(sPath As String) As Boolean
   IsFileExists = (PathFileExists(sPath) <> 0)
End Function


Public Sub CreateManifestFile()
   Dim sFileName As String
   Dim iFNum As Integer
   Dim sStr As String
   
   sFileName = App.Path & "\" & App.EXEName & ".exe.manifest"
   
   If IsFileExists(sFileName) = False Then
   
      sStr = "<?xml version=" & Chr(34) & "1.0" & Chr(34) & " encoding=" & Chr(34) & "UTF-8" & Chr(34) & " standalone=" & Chr(34) & "yes" & Chr(34) & "?>" & vbCrLf & _
            "<assembly xmlns=" & Chr(34) & "urn:schemas-microsoft-com:asm.v1" & Chr(34) & " manifestVersion=" & Chr(34) & "1.0" & Chr(34) & ">" & vbCrLf & _
            "<assemblyIdentity" & vbCrLf & _
            "    version=" & Chr(34) & "1.0.0.0" & Chr(34) & "" & vbCrLf & _
            "    processorArchitecture=" & Chr(34) & "X86" & Chr(34) & "" & vbCrLf & _
            "    name=" & Chr(34) & App.CompanyName & "." & App.ProductName & "." & App.EXEName _
            & Chr(34) & "" & vbCrLf & _
            "    type=" & Chr(34) & "win32" & Chr(34) & "" & vbCrLf & _
            "/>" & vbCrLf & _
            "<description>" & App.FileDescription & " </description>" & vbCrLf & _
            "<dependency>" & vbCrLf & _
            "    <dependentAssembly>" & vbCrLf & _
            "        <assemblyIdentity" & vbCrLf & _
            "            type=" & Chr(34) & "win32" & Chr(34) & "" & vbCrLf & _
            "            name=" & Chr(34) & "Microsoft.Windows.Common-Controls" & Chr(34) & "" & vbCrLf & _
            "            version=" & Chr(34) & "6.0.0.0" & Chr(34) & "" & vbCrLf & _
            "            processorArchitecture=" & Chr(34) & "X86" & Chr(34) & "" & vbCrLf & _
            "            publicKeyToken=" & Chr(34) & "6595b64144ccf1df" & Chr(34) & "" & vbCrLf & _
            "            language=" & Chr(34) & "*" & Chr(34) & "" & vbCrLf & _
            "        />" & vbCrLf & _
            "    </dependentAssembly>" & vbCrLf & _
            "</dependency>" & vbCrLf & _
            "</assembly>"

      iFNum = FreeFile
      Open sFileName For Binary As #iFNum
         Put #iFNum, , sStr
      Close #iFNum
   End If
End Sub

Public Function InitCommonControlsVB() As Boolean
   On Error Resume Next
   Dim iccex As tagInitCommonControlsEx
   ' Ensure CC available:
   With iccex
      .lngSize = LenB(iccex)
      .lngICC = ICC_USEREX_CLASSES
   End With
   InitCommonControlsEx iccex
   InitCommonControlsVB = (Err.Number = 0)
   On Error GoTo 0
End Function

Public Sub AddXPStyle()
   CreateManifestFile
   InitCommonControlsVB
End Sub


