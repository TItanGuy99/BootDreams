Attribute VB_Name = "modDummyFile"
Option Explicit

Public Enum DummyFormat
    AudioDummy
    DataDummy
End Enum

'this sub creates a dummy file (filled with zeros)
'dummy files are useful for load intensive programs like Quake, they _
help on load times
Public Sub MakeDummyFile(DFormat As DummyFormat, msinfo As Long)

On Error GoTo ErrorHandler
    
    Dim fn As String
    Dim fs As Long
    Dim ds As Long
    Dim cs As Long
    Dim ucds As Long
    Dim ss As Long
    
    If DFormat = AudioDummy Then
        fn = AppPath$ & "audio.raw"
    Else
        fn = AddTrailingSlash$(frmMain.txtFoldername.text) & "000DUMMY.DAT"
    End If
    
    'delete the previous dummy
    If FileExists(fn) = True Then
        Call Kill(fn)
    End If
    
    'used CD space
    ucds = msinfo
    
    ss = CLng((((lngDummySize - 2) * CLng(1024)) * CLng(1024)) / CLng(2048))
    ds = ss - ucds
    
    'folder size in bytes
    fs = GetPathSize(frmMain.txtFoldername.text) \ 2048
    
    'size in bytes to create dummy
    ds = ds - fs
    
    If DFormat = AudioDummy Then
        ds = ds * 2352
    Else
        ds = ds * 2048
    End If
    
    'create the dummy
    Call ShellWait("""" & AppPath$ & "tools\newfile.exe"" " & ds & " """ & fn & """", vbNormalFocus)
  
    Exit Sub

ErrorHandler:
    MsgBox "MakeDummyFile - modDummyFile" & vbCrLf & vbCrLf & Err & " - " & Err.Description, vbCritical, "Error"

End Sub
