Attribute VB_Name = "modAssociate"
Option Explicit

Public Sub ReadAssociation()

On Error Resume Next

    Dim b As Object
    Dim reg1 As String
    Dim reg2 As String
    Dim reg3 As String
    
    Set b = CreateObject("wscript.shell")
    
    reg1 = b.regread("HKLM\Software\Classes\DiscJuggler CD image\shell\open\command\")
    reg2 = b.regread("HKLM\Software\Classes\ISO image\shell\open\command\")
    reg3 = b.regread("HKLM\Software\Classes\CUE sheet\shell\open\command\")
    
    With frmAssociation
        If reg1 = AppEXE & " %L" Then .chkDiscJuggler.value = 1 Else .chkDiscJuggler.value = 0
        If reg2 = AppEXE & " %L" Then .chkISO.value = 1 Else .chkISO.value = 0
        If reg3 = AppEXE & " %L" Then .chkCUE.value = 1 Else .chkCUE.value = 0
    End With

End Sub

Public Sub AssociateExtension(sExt As String, sDesc As String)

On Error Resume Next

    Dim b As Object
    
    Set b = CreateObject("wscript.shell")
    
    b.regwrite "HKLM\Software\Classes\" & sExt & "\", sDesc
    b.regwrite "HKLM\Software\Classes\" & sDesc & "\", sDesc
    b.regwrite "HKLM\Software\Classes\" & sDesc & "\DefaultIcon\", AppEXE & ",1"
    b.regwrite "HKLM\Software\Classes\" & sDesc & "\shell\open\command\", AppEXE & " %L"

End Sub

Public Sub DisAssociateExtension(sExt As String, sDesc As String)

On Error Resume Next

    Dim b As Object
    
    Set b = CreateObject("wscript.shell")
    
    b.regdelete "HKLM\Software\Classes\" & sExt & "\"
    b.regdelete "HKLM\Software\Classes\" & sDesc & "\DefaultIcon\"
    b.regdelete "HKLM\Software\Classes\" & sDesc & "\shell\open\command\"
    b.regdelete "HKLM\Software\Classes\" & sDesc & "\shell\open\"
    b.regdelete "HKLM\Software\Classes\" & sDesc & "\shell\"
    b.regdelete "HKLM\Software\Classes\" & sDesc & "\"
    
End Sub
