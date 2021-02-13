Attribute VB_Name = "modCommonDialog"
Option Explicit

Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal hMem As Long)
Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
Private Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long

Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long

Private Type BrowseInfo
    hWndOwner As Long
    pIDLRoot As Long
    pszDisplayName As Long
    lpszTitle As Long
    ulFlags As Long
    lpfnCallback As Long
    lParam As Long
    iImage As Long
End Type

Private Type OPENFILENAME
    lStructSize As Long
    hWndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter  As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    Flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

Private Const BIF_RETURNONLYFSDIRS = 1
Private Const MAX_PATH = 260

Private Const OFN_ALLOWMULTISELECT = &H200
Private Const OFN_CREATEPROMPT = &H2000
Private Const OFN_ENABLEHOOK = &H20
Private Const OFN_ENABLETEMPLATE = &H40
Private Const OFN_ENABLETEMPLATEHANDLE = &H80
Private Const OFN_EXPLORER = &H80000 'new look commdlg
Private Const OFN_EXTENSIONDIFFERENT = &H400
Private Const OFN_FILEMUSTEXIST = &H1000
Private Const OFN_HIDEREADONLY = &H4
Private Const OFN_LONGNAMES = &H200000 'force long names for 3.x modules
Private Const OFN_NOCHANGEDIR = &H8
Private Const OFN_NODEREFERENCELINKS = &H100000
Private Const OFN_NOLONGNAMES = &H40000 'force no long names for 4.x modules
Private Const OFN_NONETWORKBUTTON = &H20000
Private Const OFN_NOREADONLYRETURN = &H8000
Private Const OFN_NOTESTFILECREATE = &H10000
Private Const OFN_NOVALIDATE = &H100
Private Const OFN_OVERWRITEPROMPT = &H2
Private Const OFN_PATHMUSTEXIST = &H800
Private Const OFN_READONLY = &H1
Private Const OFN_SHAREAWARE = &H4000
Private Const OFN_SHAREFALLTHROUGH = 2
Private Const OFN_SHARENOWARN = 1
Private Const OFN_SHAREWARN = 0
Private Const OFN_SHOWHELP = &H10

Public Enum DialogFlags
    ALLOWMULTISELECT = OFN_ALLOWMULTISELECT
    CREATEPROMPT = OFN_CREATEPROMPT
    ENABLEHOOK = OFN_ENABLEHOOK
    ENABLETEMPLATE = OFN_ENABLETEMPLATE
    ENABLETEMPLATEHANDLE = OFN_ENABLETEMPLATEHANDLE
    EXPLORER = OFN_EXPLORER
    EXTENSIONDIFFERENT = OFN_EXTENSIONDIFFERENT
    FILEMUSTEXIST = OFN_FILEMUSTEXIST
    HIDEREADONLY = OFN_HIDEREADONLY
    LONGNAMES = OFN_LONGNAMES
    NOCHANGEDIR = OFN_NOCHANGEDIR
    NODEREFERENCELINKS = OFN_NODEREFERENCELINKS
    NOLONGNAMES = OFN_NOLONGNAMES
    NONETWORKBUTTON = OFN_NONETWORKBUTTON
    NOREADONLYRETURN = OFN_NOREADONLYRETURN
    NOTESTFILECREATE = OFN_NOTESTFILECREATE
    NOVALIDATE = OFN_NOVALIDATE
    OVERWRITEPROMPT = OFN_OVERWRITEPROMPT
    PATHMUSTEXIST = OFN_PATHMUSTEXIST
    ReadOnly = OFN_READONLY
    SHAREAWARE = OFN_SHAREAWARE
    SHAREFALLTHROUGH = OFN_SHAREFALLTHROUGH
    SHARENOWARN = OFN_SHARENOWARN
    SHAREWARN = OFN_SHAREWARN
    ShowHelp = OFN_SHOWHELP
End Enum

Public Function BrowseForFolder() As String

    Dim iNull       As Integer
    Dim lpIDList    As Long
    Dim udtBI       As BrowseInfo
    Dim sPath       As String
    
    With udtBI
        .hWndOwner = frmMain.hwnd
        .lpszTitle = lstrcat("Browse for a folder.", "")
        .ulFlags = BIF_RETURNONLYFSDIRS
    End With
    
    lpIDList = SHBrowseForFolder(udtBI)
    If lpIDList Then
        sPath = String$(MAX_PATH, 0)
        SHGetPathFromIDList lpIDList, sPath
        CoTaskMemFree lpIDList
        iNull = InStr(sPath, vbNullChar)
        If iNull Then
            sPath = Left$(sPath, iNull - 1)
        End If
    End If

    BrowseForFolder = sPath

End Function

Public Function ShowOpen(filter As String, Optional Flags As DialogFlags) As String

On Error GoTo ErrorHandler

    Dim cdlg As OPENFILENAME
    
    cdlg.hWndOwner = frmMain.hwnd
    cdlg.hInstance = App.hInstance
    cdlg.lpstrFilter = Replace$(filter, "|", Chr$(0))
    cdlg.lpstrFile = Space$(254)
    cdlg.nMaxFile = 255
    cdlg.lpstrFileTitle = Space$(254)
    cdlg.nMaxFileTitle = 255
    cdlg.lpstrTitle = "Open"
    cdlg.Flags = Flags
    cdlg.lStructSize = Len(cdlg)
    If GetOpenFileName(cdlg) Then
        ShowOpen = cdlg.lpstrFile
        ShowOpen = Replace$(ShowOpen, Chr(0), "")
        ShowOpen = Trim$(ShowOpen)
    Else
        ShowOpen = ""
    End If
    
    Exit Function

ErrorHandler:
    MsgBox "ShowOpen - modCommonDialog" & vbCrLf & vbCrLf & Err & " - " & Err.Description, vbCritical, "Error"

End Function

Public Function ShowSave(filter As String, suffix As String, Optional Flags As DialogFlags) As String

On Error GoTo ErrorHandler

    Dim cdlg As OPENFILENAME
    
    cdlg.hWndOwner = frmMain.hwnd
    cdlg.hInstance = App.hInstance
    cdlg.lpstrFilter = Replace$(filter, "|", Chr$(0))
    cdlg.lpstrFile = Replace$(frmMain.txtCDlabel.text, ".", "") & suffix & Space$(254 - Len(frmMain.txtCDlabel.text) - Len(suffix))
    cdlg.lpstrDefExt = GetExtension$(suffix)
    cdlg.nMaxFile = 255
    cdlg.lpstrFileTitle = Space$(254)
    cdlg.nMaxFileTitle = 255
    cdlg.lpstrTitle = "Save"
    cdlg.Flags = Flags
    cdlg.lStructSize = Len(cdlg)
    If GetSaveFileName(cdlg) Then
        ShowSave = cdlg.lpstrFile
        ShowSave = Replace$(ShowSave, Chr(0), "")
        ShowSave = Trim$(ShowSave)
    Else
        ShowSave = ""
    End If

    Exit Function

ErrorHandler:
    MsgBox "ShowSave - modCommonDialog" & vbCrLf & vbCrLf & Err & " - " & Err.Description, vbCritical, "Error"

End Function
