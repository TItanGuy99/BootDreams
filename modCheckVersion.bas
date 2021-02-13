Attribute VB_Name = "modCheckVersion"
Option Explicit

Private Declare Function InternetOpen Lib "wininet" Alias "InternetOpenA" (ByVal sAgent As String, ByVal lAccessType As Long, ByVal sProxyName As String, ByVal sProxyBypass As String, ByVal lFlags As Long) As Long
Private Declare Function InternetCloseHandle Lib "wininet" (ByVal hInet As Long) As Integer
Private Declare Function InternetReadFile Lib "wininet" (ByVal hFile As Long, ByVal sBuffer As String, ByVal lNumBytesToRead As Long, lNumberOfBytesRead As Long) As Integer
Private Declare Function InternetOpenUrl Lib "wininet" Alias "InternetOpenUrlA" (ByVal hInternetSession As Long, ByVal lpszUrl As String, ByVal lpszHeaders As String, ByVal dwHeadersLength As Long, ByVal dwFlags As Long, ByVal dwContext As Long) As Long

Private Const scUserAgent = "vbUpdate"
Private Const INTERNET_OPEN_TYPE_DIRECT = 1
Private Const INTERNET_FLAG_RELOAD = &H80000000

Public Sub CheckVersion()

    Dim Update() As String
    
    Update = Split(GetURL$(strOnlineVersion), vbCrLf)
    
    If Update(0) <> App.title Then
        MsgBox "Unable to check for a update.", vbCritical, "Error"
        Exit Sub
    End If
    
    If Update(1) = strCurVersion Then
        MsgBox "You have the latest version.", vbInformation, "Information"
    Else
        If MsgBox("A new version is available." & vbNewLine & vbNewLine & "Your version: " & strCurVersion & vbNewLine & "Newest version: " & Update(1) & vbNewLine & vbNewLine & "Goto the website?", vbQuestion + vbYesNo, "Question") = vbYes Then
            Call ShellExecute(0, vbNullString, Update(2), vbNullString, vbNullString, vbNormalFocus)
        End If
    End If

End Sub

Public Function GetURL(sURL As String) As String
    
    Dim hOpen       As Long
    Dim hFile       As Long
    Dim sBuffer     As String
    Dim Ret         As Long
    Dim bSuccess    As Boolean
    
    sBuffer = Space$(1000)
    
    hOpen = InternetOpen(scUserAgent, INTERNET_OPEN_TYPE_DIRECT, vbNullString, vbNullString, 0)
    hFile = InternetOpenUrl(hOpen, sURL, vbNullString, ByVal 0&, INTERNET_FLAG_RELOAD, ByVal 0&)
    
    Do
        bSuccess = InternetReadFile(hFile, sBuffer, 1000, Ret)
        If Ret = 0 Then Exit Do
        If bSuccess Then
            GetURL$ = GetURL$ & sBuffer
        End If
        DoEvents
    Loop
    
    GetURL = Trim$(GetURL$)
    
    InternetCloseHandle hFile
    InternetCloseHandle hOpen

End Function
