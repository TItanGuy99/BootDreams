Attribute VB_Name = "modPubFunctions"
Option Explicit

Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

'INI
Public Declare Function WritePrivateProfileString& Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal AppName$, ByVal KeyName$, ByVal keydefault$, ByVal Filename$)
Public Declare Function GetPrivateProfileString& Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal AppName$, ByVal KeyName$, ByVal keydefault$, ByVal ReturnedString$, ByVal RSSize&, ByVal Filename$)

Public Declare Function PathIsDirectory Lib "shlwapi.dll" Alias "PathIsDirectoryA" (ByVal pszPath As String) As Long
Private Declare Function PathFileExists Lib "shlwapi.dll" Alias "PathFileExistsA" (ByVal pszPath As String) As Long

'ShellWait
Private Declare Function CreatePipe Lib "kernel32" (phReadPipe As Long, phWritePipe As Long, lpPipeAttributes As Any, ByVal nSize As Long) As Long
Private Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, ByVal lpBuffer As String, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, ByVal lpOverlapped As Any) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function CreateProcessA Lib "kernel32" (ByVal lpApplicationName As Long, ByVal lpCommandLine As String, ByVal lpProcessAttributes As Any, ByVal lpThreadAttributes As Any, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, ByVal lpEnvironment As Long, ByVal lpCurrentDirectory As Long, lpStartupInfo As Any, lpProcessInformation As Any) As Long

Public Const SW_HIDE = 0
Public Const SW_SHOWNORMAL = 1

Private Const INFINITE As Long = -1
Private Const NORMAL_PRIORITY_CLASS = &H20
Private Const STARTF_USESTDHANDLES = &H100
Private Const STARTF_USESHOWWINDOW = &H1

Private Type STARTUPINFO
    cb As Long
    lpReserved As Long
    lpDesktop As Long
    lpTitle As Long
    dwX As Long
    dwY As Long
    dwXSize As Long
    dwYSize As Long
    dwXCountChars As Long
    dwYCountChars As Long
    dwFillAttribute As Long
    dwFlags As Long
    wShowWindow As Integer
    cbReserved2 As Integer
    lpReserved2 As Long
    hStdInput As Long
    hStdOutput As Long
    hStdError As Long
End Type

Private Type PROCESS_INFORMATION
    hProcess As Long
    hThread As Long
    dwProcessId As Long
    dwThreadID As Long
End Type

Public Function AppPath() As String

    AppPath$ = App.Path
    If Right$(AppPath$, 1) <> "\" Then AppPath$ = AppPath$ & "\"

End Function

Public Function AppEXE() As String

    AppEXE$ = AppPath$ & App.EXEName & ".exe"

End Function

Public Function FileExists(ByVal sFile As String) As Boolean

    FileExists = CBool(PathFileExists(sFile))

End Function

Public Function StripLineEndings(ByVal text As String) As String
    
    Dim temp As String

    temp = Replace$(text, vbCrLf, "")
    text = Replace$(temp, vbLf, "")
    StripLineEndings$ = text

End Function

Public Function AddTrailingSlash(ByVal sPath As String) As String
    
    If Right$(sPath, 1) <> "\" Then
        AddTrailingSlash$ = sPath & "\"
    End If

End Function

Public Function JustTitle(ByVal sFile As String) As String

    Dim temp As Integer
    
    temp = InStrRev(sFile, "\")
    
    If temp Then
        JustTitle$ = Mid$(sFile, temp + 1)
    Else
        JustTitle$ = sFile
    End If

End Function

Public Function GetExtension(ByVal sFile As String) As String

    Dim title As String
    Dim temp As Integer
    
    title = JustTitle$(sFile)
    temp = InStrRev(title, ".")
    
    If temp Then
        GetExtension$ = LCase$(Mid$(title, temp + 1))
    Else
        GetExtension$ = ""
    End If

End Function

Public Function GetPathSize(ByVal sPathName As String) As Long

    Dim sFileName       As String
    Dim dSize           As Double
    Dim asFileName()    As String
    Dim i               As Long

    sPathName = AddTrailingSlash$(sPathName)
    sFileName = Dir$(sPathName, vbDirectory + vbHidden + vbSystem + vbReadOnly)
    
    Do While Len(sFileName) > 0
        If sFileName <> "." And sFileName <> ".." Then
            ReDim Preserve asFileName(i)
            asFileName(i) = sPathName & sFileName
            i = i + 1
        End If
        sFileName = Dir
    Loop
    
    If i > 0 Then
        For i = 0 To UBound(asFileName)
            If (GetAttr(asFileName(i)) And vbDirectory) = vbDirectory Then
                dSize = dSize + GetPathSize(asFileName(i))
            Else
                dSize = dSize + FileLen(asFileName(i))
            End If
        Next
    End If
    
    GetPathSize = dSize

End Function

Public Function ReadINI(section As String, key As String) As String

    Dim sbuff         As String
    Dim lbuffsize     As Long
    
    sbuff = Space(255)
    lbuffsize = Len(sbuff)
    lbuffsize = GetPrivateProfileString(section, key, "", sbuff, lbuffsize, AppPath$ & "tools\settings.ini")
    
    If lbuffsize > 0 Then
        ReadINI = Left(sbuff, lbuffsize)
    Else
        ReadINI = ""
    End If
    
End Function

Public Sub WriteINI(section As String, key As String, value As String)

    Call WritePrivateProfileString(section, key, value, AppPath$ & "tools\settings.ini")
    
End Sub

Public Sub ShellWait(Pathname As String, Optional ByVal WindowStyle As Long)
    
    Dim proc    As PROCESS_INFORMATION
    Dim start   As STARTUPINFO
    Dim Ret     As Long
    
    start.cb = Len(start)
    If Not IsMissing(WindowStyle) Then
        start.dwFlags = STARTF_USESHOWWINDOW
        start.wShowWindow = WindowStyle
    End If
    
    Ret = CreateProcessA(0&, Pathname, 0&, 0&, 1&, NORMAL_PRIORITY_CLASS, 0&, 0&, start, proc)
    Ret = WaitForSingleObject(proc.hProcess, INFINITE)
    Ret = CloseHandle(proc.hProcess)

End Sub

Public Function ExecuteApp(ByVal sCmdline As String) As String

    Dim proc            As PROCESS_INFORMATION
    Dim start           As STARTUPINFO
    Dim Ret             As Long
    Dim hReadPipe       As Long
    Dim hWritePipe      As Long
    Dim sOutput         As String
    Dim lngBytesRead    As Long
    Dim sBuffer         As String
    
    sBuffer = Space(256)
    Ret = CreatePipe(hReadPipe, hWritePipe, 0&, 0)
    start.cb = Len(start)
    start.dwFlags = STARTF_USESTDHANDLES Or STARTF_USESHOWWINDOW
    
    start.hStdOutput = hWritePipe
    start.hStdError = hWritePipe
    start.wShowWindow = SW_HIDE
    
    Ret = CreateProcessA(0&, sCmdline, 0&, 0&, 1&, NORMAL_PRIORITY_CLASS, 0&, 0&, start, proc)
    CloseHandle hWritePipe
    
    Do
        Ret = ReadFile(hReadPipe, sBuffer, Len(sBuffer), lngBytesRead, 0&)
        sOutput = sOutput & Left(sBuffer, lngBytesRead)
    Loop While Ret <> 0
    
    CloseHandle proc.hProcess
    CloseHandle proc.hThread
    CloseHandle hReadPipe
    
    ExecuteApp = sOutput

End Function
