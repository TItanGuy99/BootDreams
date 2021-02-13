Attribute VB_Name = "modPrevInst"
Option Explicit

Private Declare Function AttachThreadInput Lib "user32" (ByVal idAttach As Long, ByVal idAttachTo As Long, ByVal fAttach As Long) As Long
Private Declare Function GetForegroundWindow Lib "user32" () As Long
Private Declare Function SetForegroundWindow Lib "user32.dll" (ByVal hwnd As Long) As Long
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Private Declare Function EnumWindows Lib "user32.dll" (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Private Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByRef lParam As Any) As Long
Private Declare Function GetProp Lib "user32.dll" Alias "GetPropA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Private Declare Function SetProp Lib "user32.dll" Alias "SetPropA" (ByVal hwnd As Long, ByVal lpString As String, ByVal hData As Long) As Long

Private prevInstPropName As String  ' unique property value passed to IsPrevInstance
Private prevInstHwnd As Long        ' this will be the previous instance hWnd, if any
Private prevInstPropValue As Long   ' this will be the property value stored against the previous hWnd

Public Function IsPrevInstance(ByVal hwnd As Long, ByVal PropName As String, _
                        Optional ByRef propValue As Long = 1, _
                        Optional ByVal passCmdLineToPropValue As Boolean) As Long
    
    ' hWnd [in]: must be the hWnd of the instance being created (Me.hWnd)
    ' PropName [in]: unique property name, no other winodw should be using
    ' propValue [in/out]: the property value to be associated with this instance
    '     to pass command line parameters from other instances, suggest
    '     passing a textbox hWnd. This value is for your use, it must not be 0
    ' passCmdLineToPropValue [in]: since it is a common practice to pass the command line
    '     to the 1st instance when this is the 2nd instance, setting that parameter
    '     to True will allow this routine to pass the command line for you.
    
    ' Note that propValue is ByRef; therefore it can be changed by this function.
    ' Do not pass Read-Only values (i.e., do not pass Text1.hWnd as propValue)
    
    ' Return Values:
    ' If this is the first instance
    ' - function returns zero
    ' - The propValue is assigned to the PropName property
    ' - passCmdLineToPropValue is ignored
    
    ' If this is another instance...
    ' - function returns hWnd of the previous instance
    ' - propValue is the custom value set by the previous instance
    ' - if passCmdLineToPropValue, then function sends the Command$ variable
    '   to the value in propValue. Of course, this should be an hWnd of
    '   a textbox or other window that has a Text property.
    
    
    ' sanity checks
    If Trim$(PropName) = "" Then Exit Function
    If hwnd = 0 Then Exit Function
    ' should you want to add even more sanity checks, consider using
    ' the IsWindow API on passed hWnd, and also on the propValue just
    ' before using SendMessage below
    
    Const WM_SETTEXT As Long = &HC
    prevInstHwnd = 0    ' reset
    prevInstPropName = PropName ' set here so EnumWindowsProc can use it
    
    ' look for previous instance
    EnumWindows AddressOf EnumWindowsProc, hwnd
    prevInstPropName = vbNullString ' no longer need to waste extra memory
    
    If prevInstHwnd = 0 Then    ' no previous instance found
        ' ensure the property value passed is not zero
        If propValue = 0 Then propValue = 1
        ' now assign the property & value
        SetProp hwnd, PropName, propValue
    Else
        ' we do have a previous instance, set return values
        propValue = prevInstPropValue
        If passCmdLineToPropValue = True Then
            ' ensure a non-zero length string so the target's Change_Event occurs
            ' It is assumed the target will reset its text value to zero-length
            ' when done receiving this message
            SendMessage propValue, WM_SETTEXT, 0&, ByVal Command$ & " "
        End If
        IsPrevInstance = prevInstHwnd
    End If
End Function

Private Function EnumWindowsProc(ByVal hwnd As Long, ByVal lParam As Long) As Long
    
    ' the enumeration stops when the return value is zero or all windows have been enumerated
    prevInstPropValue = GetProp(hwnd, prevInstPropName)
    If prevInstPropValue = 0 Then
        EnumWindowsProc = 1 ' keep enumerating
    Else
        If hwnd = lParam Then
            ' safety check. Should this be called a 2nd time from the previous instance
            ' ensure we don't return true if this is the previous instance
            EnumWindowsProc = 1
        Else
            prevInstHwnd = hwnd
            ' stops enumerating cause we do not set its return value
        End If
    End If
End Function

Public Sub pSetForegroundWindow(ByVal hwnd As Long)

    Dim lForeThreadID As Long
    Dim lThisThreadID As Long
    Dim lReturn       As Long
    '
    ' Make a window, specified by its handle (hwnd)
    ' the foreground window.
    '
    ' If it is already the foreground window, exit.
    '
    If hwnd <> GetForegroundWindow() Then
        '
        ' Get the threads for this window and the foreground window.
        '
        lForeThreadID = GetWindowThreadProcessId(GetForegroundWindow, ByVal 0&)
        lThisThreadID = GetWindowThreadProcessId(hwnd, ByVal 0&)
        '
        ' By sharing input state, threads share their concept of
        ' the active window.
        '
        If lForeThreadID <> lThisThreadID Then
            ' Attach the foreground thread to this window.
            Call AttachThreadInput(lForeThreadID, lThisThreadID, True)
            ' Make this window the foreground window.
            lReturn = SetForegroundWindow(hwnd)
            ' Detach the foreground window's thread from this window.
            Call AttachThreadInput(lForeThreadID, lThisThreadID, False)
        Else
           lReturn = SetForegroundWindow(hwnd)
        End If
        
        'Restore this window to its normal size.
        If frmMain.WindowState = vbMinimized Then frmMain.WindowState = vbNormal
    End If

End Sub
