Attribute VB_Name = "modISOHeader"
Option Explicit

'this sub creates the second session ISO for data/data selfbooting CDs
'it will copy the first session ISO's header to a new file, pad it out to _
the minimum track size (300 sectors) and inject the IP.BIN in the bootsector
Public Sub CreateISOHeader(ByVal iso As String, ByVal ip As String)

On Error GoTo ErrorHandler
    
    blnISOSuccess = False
    
    'open IP.BIN
    Open ip For Binary As #1
        
        'create a buffer
        ReDim btIP(LOF(1) - 1) As Byte
        
        'put the IP.BIN into the buffer
        Get #1, , btIP()
        
        'open the ISO
        Open iso For Binary As #2
        
            'create a buffer
            ReDim btISOHead((2048& * 2&) - 1) As Byte
            
            'put the ISO header into the buffer, skipping over the boot sector
            Get #2, 32769, btISOHead()
            
        Close #2
        
        'build the second session ISO
        Open AppPath$ & "data02.iso" For Binary As #2
        
            'put the IP.BIN in the bootsector
            Put #2, , btIP()
            
            'copy the ISO's header
            Put #2, , btISOHead()
            
            'pad the rest of the ISO to the minimum track size
            Put #2, , String$((2048& * 282&), Chr$(0))
            
        Close #2
        
    Close #1
    
    blnISOSuccess = True
    
    Exit Sub

ErrorHandler:
    MsgBox "CreateISOHeader - modISOHeader" & vbCrLf & vbCrLf & Err & " - " & Err.Description, vbCritical, "Error"
    blnISOSuccess = False
    Close

End Sub
