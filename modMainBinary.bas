Attribute VB_Name = "modMainBinary"
Option Explicit

Public Enum BinaryType
    ELF = 1
    Katana = 2
    WinCE = 3
    Naomi = 4
    unscrambled = 5
    scrambled = 6
End Enum

Public Sub BinChecker(ByVal sFile As String)
        
On Error GoTo ErrorHandler

    Select Case ScanBinary(sFile)
    
        '//
        '// Main binary is an ELF
        '//
        Case ELF
        
            'convert ELF to binary?
            If MsgBox("The main binary is an ELF, do you want to convert it to binary?", vbYesNo + vbExclamation, "Warning") = vbYes Then
            
                'yes
                Call ShellWait("""" & AppPath$ & "tools\sh-elf-objcopy.exe"" -O binary -R .stack """ & sFile & """ """ & AddTrailingSlash$(frmMain.txtFoldername.text) & "1ST_READ_UNSCRAMBLED.BIN""", vbNormalFocus)
                
                'does the converted binary exist?
                If FileExists(AddTrailingSlash$(frmMain.txtFoldername.text) & "1ST_READ_UNSCRAMBLED.BIN") = False Then
                
                    'no
                    MsgBox "sh-objcopy was unable to convert the ELF.", vbCritical, "Error"
                    blnConvSuccess = False
                    Exit Sub
                    
                Else
                    
                    'yes
                    'delete the ELF
                    Call SetAttr(sFile, vbNormal): Call Kill(sFile)
                    
                    'scramble the binary
                    Call ShellWait("""" & AppPath$ & "tools\scramble.exe"" """ & AddTrailingSlash$(frmMain.txtFoldername.text) & "1ST_READ_UNSCRAMBLED.BIN"" """ & AddTrailingSlash$(frmMain.txtFoldername.text) & "1ST_READ.BIN""", vbNormalFocus)
                    
                    'does the scrambled binary exist?
                    If FileExists(AddTrailingSlash$(frmMain.txtFoldername.text) & "1ST_READ.BIN") = False Then
                    
                        'no
                        MsgBox "scramble was unable to scramble the unscrambled binary.", vbCritical, "Error"
                        blnConvSuccess = False
                        Exit Sub
                        
                    Else
                        
                        'yes
                        'delete the unscrambled binary
                        Call Kill(AddTrailingSlash$(frmMain.txtFoldername.text) & "1ST_READ_UNSCRAMBLED.BIN")
                        'update the main binary filename
                        strMainBinaryFilename = AddTrailingSlash$(frmMain.txtFoldername.text) & "1ST_READ.BIN"
                        
                    End If
                    
                End If
            
                blnConvSuccess = True
                Exit Sub
                
            Else
                
                'no
                blnConvSuccess = False
                Exit Sub
                
            End If
            
        '//
        '// Main binary is a Katana (SEGA) binary
        '//
        Case Katana
        
            'scramble Katana binary?
            If MsgBox("The main binary is a Katana binary. Would you like to scramble it?", vbYesNo + vbExclamation, "Warning") = vbYes Then
                
                'yes
                Call ShellWait("""" & AppPath$ & "tools\scramble.exe"" """ & sFile & """ """ & sFile & "_scrambled.bin""")
                
                'does the scrambled binary exist?
                If FileExists(sFile & "_scrambled.bin") = False Then
                
                    'no
                    MsgBox "scramble was unable to scramble the unscrambled binary.", vbCritical, "Error"
                    blnConvSuccess = False
                    Exit Sub
                    
                Else
                    
                    'yes
                    'delete unscrambled binary
                    Call SetAttr(sFile, vbNormal): Call Kill(sFile)
                    'rename now scramble Katana binary back to original unscrambled Katana binary fn
                    Name sFile & "_scrambled.bin" As sFile
                
                End If
                
            End If
            
            blnConvSuccess = True
            Exit Sub
            
        '//
        '// Main binary is a WinCE (SEGA) binary
        '//
        Case WinCE
        
            'scramble WinCE binary?
            If MsgBox("The main binary is a WinCE binary. Would you like to scramble it?", vbYesNo + vbExclamation, "Warning") = vbYes Then
            
                'yes
                Call ShellWait("""" & AppPath$ & "tools\scramble.exe"" """ & sFile & """ """ & sFile & "_scrambled.bin""")
                                
                'does the scrambled binary exist?
                If FileExists(sFile & "_scrambled.bin") = False Then
                
                    'no
                    MsgBox "scramble was unable to scramble the unscrambled binary.", vbCritical, "Error"
                    blnConvSuccess = False
                    Exit Sub
                    
                Else
                
                    'delete unscrambled binary
                    Call SetAttr(sFile, vbNormal): Call Kill(sFile)
                    'rename now scramble WinCE binary back to original unscrambled WinCE binary fn
                    Name sFile & "_scrambled.bin" As sFile
                
                End If
                
            End If
            
            blnConvSuccess = True
            Exit Sub
            
        '//
        '// Main binary is a Naomi (SEGA) binary
        '//
        Case Naomi
        
            'scramble Naomi binary?
            If MsgBox("The main binary is a Naomi binary. Do you want to scramble it?", vbYesNo + vbExclamation, "Warning") = vbYes Then
                
                'yes
                Call ShellWait("""" & AppPath$ & "tools\scramble.exe"" """ & sFile & """ """ & sFile & "_scrambled.bin""")
                
                'does the scrambled binary exist?
                If FileExists(sFile & "_scrambled.bin") = False Then
                
                    'no
                    MsgBox "scramble was unable to scramble the unscrambled binary.", vbCritical, "Error"
                    blnConvSuccess = False
                    Exit Sub
                    
                Else
                
                    'delete unscrambled binary
                    Call SetAttr(sFile, vbNormal): Call Kill(sFile)
                    'rename now scramble Naomi binary back to original unscrambled Naomi binary fn
                    Name sFile & "_scrambled.bin" As sFile
                
                End If
            
            End If
            
            blnConvSuccess = True
            Exit Sub
            
        '//
        '// Main binary is a unscrambled binary
        '//
        Case unscrambled
        
            'scramble unscrambled binary?
            If MsgBox("The main binary is unscrambled. Would you like to scramble it?", vbYesNo Or vbExclamation, "Warning") = vbYes Then
                    
                'yes
                Call ShellWait("""" & AppPath$ & "tools\scramble.exe"" """ & sFile & """ """ & sFile & "_scrambled.bin""")
                                
                'does the scrambled binary exist?
                If FileExists(sFile & "_scrambled.bin") = False Then
                
                    'no
                    MsgBox "scramble was unable to scramble the unscrambled binary.", vbCritical, "Error"
                    blnConvSuccess = False
                    Exit Sub
                    
                Else
                
                    'delete unscrambled binary
                    Call SetAttr(sFile, vbNormal): Call Kill(sFile)
                    'rename now scramble unscrambled binary back to original unscrambled unscrambled binary fn
                    Name sFile & "_scrambled.bin" As sFile
                
                End If
                
            End If
            
            blnConvSuccess = True
            Exit Sub
            
        '//
        '// Main binary is a scrambled binary
        '//
        Case scrambled
            
            'nothing needs to be done
            blnConvSuccess = True
            Exit Sub
            
    End Select

    Exit Sub

ErrorHandler:
    MsgBox "BinChecker - modBinChecker" & vbCrLf & vbCrLf & Err & " - " & Err.Description, vbCritical, "Error"
    
End Sub

Public Function ScanBinary(ByVal sFile As String) As BinaryType

On Error GoTo ErrorHandler

    Dim sELF As String, sABC As String, sNumber1 As String, sNumber2 As String
    Dim sDreamSNES As String, sBoR As String, sPunch As String, sTetris As String
    Dim sNetBSD As String, sFISA As String, sVMUFrog As String
    Dim sUnknown1 As String, sUnknown2 As String
    Dim sKatana As String, sWinCE As String, sNaomi As String
    
    If FileLen(sFile) < 1024 Then
        ScanBinary = scrambled
        Exit Function
    End If
    
    'unscrambled binary strings
    sELF = Chr(127) & "elf"
    sABC = "abcdefghijklmnopqrstuvwxyz"
    sNumber1 = "1234567890"
    sNumber2 = "0123456789"
    sDreamSNES = "abcdefghijklmnopqrstuvwxyz.0123456789-"
    sBoR = "0123456789abcdef....inf.nan.0123456789abcdef....(null)...": sBoR = Replace(sBoR, ".", Chr(0))
    sNetBSD = "$%&'()*+,-./0123456789:;<=>?@abcdefghijklmnopqrstuvwxyz[\]^_`abcdefghijklmnopqrstuvwxyz{|}~"
    sPunch = "portdev infoenblstatradrtoutdrqcfuncend"
    sTetris = "abcdefghijklefghijklmnopqrstuvwxyz!@#$%^&*()"
    sFISA = "warning: flash read error"
    sVMUFrog = "vmufrog r1, by: the black frog,|http://www.theblackfrog.8m.com": sVMUFrog = Replace(sVMUFrog, "|", Chr(0))
    sUnknown1 = "#...'...*...-.../...2...4...7...9...;...=...?...a...c...e...g...i...j...l...n...o...q...r...t...u...w...x...z...": sUnknown1 = Replace(sUnknown1, ".", Chr(0))
    sUnknown2 = "0123456789abcdef....(null)..0123456789abcdef": sUnknown2 = Replace(sUnknown2, ".", Chr(0))
    sKatana = "shinobi library for dreamcast version "
    sWinCE = "w.i.n.d.o.w.s. .c.e. .k.e.r.n.e.l. .f.o.r. .h.i.t.a.c.h.i. .s.h.": sWinCE = Replace(sWinCE, ".", Chr(0))
    sNaomi = "copyright (c) sega enterprises,ltd." & Chr(0) & "naomi library ver "

    Open sFile For Binary As #1
        ReDim ByteArray(LOF(1) - 1) As Byte
        Get #1, , ByteArray
        sFile = LCase(StrConv(ByteArray, vbUnicode))
    Close #1
    
    If Left$(sFile, 4) = sELF Then
        ScanBinary = ELF
        Exit Function
    End If
    
    If InStr(sFile, sKatana) > 0 Then
        ScanBinary = Katana
        Exit Function
    End If
    
    If InStr(sFile, sWinCE) > 0 Then
        ScanBinary = WinCE
        Exit Function
    End If
    
    If InStr(sFile, sNaomi) > 0 Then
        ScanBinary = Naomi
        Exit Function
    End If
    
    If InStr(sFile, sABC & sNumber1) > 0 Then
        ScanBinary = unscrambled
        Exit Function
    End If
    
    If InStr(sFile, sABC & sNumber2) > 0 Then
        ScanBinary = unscrambled
        Exit Function
    End If
    
    If InStr(sFile, sNumber1 & sABC) > 0 Then
        ScanBinary = unscrambled
        Exit Function
    End If
    
    If InStr(sFile, sNumber2 & sABC) > 0 Then
        ScanBinary = unscrambled
        Exit Function
    End If
    
    If InStr(sFile, sDreamSNES) > 0 Then
        ScanBinary = unscrambled
        Exit Function
    End If
    
    If InStr(sFile, sBoR) > 0 Then
        ScanBinary = unscrambled
        Exit Function
    End If
    
    If InStr(sFile, sNetBSD) > 0 Then
        ScanBinary = unscrambled
        Exit Function
    End If
    
    If InStr(sFile, sTetris) > 0 Then
        ScanBinary = unscrambled
        Exit Function
    End If
    
    If InStr(sFile, sPunch) > 0 Then
        ScanBinary = unscrambled
        Exit Function
    End If
    
    If InStr(sFile, sFISA) > 0 Then
        ScanBinary = unscrambled
        Exit Function
    End If
    
    If InStr(sFile, sVMUFrog) > 0 Then
        ScanBinary = unscrambled
        Exit Function
    End If
    
    If InStr(sFile, sUnknown1) > 0 Then
        ScanBinary = unscrambled
        Exit Function
    End If
    
    If InStr(sFile, sUnknown2) > 0 Then
        ScanBinary = unscrambled
        Exit Function
    End If

    ScanBinary = scrambled

    Exit Function

ErrorHandler:
    MsgBox "ScanBinary - modBinChecker" & vbCrLf & vbCrLf & Err & " - " & Err.Description, vbCritical, "Error"

End Function

'this sub tries to find the main binary filename
'checks for 1ST_READ.BIN, checks the set binary in the IP.BIN, checks if it's _
the only file (excluding the IP.BIN file) in the selfboot folder. If it's still _
not found, the user has to select it theirselfs
Public Sub GetMainBinary()

    strMainBinaryFilename = ""

    'does the current folder have a 1ST_READ.BIN?
    If FileExists(AddTrailingSlash$(frmMain.txtFoldername.text) & "1ST_READ.BIN") = True Then
    
        'yes, so set the main binary
        strMainBinaryFilename = AddTrailingSlash$(frmMain.txtFoldername.text) & "1ST_READ.BIN"
        frmSelectMainBinary.Canceled = False
        Exit Sub
        
    Else
    
        'no
        'does the current folder have a IP.BIN?
        If FileExists(AddTrailingSlash$(frmMain.txtFoldername.text) & "IP.BIN") = True Then
        
            'yes, get the main binary from the IP.BIN file
            Open AddTrailingSlash$(frmMain.txtFoldername.text) & "IP.BIN" For Binary As #1
                
                ReDim btMainBinary(15) As Byte 'create a buffer
                Get #1, 97, btMainBinary 'put the binary field in the IP.BIN into the buffer
                strMainBinaryFilename = AddTrailingSlash$(frmMain.txtFoldername.text) & Trim$(StrConv(btMainBinary, vbUnicode))
                
            Close #1
            
            'does the main binary from the IP.BIN exist?
            If FileExists(strMainBinaryFilename) = False Then
            
                'no
                strMainBinaryFilename = ""
                
                '//
                '// see frmSelectMainBinary.Form_Load; first load
                '//
                Load frmSelectMainBinary
                Unload frmSelectMainBinary
                If frmSelectMainBinary.Canceled = True Then Exit Sub
                
                'was the main binary auto detected?
                If strMainBinaryFilename <> "" Then
                    
                    'yes
                    frmSelectMainBinary.Canceled = False
                    Exit Sub
                    
                Else
                
                    'no
                    MsgBox "The main binary could not be found. Please manually select it.", vbExclamation, "Warning"
                    
                    '//
                    '// see frmSelectMainBinary.Form_Load; second load
                    '//
                    frmSelectMainBinary.Show vbModal
                    If frmSelectMainBinary.Canceled = True Then Exit Sub
                    Exit Sub
                
                End If
            
            Else
            
                'yes, exit with main binary same as binary in IP.BIN
                frmSelectMainBinary.Canceled = False
                Exit Sub
                
            End If
            
        Else
        
            'no
            'see frmSelectMainBinary.Form_Load; first load
            Load frmSelectMainBinary
            Unload frmSelectMainBinary
            If frmSelectMainBinary.Canceled = True Then Exit Sub
            
            'was the main binary auto detected?
            If strMainBinaryFilename = "" Then
            
                'no
                MsgBox "The main binary could not be found. Please manually select it.", vbExclamation, "Warning"
                
                'see frmSelectMainBinary.Form_Load; second load
                frmSelectMainBinary.Show vbModal
                If frmSelectMainBinary.Canceled = True Then Exit Sub
                
            End If
            
        End If
        
    End If

End Sub
