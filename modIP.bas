Attribute VB_Name = "modIP"
Option Explicit

Private Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long

Public Sub MakeIP()

On Error GoTo ErrorHandler

    Dim strHwID         As String * 16
    Dim strMkrID        As String * 16
    Dim strDevInfo      As String * 16
    Dim strAreaSyms     As String * 8
    Dim strPeriphs      As String * 8
    Dim strProdNum      As String * 10
    Dim strProdVer      As String * 6
    Dim strRelDate      As String * 16
    Dim strFiletitle    As String * 16
    Dim strNameComp     As String * 16
    Dim strNameSoft     As String * 128
    Dim ip              As String
    
    'do we need to make a IP.BIN?
    If FileExists(AddTrailingSlash$(frmMain.txtFoldername.text) & "IP.BIN") = True Then
    
        'no
        'but make sure the IP.BIN points to our binary
        
        'open the IP.BIN
        Open AddTrailingSlash$(frmMain.txtFoldername.text) & "IP.BIN" For Binary As #1
        
            'create a buffer
            ReDim btMainBinary(15) As Byte
            
            'put the main binary field into the buffer
            Get #1, 97, btMainBinary
            
        Close #1
        
        'is the IP.BIN pointing to our main binary?
        If Trim$(StrConv(btMainBinary, vbUnicode)) <> UCase$(JustTitle$(strMainBinaryFilename)) Then
            
            'no
            'should BootDreams fix this problem automatically?
            If MsgBox("The IP.BIN is not pointing to the main binary. Would you like to fix this now?", vbYesNo + vbCritical, "Error") = vbYes Then
                
                'yes
                'open the IP.BIN
                Open AddTrailingSlash$(frmMain.txtFoldername.text) & "IP.BIN" For Binary As #1
                
                    'fill the main binary area with spaces
                    Put #1, 97, StrConv(Space$(16), vbFromUnicode)
                    
                    'get the main binary in uppercase
                    strFiletitle = UCase$(JustTitle$(strMainBinaryFilename))
                    
                    'write the new main binary
                    Put #1, 97, strFiletitle
                    
                Close #1
                
                blnIP = True
                Exit Sub
                
            Else
                
                'no, going on would make a coaster
                blnIP = False
                Exit Sub
                
            End If
            
        Else
            
            'yes, nothing to be done
            blnIP = True
            Exit Sub
            
        End If
        
    Else
        
        'yes
        'should BootDreams make a IP.BIN?
        If MsgBox("You are missing a IP.BIN. Would you like to create one?", vbYesNo + vbCritical, "Error") = vbNo Then
        
            'no
            blnIP = False
            Exit Sub
            
        Else
        
            'yes
            'default IP.BIN meta info with VGA box enabled
            strHwID = "SEGA SEGAKATANA"
            strMkrID = "SEGA ENTERPRISES"
            strDevInfo = "B6D8 CD-ROM1/1"
            strAreaSyms = "JUE"
            strPeriphs = "E000010"
            strProdNum = "T0000"
            strProdVer = "V1.000"
            strRelDate = Format$(Now, "yyyymmdd")
            strFiletitle = UCase$(JustTitle$(strMainBinaryFilename)) 'main binary requires uppercase
            strNameComp = "fackue"
            strNameSoft = "BootDreams"
            
            'open the IP.BIN template
            Open AppPath$ & "tools\IP.TMPL" For Binary As #1
            
                'create a buffer
                ReDim btTemplate(32511) As Byte
                
                'put the IP.BIN template into the buffer
                Get #1, 257, btTemplate()
                
                'build a new IP.BIN
                Open AddTrailingSlash$(frmMain.txtFoldername.text) & "IP.BIN" For Binary As #2
                
                    'write the meta info
                    Put #2, , strHwID & _
                              strMkrID & _
                              strDevInfo & _
                              strAreaSyms & _
                              strPeriphs & _
                              strProdNum & _
                              strProdVer & _
                              strRelDate & _
                              strFiletitle & _
                              strNameComp & _
                              strNameSoft
                    
                    'copy the IP.BIN template
                    Put #2, , btTemplate
                    
                Close #2
                
            Close #1
    
            blnIP = True
            
        End If
        
    End If
    
    Exit Sub

ErrorHandler:
    MsgBox "MakeIP - modIP" & vbCrLf & vbCrLf & Err & " - " & Err.Description, vbCritical, "Error"
    Close
    blnIP = False

End Sub

'this sub injects a MR logo into a IP.BIN
Public Sub InsertMRLogo(ByVal mr As String)

On Error GoTo ErrorHandler
    
    'open the IP.BIN
    Open AddTrailingSlash$(frmMain.txtFoldername.text) & "IP.BIN" For Binary As #1
    
        'open the MR logo
        Open mr For Binary As #2
        
            'create a buffer
            ReDim btMR(LOF(2) - 1) As Byte
            
            'put the MR logo into the buffer
            Get #2, , btMR
            
        Close #2
        
        'fill the MR area with zeros
        Put #1, 14369, String$(8192, Chr$(0))
        
        'copy the MR into the IP.BIN
        Put #1, 14369, btMR
        
    Close #1
    
    Exit Sub

ErrorHandler:
    MsgBox "InsertMRLogo - modIP" & vbCrLf & vbCrLf & Err & " - " & Err.Description, vbCritical, "Error"

End Sub

'View MR logo
Public Sub ViewMRLogo(mr As String)
    
On Error GoTo ErrorHandler

    Dim strMRData       As String
    Dim lngMRfs         As Long
    Dim lngMRoffset     As Long
    Dim lngMRwidth      As Long
    Dim lngMRheight     As Long
    Dim lngMRcolors     As Long
    Dim strMRpallete    As String
    Dim strPalette()    As String
    Dim btCompMR()      As Byte
    Dim btUncompMR()    As Byte
    Dim Run             As String
    Dim i As Long, j As Long, k As Long
    Dim x1 As Long, y1 As Long
    Dim R As Byte, G As Byte, b As Byte
    
    'open the MR
    Open mr For Binary As #1
    
        'create a buffer
        ReDim btMR(LOF(1) - 1) As Byte
        
        lngMRfs = LOF(1)
        
        'put the MR into the buffer
        Get #1, , btMR()
        
    Close #1
    
    'convert the MR to something readable
    strMRData = StrConv(btMR(), vbUnicode)
    
    'get bitmap offset
    lngMRoffset = CLng("&H" & Hex$(Asc(Mid$(strMRData, 12, 1))) & String(2 - Len(Hex$(Asc(Mid$(strMRData, 11, 1)))), "0") & Hex$(Asc(Mid$(strMRData, 11, 1))))
    
    'get bitmap resolution
    lngMRwidth = CLng("&H" & Hex$(Asc(Mid$(strMRData, 16, 1))) & String(2 - Len(Hex$(Asc(Mid$(strMRData, 15, 1)))), "0") & Hex$(Asc(Mid$(strMRData, 15, 1))))
    lngMRheight = Asc(Mid$(strMRData, 19, 1))
    
    'get bitmap colors/palette
    lngMRcolors = CLng("&H" & Hex$(Asc(Mid$(strMRData, 28, 1))) & String(2 - Len(Hex$(Asc(Mid$(strMRData, 27, 1)))), "0") & Hex$(Asc(Mid$(strMRData, 27, 1))))
    strMRpallete = Mid$(strMRData, 31, lngMRcolors * 4)
    
    'loop the palette for each color
    For i = 0 To lngMRcolors - 1
    
        'resize palette array
        ReDim Preserve strPalette(i)
        
        'add color to to palette array
        strPalette(i) = Mid$(strMRpallete, i * 4 + 1, 4)
        
    Next
    
    'get compressed bitmap and convert to binary
    btCompMR = StrConv(Mid$(strMRData, lngMRoffset + 1), vbFromUnicode)
    
    'initialize variables
    i = 0
    j = 0
    
    '** CREDIT: Decompression routine ported from kRYPT_'s mrtool **
    
    'create the picture buffer
    ReDim btUncompMR(lngMRwidth * lngMRheight - 1)
    
    'decompress bitmap
    Do
    
        If btCompMR(i) < &H80 Then
        
            btUncompMR(j) = btCompMR(i) 'the bytes lower than 128 are recopied just as they are in the bitmap
            
            j = j + 1
            
            If j > UBound(btUncompMR) Then Exit Do
            
        Else
        
            If (btCompMR(i) = &H82) And (btCompMR(i + 1) >= &H80) Then
            
                'the tag &H82 is followed Nb of points decoded in Run
                Run = btCompMR(i + 1) - &H80 + &H100
                
                For k = 1 To Run
                    btUncompMR(j) = btCompMR(i + 2)  'by retaining only the 1° byte for each point
                    j = j + 1
                    If j > UBound(btUncompMR) Then Exit Do
                Next
                
                i = i + 2
                
            ElseIf btCompMR(i) = &H81 Then
            
                'the tag &H81 is followed of a byte giving Nb of points directly
                Run = btCompMR(i + 1)
                
                For k = 1 To Run
                    btUncompMR(j) = btCompMR(i + 2) 'idem : 1° byte on 2
                    j = j + 1
                    If j > UBound(btUncompMR) Then Exit Do
                Next
                
                i = i + 2
                
            Else
            
                'if > &H82 => code for Nb of points decoded in run
                Run = btCompMR(i) - &H80
                
                For k = 1 To Run
                    btUncompMR(j) = btCompMR(i + 1) 'coded on only one byte
                    j = j + 1
                    If j > UBound(btUncompMR) Then Exit Do
                Next
                
                i = i + 1
                
            End If
            
        End If
        
        i = i + 1
        
    Loop Until i >= (lngMRfs - lngMRoffset - 1)
    
    'resize the MR picture box to MR's resolution
    frmSelectMRLogo.picMR.Width = (lngMRwidth * Screen.TwipsPerPixelX)
    frmSelectMRLogo.picMR.Height = (lngMRheight * Screen.TwipsPerPixelY)
   
    'reinitialize variable
    i = 0
    
    'loop the y coordinates
    For y1 = 0 To lngMRheight - 1
    
        'loop the x coordinates
        For x1 = 0 To lngMRwidth - 1
        
            'get RGB values from palette for current pixel
            If btUncompMR(i) > UBound(strPalette) Then
                'fix for Windows CE
                R = Asc(Mid$(strPalette(0), 3, 1))
                G = Asc(Mid$(strPalette(0), 2, 1))
                b = Asc(Mid$(strPalette(0), 1, 1))
            Else
                R = Asc(Mid$(strPalette(btUncompMR(i)), 3, 1))
                G = Asc(Mid$(strPalette(btUncompMR(i)), 2, 1))
                b = Asc(Mid$(strPalette(btUncompMR(i)), 1, 1))
            End If
            
            'draw the pixel
            SetPixel frmSelectMRLogo.picMR.hdc, x1, y1, RGB(R, G, b)
            
            'next pixel data
            i = i + 1
            
        Next
        
    Next
    
    'rest the drawn pixels in the picture box
    frmSelectMRLogo.picMR.Picture = frmSelectMRLogo.picMR.Image
    frmSelectMRLogo.picMR.Refresh

    Exit Sub

ErrorHandler:
    Close
    MsgBox "ViewMRLogo - modIP" & vbCrLf & vbCrLf & Err & " - " & Err.Description, vbCritical, "Error"
    
End Sub

