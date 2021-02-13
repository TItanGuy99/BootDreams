Attribute VB_Name = "modCDimage"
Option Explicit

Public Enum NeroFormat
    DAO = 1
    TAO = 2
End Enum

Public Sub Create_DiscJuggler_Image()

On Error GoTo ErrorHandler

    Dim sFile As String
    Dim temp As String
    
    '************************
    'GET MAIN BINARY FILENAME
    '************************
    Call GetMainBinary
    If frmSelectMainBinary.Canceled = True Then Exit Sub
    
    '********************
    '1ST_READ.BIN CHECKER
    '********************
    Call BinChecker(strMainBinaryFilename)
    If blnConvSuccess = False Then Exit Sub
    
    '***********
    'MAKE IP.BIN
    '***********
    Call MakeIP
    If blnIP = False Then Exit Sub
    
    '******************
    'INSERT IP.BIN LOGO
    '******************
    'inject a MR logo?
    If frmMain.mnuExtrasInsertMRlogo.Checked = True Then
    
        'yes
        frmSelectMRLogo.Show vbModal
        If frmSelectMRLogo.Canceled = True Then Exit Sub
        
        Call InsertMRLogo(strMRFilename)
        
    End If
    
    '***************
    'MAKE DUMMY FILE
    '***************
    'create a dummy file?
    If lngDummySize <> 0 Then
    
        '//
        '// NOTE: cdi4dc currently does not support a way to use audio dummys
        '//
        
        'yes
        Call MakeDummyFile(DataDummy, 11702)
        
    End If
    
    '****************
    'SAVE FILE DIALOG
    '****************
    sFile = ShowSave("DiscJuggler images (*.cdi)|*.cdi", ".cdi", OVERWRITEPROMPT)
    If sFile = "" Then Exit Sub
    
    '********
    'MAKE ISO
    '********
    'IP.BIN in the bootsector only?
    If frmMain.mnuExtrasBootsectorOnly.Checked = False Then
        
        'no, create the game ISO
        If frmMain.cboDiscFormat.text = "Audio\Data" Then
            Call ShellWait("""" & AppPath$ & "tools\mkisofs.exe"" -G """ & AddTrailingSlash$(frmMain.txtFoldername.text) & "IP.BIN"" -C 0,11702 " & IIf(frmMain.mnuExtrasISOSettingsJoliet.Checked = True, "-J ", "") & IIf(frmMain.mnuExtrasISOSettingsFullFilenames.Checked = True, "-l ", "") & IIf(frmMain.mnuExtrasISOSettingsRockRidge.Checked = True, "-r ", "") & "-V """ & frmMain.txtCDlabel.text & """ -o """ & AppPath$ & "data.iso"" """ & frmMain.txtFoldername.text & """", vbNormalFocus)
        ElseIf frmMain.cboDiscFormat.text = "Data\Data" Then
            Call ShellWait("""" & AppPath$ & "tools\mkisofs.exe"" -G """ & AddTrailingSlash$(frmMain.txtFoldername.text) & "IP.BIN"" " & IIf(frmMain.mnuExtrasISOSettingsJoliet.Checked = True, "-J ", "") & IIf(frmMain.mnuExtrasISOSettingsFullFilenames.Checked = True, "-l ", "") & IIf(frmMain.mnuExtrasISOSettingsRockRidge.Checked = True, "-r ", "") & "-V """ & frmMain.txtCDlabel.text & """ -o """ & AppPath$ & "data.iso"" """ & frmMain.txtFoldername.text & """", vbNormalFocus)
        End If
        
    Else
        
        'yes, move the IP.BIN out of the selfboot folder
        Name AddTrailingSlash$(frmMain.txtFoldername.text) & "IP.BIN" As AppPath$ & "IP.BIN"
        
        'create the game ISO
        If frmMain.cboDiscFormat.text = "Audio\Data" Then
            Call ShellWait("""" & AppPath$ & "tools\mkisofs.exe"" -G """ & AppPath$ & "IP.BIN"" -C 0,11702 " & IIf(frmMain.mnuExtrasISOSettingsJoliet.Checked = True, "-J ", "") & IIf(frmMain.mnuExtrasISOSettingsFullFilenames.Checked = True, "-l ", "") & IIf(frmMain.mnuExtrasISOSettingsRockRidge.Checked = True, "-r ", "") & "-V """ & frmMain.txtCDlabel.text & """ -o """ & AppPath$ & "data.iso"" """ & frmMain.txtFoldername.text & """", vbNormalFocus)
        ElseIf frmMain.cboDiscFormat.text = "Data\Data" Then
            Call ShellWait("""" & AppPath$ & "tools\mkisofs.exe"" -G """ & AppPath$ & "IP.BIN"" " & IIf(frmMain.mnuExtrasISOSettingsJoliet.Checked = True, "-J ", "") & IIf(frmMain.mnuExtrasISOSettingsFullFilenames.Checked = True, "-l ", "") & IIf(frmMain.mnuExtrasISOSettingsRockRidge.Checked = True, "-r ", "") & "-V """ & frmMain.txtCDlabel.text & """ -o """ & AppPath$ & "data.iso"" """ & frmMain.txtFoldername.text & """", vbNormalFocus)
        End If
        
        'move the IP.BIN back to the selfboot folder
        Name AppPath$ & "IP.BIN" As AddTrailingSlash$(frmMain.txtFoldername.text) & "IP.BIN"
    
    End If
    
    'was the ISO built?
    If FileExists(AppPath$ & "data.iso") = False Then
        'no
        MsgBox "data.iso could not be found", vbCritical, "Error"
        Exit Sub
        
    End If
    
    '**********************
    'MAKE DISCJUGGLER IMAGE
    '**********************
    If frmMain.cboDiscFormat.text = "Audio\Data" Then
        Call ShellWait("""" & AppPath$ & "tools\cdi4dc.exe"" """ & AppPath$ & "data.iso"" """ & sFile & """", vbNormalFocus)
    ElseIf frmMain.cboDiscFormat.text = "Data\Data" Then
        Call ShellWait("""" & AppPath$ & "tools\cdi4dc.exe"" """ & AppPath$ & "data.iso"" """ & sFile & """ -d", vbNormalFocus)
    End If
    
    'was the CDI built?
    If FileExists(sFile) = False Then
        'no
        MsgBox "cdi4dc could not create " & JustTitle$(sFile) & ".", vbCritical, "Error"
        Exit Sub
        
    End If
    
    '**********
    'DELETE ISO
    '**********
    Call Kill(AppPath$ & "data.iso")
    
    '********
    'FINISHED
    '********
    MsgBox "The DiscJuggler image was successfully created.", vbInformation, "Information"
    
    Exit Sub

ErrorHandler:
    MsgBox "Create_DiscJuggler_Image - modCDimage" & vbCrLf & vbCrLf & Err & " - " & Err.Description, vbCritical, "Error"

End Sub

Public Sub Create_Nero_Image(ImageFormat As NeroFormat)

On Error GoTo ErrorHandler

    Dim sFile As String
    Dim sFilter As String
    
    '************************
    'GET MAIN BINARY FILENAME
    '************************
    Call GetMainBinary
    If frmSelectMainBinary.Canceled = True Then Exit Sub
    
    '********************
    '1ST_READ.BIN CHECKER
    '********************
    Call BinChecker(strMainBinaryFilename)
    If blnConvSuccess = False Then Exit Sub
    
    '***********
    'MAKE IP.BIN
    '***********
    Call MakeIP
    If blnIP = False Then Exit Sub
    
    '******************
    'INSERT IP.BIN LOGO
    '******************
    'inject a MR logo?
    If frmMain.mnuExtrasInsertMRlogo.Checked = True Then
    
        'yes
        frmSelectMRLogo.Show vbModal
        If frmSelectMRLogo.Canceled = True Then Exit Sub
        
        Call InsertMRLogo(strMRFilename)
        
    End If
    
    '***************
    'MAKE DUMMY FILE
    '***************
    'create a dummy file?
    If lngDummySize <> 0 Then
    
        '//
        '// NOTE: cdi4dc currently does not support a way to use audio dummys
        '//
        
        'yes
        Call MakeDummyFile(DataDummy, 11702)
        
    End If

    '****************
    'SAVE FILE DIALOG
    '****************
    sFile = ShowSave("Nero images (*.nrg)|*.nrg", " " & IIf(ImageFormat = DAO, "DAO", "TAO") & ".nrg", OVERWRITEPROMPT)
    If sFile = "" Then Exit Sub
    
    '********
    'MAKE ISO
    '********
    'IP.BIN in the bootsector only?
    If frmMain.mnuExtrasBootsectorOnly.Checked = False Then
        
        'no, create the game ISO
        If frmMain.cboDiscFormat.text = "Audio\Data" Then
            Call ShellWait("""" & AppPath$ & "tools\mkisofs.exe"" -G """ & AddTrailingSlash$(frmMain.txtFoldername.text) & "IP.BIN"" -C 0,11702 " & IIf(frmMain.mnuExtrasISOSettingsJoliet.Checked = True, "-J ", "") & IIf(frmMain.mnuExtrasISOSettingsFullFilenames.Checked = True, "-l ", "") & IIf(frmMain.mnuExtrasISOSettingsRockRidge.Checked = True, "-r ", "") & "-V """ & frmMain.txtCDlabel.text & """ -o """ & AppPath$ & "data.iso"" """ & frmMain.txtFoldername.text & """", vbNormalFocus)
        ElseIf frmMain.cboDiscFormat.text = "Data\Data" Then
            Call ShellWait("""" & AppPath$ & "tools\mkisofs.exe"" -G """ & AddTrailingSlash$(frmMain.txtFoldername.text) & "IP.BIN"" " & IIf(frmMain.mnuExtrasISOSettingsJoliet.Checked = True, "-J ", "") & IIf(frmMain.mnuExtrasISOSettingsFullFilenames.Checked = True, "-l ", "") & IIf(frmMain.mnuExtrasISOSettingsRockRidge.Checked = True, "-r ", "") & "-V """ & frmMain.txtCDlabel.text & """ -o """ & AppPath$ & "data.iso"" """ & frmMain.txtFoldername.text & """", vbNormalFocus)
        End If
        
    Else
        
        'no, move the IP.BIN out of the selfboot folder
        Name AddTrailingSlash$(frmMain.txtFoldername.text) & "IP.BIN" As AppPath$ & "IP.BIN"
        
        'create the game ISO
        If frmMain.cboDiscFormat.text = "Audio\Data" Then
            Call ShellWait("""" & AppPath$ & "tools\mkisofs.exe"" -G """ & AppPath$ & "IP.BIN"" -C 0,11702 " & IIf(frmMain.mnuExtrasISOSettingsJoliet.Checked = True, "-J ", "") & IIf(frmMain.mnuExtrasISOSettingsFullFilenames.Checked = True, "-l ", "") & IIf(frmMain.mnuExtrasISOSettingsRockRidge.Checked = True, "-r ", "") & "-V """ & frmMain.txtCDlabel.text & """ -o """ & AppPath$ & "data.iso"" """ & frmMain.txtFoldername.text & """", vbNormalFocus)
        ElseIf frmMain.cboDiscFormat.text = "Data\Data" Then
            Call ShellWait("""" & AppPath$ & "tools\mkisofs.exe"" -G """ & AppPath$ & "IP.BIN"" " & IIf(frmMain.mnuExtrasISOSettingsJoliet.Checked = True, "-J ", "") & IIf(frmMain.mnuExtrasISOSettingsFullFilenames.Checked = True, "-l ", "") & IIf(frmMain.mnuExtrasISOSettingsRockRidge.Checked = True, "-r ", "") & "-V """ & frmMain.txtCDlabel.text & """ -o """ & AppPath$ & "data.iso"" """ & frmMain.txtFoldername.text & """", vbNormalFocus)
        End If
        
        'move the IP.BIN back to the selfboot folder
        Name AppPath$ & "IP.BIN" As AddTrailingSlash$(frmMain.txtFoldername.text) & "IP.BIN"
    
    End If
    
    'was the ISO built?
    If FileExists(AppPath$ & "data.iso") = False Then
        'no
        MsgBox "data.iso could not be found", vbCritical, "Error"
        Exit Sub
        
    End If
    
    '**********************
    'MAKE DISCJUGGLER IMAGE
    '**********************
    If frmMain.cboDiscFormat.text = "Audio\Data" Then
        Call ShellWait("""" & AppPath$ & "tools\cdi4dc.exe"" """ & AppPath$ & "data.iso"" """ & AppPath$ & "data.cdi""", vbNormalFocus)
    ElseIf frmMain.cboDiscFormat.text = "Data\Data" Then
        Call ShellWait("""" & AppPath$ & "tools\cdi4dc.exe"" """ & AppPath$ & "data.iso"" """ & AppPath$ & "data.cdi"" -d", vbNormalFocus)
    End If
    
    'was the CDI built?
    If FileExists(AppPath$ & "data.cdi") = False Then
        'no
        MsgBox "cdi4dc could not create data.cdi.", vbCritical, "Error"
        Exit Sub
        
    End If
    
    '**********
    'DELETE ISO
    '**********
    Call Kill(AppPath$ & "data.iso")
    
    '***********************
    'CONVERT TO A NERO IMAGE
    '***********************
    Call ShellWait("""" & AppPath$ & "tools\cdi2nero.exe"" """ & AppPath$ & "data.cdi"" " & IIf(ImageFormat = DAO, "1 ", "2 ") & """" & sFile & """", vbNormalFocus)
    
    'was the NRG built?
    If FileExists(sFile) = False Then
        'no
        MsgBox "cdi2nero could not create " & JustTitle$(sFile) & ".", vbCritical, "Error"
        Exit Sub
        
    End If
    
    '************************
    'DELETE DISCJUGGLER IMAGE
    '************************
    Call Kill(AppPath$ & "data.cdi")
    
    '********
    'FINISHED
    '********
    MsgBox "The Nero (" & IIf(ImageFormat = DAO, "DAO", "TAO") & ") image was successfully created.", vbInformation, "Information"
    
    Exit Sub

ErrorHandler:
    MsgBox "Create_Nero_Image - modCDimage" & vbCrLf & vbCrLf & Err & " - " & Err.Description, vbCritical, "Error"
    
End Sub

Public Sub Create_Alcohol120_Image()

On Error GoTo ErrorHandler

    Dim sFile As String
    Dim sFilter As String
    Dim msinfo As Long
    Dim temp As String
    
    '************************
    'GET MAIN BINARY FILENAME
    '************************
    Call GetMainBinary
    If frmSelectMainBinary.Canceled = True Then Exit Sub
    
    '********************
    '1ST_READ.BIN CHECKER
    '********************
    Call BinChecker(strMainBinaryFilename)
    If blnConvSuccess = False Then Exit Sub
    
    '***********
    'MAKE IP.BIN
    '***********
    Call MakeIP
    If blnIP = False Then Exit Sub
    
    '******************
    'INSERT IP.BIN LOGO
    '******************
    If frmMain.mnuExtrasInsertMRlogo.Checked = True Then
        frmSelectMRLogo.Show vbModal
        If frmSelectMRLogo.Canceled = True Then Exit Sub
        Call InsertMRLogo(strMRFilename)
    End If
    
    '***************
    'MAKE DUMMY FILE
    '***************
    If lngDummySize <> 0 Then
         If frmMain.mnuExtrasAddCDDATracks.Checked = False Then
         
            If frmMain.cboDiscFormat = "Audio\Data" Then
                Call MakeDummyFile(AudioDummy, 11440)
            ElseIf frmMain.cboDiscFormat = "Data\Data" Then
                Call MakeDummyFile(DataDummy, 11702)
            End If
            
        End If
    End If

    '***************
    'GET CDDA TRACKS
    '***************
    If frmMain.mnuExtrasAddCDDATracks.Checked = True Then
        frmSelectCDDATracks.Show vbModal
        If frmSelectCDDATracks.Canceled = True Then Exit Sub
    End If
    
    '****************
    'SAVE FILE DIALOG
    '****************
    sFile = ShowSave("Alcohol 120% images (*.mds)|*.mds", ".mds", OVERWRITEPROMPT)
    If sFile = "" Then Exit Sub
    
    '********
    'MAKE ISO
    '********
    If frmMain.mnuExtrasBootsectorOnly.Checked = False Then
        If frmMain.mnuExtrasAddCDDATracks.Checked = True Then
            If frmMain.cboDiscFormat.text = "Audio\Data" Then
                
                'make the game session
                'Update by Luiz Nai 02/13/2021
                Dim myCommand As String
                myCommand = """" & AppPath$ & "tools\lbacalc.exe"" " & strCDDAFilenames
                msinfo = ExecuteApp(myCommand)
                
                If lngDummySize <> 0 Then
                    Call MakeDummyFile(DataDummy, msinfo)
                End If
                
                Call ShellWait("""" & AppPath$ & "tools\mkisofs.exe"" -G """ & AddTrailingSlash$(frmMain.txtFoldername.text) & "IP.BIN"" -V """ & frmMain.txtCDlabel.text & """ -C 0," & msinfo & " " & IIf(frmMain.mnuExtrasISOSettingsRockRidge.Checked = True, "-r ", "") & IIf(frmMain.mnuExtrasISOSettingsJoliet.Checked = True, "-J ", "") & IIf(frmMain.mnuExtrasISOSettingsFullFilenames.Checked = True, "-l ", "") & "-o """ & AppPath$ & "data.iso"" """ & frmMain.txtFoldername.text & """", vbNormalFocus)
        
            Else
                
                MsgBox "ERROR: MDS4DC does not supported this disc format with CDDA tracks.", vbCritical, "Error"
                Exit Sub
        
            End If
        Else 'no CDDA
        
            'make the game session
        
            If frmMain.cboDiscFormat.text = "Audio\Data" Then
            
                If lngDummySize <> 0 Then
                    msinfo = ExecuteApp("""" & AppPath$ & "tools\lbacalc.exe"" """ & AppPath$ & "audio.raw""")
                    Call ShellWait("""" & AppPath$ & "tools\mkisofs.exe"" -G """ & AddTrailingSlash$(frmMain.txtFoldername.text) & "IP.BIN"" -V """ & frmMain.txtCDlabel.text & """ -C 0," & msinfo & " " & IIf(frmMain.mnuExtrasISOSettingsRockRidge.Checked = True, "-r ", "") & IIf(frmMain.mnuExtrasISOSettingsJoliet.Checked = True, "-J ", "") & IIf(frmMain.mnuExtrasISOSettingsFullFilenames.Checked = True, "-l ", "") & "-o """ & AppPath$ & "data.iso"" """ & frmMain.txtFoldername.text & """", vbNormalFocus)
                Else
                    Call ShellWait("""" & AppPath$ & "tools\mkisofs.exe"" -G """ & AddTrailingSlash$(frmMain.txtFoldername.text) & "IP.BIN"" -V """ & frmMain.txtCDlabel.text & """ -C 0,11702 " & IIf(frmMain.mnuExtrasISOSettingsRockRidge.Checked = True, "-r ", "") & IIf(frmMain.mnuExtrasISOSettingsJoliet.Checked = True, "-J ", "") & IIf(frmMain.mnuExtrasISOSettingsFullFilenames.Checked = True, "-l ", "") & "-o """ & AppPath$ & "data.iso"" """ & frmMain.txtFoldername.text & """", vbNormalFocus)
                End If
                
            ElseIf frmMain.cboDiscFormat.text = "Data\Data" Then
                
                Call ShellWait("""" & AppPath$ & "tools\mkisofs.exe"" -G """ & AddTrailingSlash$(frmMain.txtFoldername.text) & "IP.BIN"" -V """ & frmMain.txtCDlabel.text & """ " & IIf(frmMain.mnuExtrasISOSettingsRockRidge.Checked = True, "-r ", "") & IIf(frmMain.mnuExtrasISOSettingsJoliet.Checked = True, "-J ", "") & IIf(frmMain.mnuExtrasISOSettingsFullFilenames.Checked = True, "-l ", "") & "-o """ & AppPath$ & "data.iso"" """ & frmMain.txtFoldername.text & """", vbNormalFocus)
            
            End If
            
        End If
    Else 'bootsector only
        Name AddTrailingSlash$(frmMain.txtFoldername.text) & "IP.BIN" As AppPath$ & "IP.BIN"
        
        If frmMain.mnuExtrasAddCDDATracks.Checked = True Then
            
            'make the game session
            
            msinfo = ExecuteApp("""" & AppPath$ & "tools\lbacalc.exe"" " & strCDDAFilenames)
            
            If lngDummySize <> 0 Then
                Call MakeDummyFile(DataDummy, msinfo)
            End If
            
            Call ShellWait("""" & AppPath$ & "tools\mkisofs.exe"" -G """ & AppPath$ & "IP.BIN"" -V """ & frmMain.txtCDlabel.text & """ -C 0," & msinfo & " " & IIf(frmMain.mnuExtrasISOSettingsRockRidge.Checked = True, "-r ", "") & IIf(frmMain.mnuExtrasISOSettingsJoliet.Checked = True, "-J ", "") & IIf(frmMain.mnuExtrasISOSettingsFullFilenames.Checked = True, "-l ", "") & "-o """ & AppPath$ & "data.iso"" """ & frmMain.txtFoldername.text & """", vbNormalFocus)
            
        Else
            
            'make the game session
            
            If frmMain.cboDiscFormat.text = "Audio\Data" Then
                
                If lngDummySize <> 0 Then
                    msinfo = ExecuteApp("""" & AppPath$ & "tools\lbacalc.exe"" """ & AppPath$ & "audio.raw""")
                    Call ShellWait("""" & AppPath$ & "tools\mkisofs.exe"" -G """ & AppPath$ & "IP.BIN"" -V """ & frmMain.txtCDlabel.text & """ -C 0," & msinfo & " " & IIf(frmMain.mnuExtrasISOSettingsRockRidge.Checked = True, "-r ", "") & IIf(frmMain.mnuExtrasISOSettingsJoliet.Checked = True, "-J ", "") & IIf(frmMain.mnuExtrasISOSettingsFullFilenames.Checked = True, "-l ", "") & "-o """ & AppPath$ & "data.iso"" """ & frmMain.txtFoldername.text & """", vbNormalFocus)
                Else
                    Call ShellWait("""" & AppPath$ & "tools\mkisofs.exe"" -G """ & AppPath$ & "IP.BIN"" -V """ & frmMain.txtCDlabel.text & """ -C 0,11702 " & IIf(frmMain.mnuExtrasISOSettingsRockRidge.Checked = True, "-r ", "") & IIf(frmMain.mnuExtrasISOSettingsJoliet.Checked = True, "-J ", "") & IIf(frmMain.mnuExtrasISOSettingsFullFilenames.Checked = True, "-l ", "") & "-o """ & AppPath$ & "data.iso"" """ & frmMain.txtFoldername.text & """", vbNormalFocus)
                End If
            
            ElseIf frmMain.cboDiscFormat.text = "Data\Data" Then
                
                Call ShellWait("""" & AppPath$ & "tools\mkisofs.exe"" -G """ & AppPath$ & "IP.BIN"" -V """ & frmMain.txtCDlabel.text & """ " & IIf(frmMain.mnuExtrasISOSettingsRockRidge.Checked = True, "-r ", "") & IIf(frmMain.mnuExtrasISOSettingsJoliet.Checked = True, "-J ", "") & IIf(frmMain.mnuExtrasISOSettingsFullFilenames.Checked = True, "-l ", "") & "-o """ & AppPath$ & "data.iso"" """ & frmMain.txtFoldername.text & """", vbNormalFocus)
            
            End If
        
        End If
        
        'move the IP.BIN back
        Name AppPath$ & "IP.BIN" As AddTrailingSlash$(frmMain.txtFoldername.text) & "IP.BIN"
        
    End If
    
    If FileExists(AppPath$ & "data.iso") = False Then
        MsgBox "data.iso could not be found.", vbCritical, "Error"
        Exit Sub
    End If
    
    '***********************
    'MAKE ALCOHOL 120% IMAGE
    '***********************
    If frmMain.mnuExtrasAddCDDATracks.Checked = True Then
        Call ShellWait("""" & AppPath$ & "tools\mds4dc.exe"" -c """ & sFile & """ """ & AppPath$ & "data.iso"" " & strCDDAFilenames, vbNormalFocus)
    Else
    
        If frmMain.cboDiscFormat.text = "Audio\Data" Then
            
            If lngDummySize <> 0 Then
                Call ShellWait("""" & AppPath$ & "tools\mds4dc.exe"" -c """ & sFile & """ """ & AppPath$ & "data.iso"" """ & AppPath$ & "audio.raw""", vbNormalFocus)
            Else
                Call ShellWait("""" & AppPath$ & "tools\mds4dc.exe"" -a """ & sFile & """ """ & AppPath$ & "data.iso""", vbNormalFocus)
            End If
        
        ElseIf frmMain.cboDiscFormat.text = "Data\Data" Then
            
            Call ShellWait("""" & AppPath$ & "tools\mds4dc.exe"" -d """ & sFile & """ """ & AppPath$ & "data.iso""", vbNormalFocus)
        
        End If
        
    End If
    
    If FileExists(sFile) = False Then
        MsgBox "mds4dc could not create " & JustTitle$(sFile) & ".", vbCritical, "Error"
        Exit Sub
    End If
    
    '*****************
    'DELETE TEMP FILES
    '*****************
    If frmMain.mnuExtrasAddCDDATracks.Checked = True Then
        Call Kill(AppPath$ & "cdda\temp\*.*")
        Call RmDir(AppPath$ & "cdda\temp")
    Else
        If lngDummySize <> 0 Then
            If frmMain.cboDiscFormat = "Audio\Data" Then
                Call Kill(AppPath$ & "audio.raw")
            End If
        End If
    End If
    
    Call Kill(AppPath$ & "data.iso")
    
    '********
    'FINISHED
    '********
    If frmMain.mnuExtrasAddCDDATracks.Checked = True Then
        MsgBox "The Alcohol 120% image with CDDA was successfully created.", vbInformation, "Information"
    Else
        MsgBox "The Alcohol 120% image was successfully created.", vbInformation, "Information"
    End If
    
    Exit Sub

ErrorHandler:
    MsgBox "Create_Alcohol120_Image - modCDimage" & vbCrLf & vbCrLf & Err & " - " & Err.Description, vbCritical, "Error"

End Sub
