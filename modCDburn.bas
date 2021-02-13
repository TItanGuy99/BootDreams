Attribute VB_Name = "modCDburn"
Option Explicit

Public Sub Burn_AudioData_CD()

'******************************
'** Audio\Data implementaion **
'******************************
' The dummy file can be an audio dummy or a data dummy. If you're burning CDDA then a _
' data dummy is created, otherwise a audio dummy is created.

On Error GoTo ErrorHandler

    Dim temp As Integer
    Dim msinfo As String
    Dim minfo As String
    Dim strDiscStatus As String
    Dim msinfoString As String
    Dim msinfoStarts As Byte
    Dim msinfoDummy As Long
        
    '***********************
    'MAKE SURE CDRECORD RUNS
    '***********************
    If DoesCDRecordStart = False Then
        MsgBox "Could not start cdrecord, make sure Cygwin is not open.", vbCritical, "Error"
        Exit Sub
    End If
    
    '****************
    'GET CD BURNER ID
    '****************
    Load frmSelectCDBurner
    
    If blnASPI = False Then
        Unload frmSelectCDBurner
        Exit Sub
    End If
    
    If blnDrives = False Then
        Unload frmSelectCDBurner
        Exit Sub
    End If
    
    frmSelectCDBurner.Show vbModal
    If frmSelectCDBurner.Canceled = True Then Exit Sub
    
    'output example 1: _
    /cygdrive/d/source code/bootdreams_105a_src/tools/cdrecord: Input/Output error. test unit ready: scsi sendcmd: no error _
    CDB:  00 00 00 00 00 00 _
    status: 0x2 (CHECK CONDITION) _
    Sense Bytes: 70 00 02 00 00 00 00 0A 00 00 00 00 3A 00 00 00 00 00 _
    Sense Key: 0x2 Not Ready, Segment 0 _
    Sense Code: 0x3A Qual 0x00 (medium not present) Fru 0x0 _
    Sense flags: Blk 0 (not valid) _
    cmd finished after 0.000s timeout 40s _
    /cygdrive/d/source code/bootdreams_105a_src/tools/cdrecord: No disk / Wrong disk!
    
    'output example 2: _
    data type:                standard _
    disk status:              incomplete/appendable _
    session status:           empty _
    BG format status:         none _
    first track:              1 _
    number of sessions:       22 _
    first track in last sess: 22 _
    last track in last sess:  22
    
Retry:
    minfo = ExecuteApp("""" & AppPath$ & "tools\cdrecord.exe"" dev=" & strDrvID & " -minfo")
    
    'is there a disk in the drive?
    If InStr(1, minfo, "No disk") = 0 Then
    
        'yes, parse the disk status
        temp = InStr(1, minfo, "disk status:")
        strDiscStatus = Mid$(minfo, temp + 26, InStr(temp + 26, minfo, vbLf) - (temp + 26))
        
        'is the disc empty?
        If strDiscStatus <> "empty" Then
            'no, did the user click retry?
            If MsgBox("Insert a blank CD-R into the drive and click retry.", vbCritical + vbRetryCancel, "Error") = vbRetry Then
                'yes
                GoTo Retry
            Else
                'no
                Exit Sub
            End If
        End If
        
    Else
    
        'no, did the user click retry?
        If MsgBox("Insert a CD-R in the drive and click retry.", vbCritical + vbRetryCancel, "Error") = vbRetry Then
            'yes
            GoTo Retry
        Else
            'no
            Exit Sub
        End If
    
    End If
    
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
    
        'yes
        'are we burning a selfbooting disc with CDDA?
        If frmMain.mnuExtrasAddCDDATracks.Checked = False Then
        
            'no, so create a (audio) dummy file
            Call MakeDummyFile(AudioDummy, 11702)
            
        End If
        
    End If

    '***************
    'GET CDDA TRACKS
    '***************
    If frmMain.mnuExtrasAddCDDATracks.Checked = True Then
    
        frmSelectCDDATracks.Show vbModal
        If frmSelectCDDATracks.Canceled = True Then Exit Sub
        
    End If
    
    '**********************
    'BURN THE AUDIO SESSION
    '**********************
    'are we burning a selfboot disc with CDDA?
    If frmMain.mnuExtrasAddCDDATracks.Checked = True Then
        
        'yes
        'burn the first session (CDDA)
        Call ShellWait("""" & AppPath$ & "tools\cdrecord.exe"" dev=" & strDrvID & " gracetime=2 -v driveropts=burnfree speed=" & frmMain.cboBurnSpeed.text & " -multi -tao -pad -swab -audio " & strCDDAFilenames, vbNormalFocus)
    
    Else
        
        'no
        'was a dummy file created?
        If lngDummySize <> 0 Then
            
            'yes
            'burn the first session (audio dummy)
            Call ShellWait("""" & AppPath$ & "tools\cdrecord.exe"" dev=" & strDrvID & " gracetime=2 -v driveropts=burnfree speed=" & frmMain.cboBurnSpeed.text & " -multi -tao -pad -swab -audio """ & AppPath$ & "audio.raw""", vbNormalFocus)
        
        Else
            
            'no
            'burn the first session
            Call ShellWait("""" & AppPath$ & "tools\cdrecord.exe"" dev=" & strDrvID & " gracetime=2 -v driveropts=burnfree speed=" & frmMain.cboBurnSpeed.text & " -multi -tao -pad -swab -audio """ & AppPath$ & "tools\audio.raw""", vbNormalFocus)
        
        End If
        
    End If
    
    '***********************
    'DELETE TEMP CDDA TRACKS
    '***********************
    'did we burn CDDA?
    If frmMain.mnuExtrasAddCDDATracks.Checked = True Then
    
        'yes
        Call Kill(AppPath$ & "cdda\temp\*.*")
        Call RmDir(AppPath$ & "cdda\temp")
        
    End If
    
    '**********
    'GET MSINFO
    '**********
    msinfoString = ExecuteApp("""" & AppPath$ & "tools\cdrecord.exe"" dev=" & strDrvID & " -msinfo")
    
    'output example: _
    0,11702
    
    'is the output something other than the msinfo?
    If InStr(1, msinfoString, "cdrecord") <> 0 Then
    
        'yes
        MsgBox "CDRecord could not get the msinfo." & vbCrLf & vbCrLf & msinfoString, vbCritical, "Error"
        Exit Sub
        
    End If
    
    'remove line endings
    msinfo = StripLineEndings$(msinfoString)
    
    '***************
    'MAKE DUMMY FILE
    '***************
    'are we creating a dummy file?
    If lngDummySize <> 0 Then
    
        'yes
        'did we burn CDDA tracks?
        If frmMain.mnuExtrasAddCDDATracks.Checked = True Then
        
            'yes
            'parse the msinfo after the ","
            msinfoStarts = InStr(msinfoString, ",") + 1
            msinfoString = Mid$(msinfoString, msinfoStarts)
            msinfoDummy = msinfoString 'ex: 0,11702 -> 11702
            
            'create a (data) dummy file
            Call MakeDummyFile(DataDummy, msinfoDummy)
            
        End If
        
    End If
    
    '********
    'MAKE ISO
    '********
    'keep the IP.BIN in the filesystem?
    If frmMain.mnuExtrasBootsectorOnly.Checked = False Then
        
        'yes
        'create the ISO
        Call ShellWait("""" & AppPath$ & "tools\mkisofs.exe"" -G """ & AddTrailingSlash$(frmMain.txtFoldername.text) & "IP.BIN"" -C " & msinfo & " " & IIf(frmMain.mnuExtrasISOSettingsJoliet.Checked = True, "-J ", "") & IIf(frmMain.mnuExtrasISOSettingsFullFilenames.Checked = True, "-l ", "") & IIf(frmMain.mnuExtrasISOSettingsRockRidge.Checked = True, "-r ", "") & "-V """ & frmMain.txtCDlabel.text & """ -o """ & AppPath$ & "data.iso"" """ & frmMain.txtFoldername.text & """", vbNormalFocus)
    
    Else
    
        'no
        'move the IP.BIN out of the selfbooting folder
        Name AddTrailingSlash$(frmMain.txtFoldername.text) & "IP.BIN" As AppPath$ & "IP.BIN"
        
        'create the ISO
        Call ShellWait("""" & AppPath$ & "tools\mkisofs.exe"" -G """ & AppPath$ & "IP.BIN"" -C " & msinfo & IIf(frmMain.mnuExtrasISOSettingsJoliet.Checked = True, "-J ", "") & IIf(frmMain.mnuExtrasISOSettingsFullFilenames.Checked = True, "-l ", "") & IIf(frmMain.mnuExtrasISOSettingsRockRidge.Checked = True, "-r ", "") & " -V """ & frmMain.txtCDlabel.text & """ -o """ & AppPath$ & "data.iso"" """ & frmMain.txtFoldername.text & """", vbNormalFocus)
        
        'move the IP.BIN back in selfbooting folder
        Name AppPath$ & "IP.BIN" As AddTrailingSlash$(frmMain.txtFoldername.text) & "IP.BIN"
        
    End If
    
    'was the ISO built?
    If FileExists(AppPath$ & "data.iso") = False Then
    
        'no
        MsgBox "data.iso could not be found", vbCritical, "Error"
        Exit Sub
        
    End If
    
    '********
    'BURN ISO
    '********
    'burn the second session ISO
    Call ShellWait("""" & AppPath$ & "tools\cdrecord.exe"" dev=" & strDrvID & " gracetime=2 -v driveropts=burnfree speed=" & frmMain.cboBurnSpeed.text & " -eject -tao -xa """ & AppPath$ & "data.iso""", vbNormalFocus)
    
    '********
    'FINISHED
    '********
    If MsgBox("The CD was successfully written." & vbCrLf & "Do you want to delete the temporary files?", vbYesNo + vbInformation, "Information") = vbYes Then
        
        'did we create a dummy file?
        If lngDummySize <> 0 Then
            
            'yes
            'did we burn CDDA tracks?
            If frmMain.mnuExtrasAddCDDATracks.Checked = False Then
            
                'no
                'did we burn a Audio\Data selfbooting CD?
                If frmMain.cboDiscFormat.text = "Audio\Data" Then
                
                    'yes
                    Call Kill(AppPath$ & "audio.raw")
                
                End If
            
            End If
        
        End If
        
        Call Kill(AppPath$ & "data.iso")
        
    End If

    Exit Sub

ErrorHandler:
    MsgBox "Burn_AudioData_CD - modCDburn" & vbCrLf & vbCrLf & Err & " - " & Err.Description, vbCritical, "Error"
    
End Sub

Public Sub Burn_DataData_CD()

'*****************************
'** Data\Data implementaion **
'*****************************
' The IP.BIN is NEVER put in the first session, always the second (header ISO).
' The dummy file is ALWAYS a data dummy even with CDDA -- since the game filesystem is _
  in the first session/track, 000DUMMY.DAT is always the first thing burned.
' CDDA is burned in the first session, right after the game filesystem.

On Error GoTo ErrorHandler

    Dim temp As Integer
    Dim strDiscStatus As String
    Dim msinfo As Long
    Dim minfo As String
        
    '***********************
    'MAKE SURE CDRECORD RUNS
    '***********************
    If DoesCDRecordStart = False Then
        MsgBox "Could not start cdrecord, make sure Cygwin is not open.", vbCritical, "Error"
        Exit Sub
    End If
    
    '****************
    'GET CD BURNER ID
    '****************
    Load frmSelectCDBurner
    
    If blnASPI = False Then
        Unload frmSelectCDBurner
        Exit Sub
    End If
    
    If blnDrives = False Then
        Unload frmSelectCDBurner
        Exit Sub
    End If
    
    frmSelectCDBurner.Show vbModal
    If frmSelectCDBurner.Canceled = True Then Exit Sub
    
    'output example 1: _
    /cygdrive/d/source code/bootdreams_105a_src/tools/cdrecord: Input/Output error. test unit ready: scsi sendcmd: no error _
    CDB:  00 00 00 00 00 00 _
    status: 0x2 (CHECK CONDITION) _
    Sense Bytes: 70 00 02 00 00 00 00 0A 00 00 00 00 3A 00 00 00 00 00 _
    Sense Key: 0x2 Not Ready, Segment 0 _
    Sense Code: 0x3A Qual 0x00 (medium not present) Fru 0x0 _
    Sense flags: Blk 0 (not valid) _
    cmd finished after 0.000s timeout 40s _
    /cygdrive/d/source code/bootdreams_105a_src/tools/cdrecord: No disk / Wrong disk!
    
    'output example 2: _
    data type:                standard _
    disk status:              incomplete/appendable _
    session status:           empty _
    BG format status:         none _
    first track:              1 _
    number of sessions:       22 _
    first track in last sess: 22 _
    last track in last sess:  22
    
Retry:
    minfo = ExecuteApp("""" & AppPath$ & "tools\cdrecord.exe"" dev=" & strDrvID & " -minfo")
    
    'is there a disk in the drive?
    If InStr(1, minfo, "No disk") > 0 Then
    
        'no, did the user click retry?
        If MsgBox("Insert a CD-R in the drive and click retry.", vbCritical + vbRetryCancel, "Error") = vbRetry Then
            'yes
            GoTo Retry
        Else
            'no
            Exit Sub
        End If
        
    Else
    
        'yes, parse the disk status
        temp = InStr(1, minfo, "disk status:")
        strDiscStatus = Mid$(minfo, temp + 26, InStr(temp + 26, minfo, vbLf) - (temp + 26))
        
        'is the disc empty?
        If strDiscStatus <> "empty" Then
            'no, did the user click retry?
            If MsgBox("Insert a blank CD-R into the drive and click retry.", vbCritical + vbRetryCancel, "Error") = vbRetry Then
                'yes
                GoTo Retry
            Else
                'no
                Exit Sub
            End If
        End If
    
    End If
    
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
    'GET CDDA TRACKS
    '***************
    If frmMain.mnuExtrasAddCDDATracks.Checked = True Then
    
        frmSelectCDDATracks.Show vbModal
        If frmSelectCDDATracks.Canceled = True Then Exit Sub
        
    End If
    
    '***************
    'MAKE DUMMY FILE
    '***************
    'create a dummy file?
    If lngDummySize <> 0 Then
    
        'yes, burn CDDA tracks?
        If frmMain.mnuExtrasAddCDDATracks.Checked = True Then
        
            'the LBA size of CDDA (pregap included)
            msinfo = ExecuteApp("""" & AppPath$ & "tools\lbacalc.exe"" " & strCDDAFilenames)
            
            'pregap, postgap and header
            msinfo = msinfo + 150 + 150 + 300
            
            'create a data dummy
            Call MakeDummyFile(DataDummy, msinfo)
            
        Else
        
            'no
            Call MakeDummyFile(DataDummy, 11702)
            
        End If
        
    End If
    
    '*********
    'MAKE ISOS
    '*********
    'move the IP.BIN
    If frmMain.mnuExtrasBootsectorOnly.Checked = True Then
        Name AddTrailingSlash$(frmMain.txtFoldername.text) & "IP.BIN" As AppPath$ & "IP.BIN"
    End If
    
    'make the ISO
    Call ShellWait("""" & AppPath$ & "tools\mkisofs.exe"" " & IIf(frmMain.mnuExtrasISOSettingsJoliet.Checked = True, "-J ", "") & IIf(frmMain.mnuExtrasISOSettingsFullFilenames.Checked = True, "-l ", "") & IIf(frmMain.mnuExtrasISOSettingsRockRidge.Checked = True, "-r ", "") & "-V """ & frmMain.txtCDlabel.text & """ -o """ & AppPath$ & "data01.iso"" """ & frmMain.txtFoldername.text & """", vbNormalFocus)
    
    'move the IP.BIN back
    If frmMain.mnuExtrasBootsectorOnly.Checked = True Then
        Name AppPath$ & "IP.BIN" As AddTrailingSlash$(frmMain.txtFoldername.text) & "IP.BIN"
    End If
    
    'make sure the ISO was built
    If FileExists(AppPath$ & "data01.iso") = False Then
        MsgBox "data01.iso could not be found.", vbCritical, "Error"
        Exit Sub
    End If
    
    'create the second session ISO
    Call CreateISOHeader(AppPath$ & "data01.iso", AddTrailingSlash$(frmMain.txtFoldername.text) & "IP.BIN")
    
    'make sure the ISO header was built
    If blnISOSuccess = False Then
        MsgBox "The ISO header creation was not successful.", vbCritical, "Error"
        Exit Sub
    End If
    
    '*******
    'BURN CD
    '*******
    'do we want to burn CDDA tracks?
    If frmMain.mnuExtrasAddCDDATracks.Checked = True Then
        'yes
        Call ShellWait("""" & AppPath$ & "tools\cdrecord.exe"" dev=" & strDrvID & " gracetime=2 -v driveropts=burnfree speed=" & frmMain.cboBurnSpeed.text & " -multi -tao -xa """ & AppPath$ & "data01.iso"" -pad -swab -audio " & strCDDAFilenames, vbNormalFocus)
    Else
        'no
        Call ShellWait("""" & AppPath$ & "tools\cdrecord.exe"" dev=" & strDrvID & " gracetime=2 -v driveropts=burnfree speed=" & frmMain.cboBurnSpeed.text & " -multi -tao -xa """ & AppPath$ & "data01.iso""", vbNormalFocus)
    End If
    
    'burn the second session
    Call ShellWait("""" & AppPath$ & "tools\cdrecord.exe"" dev=" & strDrvID & " gracetime=2 -v driveropts=burnfree speed=" & frmMain.cboBurnSpeed.text & " -eject -tao -xa """ & AppPath$ & "data02.iso""", vbNormalFocus)
        
    '***********************
    'DELETE TEMP CDDA TRACKS
    '***********************
    'did we burn CDDA tracks?
    If frmMain.mnuExtrasAddCDDATracks.Checked = True Then
        
        'yes
        Call Kill(AppPath$ & "cdda\temp\*.*")
        Call RmDir(AppPath$ & "cdda\temp")
        
    End If
        
    '********
    'FINISHED
    '********
    If MsgBox("The CD was successfully written." & vbCrLf & "Do you want to delete the temporary files?", vbYesNo + vbInformation, "Information") = vbYes Then
        Call Kill(AppPath$ & "data01.iso")
        Call Kill(AppPath$ & "data02.iso")
    End If
    
    Exit Sub

ErrorHandler:
    MsgBox "Burn_DataData_CD - modCDburn" & vbCrLf & vbCrLf & Err & " - " & Err.Description, vbCritical, "Error"
    
End Sub

Public Sub Burn_Multisession_CD()

On Error GoTo ErrorHandler
    
    Dim strDiscStatus As String
    Dim temp As String
    Dim msinfo As String
    Dim minfo As String
        
    '***********************
    'MAKE SURE CDRECORD RUNS
    '***********************
    If DoesCDRecordStart = False Then
        MsgBox "Could not start cdrecord, make sure Cygwin is not open.", vbCritical, "Error"
        Exit Sub
    End If
    
    '****************
    'GET CD BURNER ID
    '****************
    Load frmSelectCDBurner
    
    If blnASPI = False Then
        Unload frmSelectCDBurner
        Exit Sub
    End If
    
    If blnDrives = False Then
        Unload frmSelectCDBurner
        Exit Sub
    End If
    
    frmSelectCDBurner.Show vbModal
    If frmSelectCDBurner.Canceled = True Then Exit Sub
    
    'output example 1: _
    /cygdrive/d/source code/bootdreams_105a_src/tools/cdrecord: Input/Output error. test unit ready: scsi sendcmd: no error _
    CDB:  00 00 00 00 00 00 _
    status: 0x2 (CHECK CONDITION) _
    Sense Bytes: 70 00 02 00 00 00 00 0A 00 00 00 00 3A 00 00 00 00 00 _
    Sense Key: 0x2 Not Ready, Segment 0 _
    Sense Code: 0x3A Qual 0x00 (medium not present) Fru 0x0 _
    Sense flags: Blk 0 (not valid) _
    cmd finished after 0.000s timeout 40s _
    /cygdrive/d/source code/bootdreams_105a_src/tools/cdrecord: No disk / Wrong disk!
    
    'output example 2: _
    data type:                standard _
    disk status:              incomplete/appendable _
    session status:           empty _
    BG format status:         none _
    first track:              1 _
    number of sessions:       22 _
    first track in last sess: 22 _
    last track in last sess:  22
        
Retry:
    minfo = ExecuteApp("""" & AppPath$ & "tools\cdrecord.exe"" dev=" & strDrvID & " -minfo")
    
    'is there a disk in the drive?
    If InStr(1, minfo, "No disk") = 0 Then
    
        'yes, parse the disk status
        temp = InStr(1, minfo, "disk status:")
        strDiscStatus = Mid$(minfo, temp + 26, InStr(temp + 26, minfo, vbLf) - (temp + 26))
        
        'is the disc empty?
        If strDiscStatus = "empty" Then
            'yes
            msinfo = "0,0"
        End If
        
    Else
    
        'no, did the user click retry?
        If MsgBox("Insert a CD-R in the drive and click retry.", vbCritical + vbRetryCancel, "Error") = vbRetry Then
            'yes
            GoTo Retry
        Else
            'no
            Exit Sub
        End If
    
    End If
    
    '********
    'MAKE ISO
    '********
    'is the msinfo already set?
    If msinfo = "" Then
    
        'no, so grab it
        msinfo = ExecuteApp("""" & AppPath$ & "tools\cdrecord.exe"" dev=" & strDrvID & " -msinfo")
        
    End If
                
    'output example: _
    178335,186595
            
    'is the output something other than the msinfo?
    If InStr(1, msinfo, "cdrecord") > 0 Then
            
        'yes
        MsgBox "CDRecord could not get the msinfo." & vbCrLf & vbCrLf & msinfo, vbCritical, "Error"
        Exit Sub
        
    End If
    
    msinfo = StripLineEndings$(msinfo)
    
    Call ShellWait("""" & AppPath$ & "tools\mkisofs.exe"" -C " & msinfo & " " & IIf(msinfo = "0,0", "", IIf(frmMain.cbMerge.value = 1, "-M " & strDrvID & " ", "")) & IIf(frmMain.mnuExtrasISOSettingsJoliet.Checked = True, "-J ", "") & IIf(frmMain.mnuExtrasISOSettingsFullFilenames.Checked = True, "-l ", "") & IIf(frmMain.mnuExtrasISOSettingsRockRidge.Checked = True, "-r ", "") & "-V """ & frmMain.txtCDlabel.text & """ -o """ & AppPath$ & "data.iso"" """ & frmMain.txtFoldername.text & """", vbNormalFocus)
    
    'was the ISO built?
    If FileExists(AppPath$ & "data.iso") = False Then
    
        'no
        MsgBox "data.iso could not be found", vbCritical, "Error"
        Exit Sub
        
    End If
    
    '********
    'BURN ISO
    '********
    Call ShellWait("""" & AppPath$ & "tools\cdrecord.exe"" dev=" & strDrvID & " gracetime=2 -v driveropts=burnfree speed=" & frmMain.cboBurnSpeed.text & " -eject -multi -tao " & IIf(frmMain.mnuExtrasISOSettingsImgRecordMode1.Checked = True, "-data ", "-xa ") & """" & AppPath$ & "data.iso""", vbNormalFocus)
    
    '********
    'FINISHED
    '********
    If MsgBox("The multisession CD was successfully written." & vbCrLf & "Do you want to delete the temporary files?", vbYesNo + vbInformation, "Information") = vbYes Then
        Call Kill(AppPath$ & "data.iso")
    End If
        
    Exit Sub

ErrorHandler:
    MsgBox "Burn_Multisession_CD - modCDburn" & vbCrLf & vbCrLf & Err & " - " & Err.Description, vbCritical, "Error"
        
End Sub

Public Sub Burn_DiscJuggler_Image()

On Error GoTo ErrorHandler

    Dim strDiscStatus As String
    Dim minfo As String
    Dim cdiInfo As String
    Dim strSession As String
    Dim temp As Long
    Dim tmp1 As Long
    Dim tmp2 As Long
    Dim lngSessions As Long
    Dim lngTracks As Long
    Dim strTrackType As String
    Dim i As Long, j As Long
    Dim strCDIInfo() As String
        
    '***********************
    'MAKE SURE CDRECORD RUNS
    '***********************
    If DoesCDRecordStart = False Then
        MsgBox "Could not start cdrecord, make sure Cygwin is not open.", vbCritical, "Error"
        Exit Sub
    End If

    '****************
    'GET CD BURNER ID
    '****************
    Load frmSelectCDBurner
    
    If blnASPI = False Then
        Unload frmSelectCDBurner
        Exit Sub
    End If
    
    If blnDrives = False Then
        Unload frmSelectCDBurner
        Exit Sub
    End If
    
    frmSelectCDBurner.Show vbModal
    If frmSelectCDBurner.Canceled = True Then Exit Sub
    
    'output example 1: _
    /cygdrive/d/source code/bootdreams_105a_src/tools/cdrecord: Input/Output error. test unit ready: scsi sendcmd: no error _
    CDB:  00 00 00 00 00 00 _
    status: 0x2 (CHECK CONDITION) _
    Sense Bytes: 70 00 02 00 00 00 00 0A 00 00 00 00 3A 00 00 00 00 00 _
    Sense Key: 0x2 Not Ready, Segment 0 _
    Sense Code: 0x3A Qual 0x00 (medium not present) Fru 0x0 _
    Sense flags: Blk 0 (not valid) _
    cmd finished after 0.000s timeout 40s _
    /cygdrive/d/source code/bootdreams_105a_src/tools/cdrecord: No disk / Wrong disk!
    
    'output example 2: _
    data type:                standard _
    disk status:              incomplete/appendable _
    session status:           empty _
    BG format status:         none _
    first track:              1 _
    number of sessions:       22 _
    first track in last sess: 22 _
    last track in last sess:  22
    
Retry:
    minfo = ExecuteApp("""" & AppPath$ & "tools\cdrecord.exe"" dev=" & strDrvID & " -minfo")
    
    'is there a disk in the drive?
    If InStr(1, minfo, "No disk") = 0 Then
    
        'yes, parse the disk status
        temp = InStr(1, minfo, "disk status:")
        strDiscStatus = Mid$(minfo, temp + 26, InStr(temp + 26, minfo, vbLf) - (temp + 26))
        
        'is the disc empty?
        If strDiscStatus <> "empty" Then
            'no, did the user insert a blank CD-R?
            If MsgBox("Insert a blank CD-R into the drive and click retry.", vbCritical + vbRetryCancel, "Error") = vbRetry Then
                'yes
                GoTo Retry
            Else
                'no
                Exit Sub
            End If
        End If
        
    Else
    
        'no, did the user insert a CD-R?
        If MsgBox("Insert a CD-R in the drive and click retry.", vbCritical + vbRetryCancel, "Error") = vbRetry Then
            'yes
            GoTo Retry
        Else
            'no
            Exit Sub
        End If
    
    End If
    
    '*********************
    'RIP DISCJUGGLER IMAGE
    '*********************
    cdiInfo = ExecuteApp("""" & AppPath$ & "tools\cdirip.exe"" """ & frmMain.txtFilename.text & """ -info")
    
    
    'reinitialize global track count
    temp = 0
    
    'parse the session count from cdiInfo
    tmp1 = InStr(1, cdiInfo, " session(s)")
    tmp2 = InStrRev(cdiInfo, "Found ", tmp1)
    lngSessions = Mid$(cdiInfo, tmp2 + 6, tmp1 - (tmp2 + 6))
    
    'warn about images with more than 2 session
    If lngSessions > 2 Then
        If MsgBox("DiscJuggler images with more than 2 sessions has not been fully tested," & vbCrLf & "but you may continue. Do you wish to continue to burn this image?", vbYesNo + vbExclamation + vbDefaultButton2, "Warning") = vbNo Then
            Exit Sub
        End If
    End If
    
    'initialize array
    ReDim strCDIInfo(lngSessions - 1, 0) As String
    
    'loop the session output in cdiInfo
    For i = 1 To lngSessions
    
        'parse the track count for current session from cdiInfo
        tmp1 = InStr(1, cdiInfo, "Session " & CStr(i) & " has ") + (13 + Len(CStr(i)))
        tmp2 = InStr(tmp1, cdiInfo, " ")
        lngTracks = Mid$(cdiInfo, tmp1, tmp2 - tmp1)
        
        'create an array, first diminsion = number of sessions, 2nd diminsion = highest _
        track count in any session
        If (lngTracks - 1) > UBound(strCDIInfo, 2) Then
            ReDim strCDIInfo(lngSessions - 1, lngTracks - 1) As String
        End If
        
        'loop the track output in cdiInfo
        For j = 1 To lngTracks
        
            'global track number
            temp = temp + 1
            
            'parse the track type output in cdiInfo
            tmp1 = InStr(1, cdiInfo, CStr(temp) & "  Type: ") + (8 + Len(CStr(temp)))
            tmp2 = InStr(tmp1, cdiInfo, "  Size: ")
            
            strTrackType = Mid$(cdiInfo, tmp1, tmp2 - tmp1)
            
            'make sure it's a known track type
            If strTrackType <> "Mode1/2048" And strTrackType <> "Mode2/2336" And strTrackType <> "Audio/2352" Then
            
                MsgBox "Unknown track type: " & strTrackType, vbCritical, "Error"
                Exit Sub
            
            End If
            
            'add current track type to array
            strCDIInfo(i - 1, j - 1) = strTrackType
        
        Next
        
    Next
    
    '//NOTE: cdrecord 2.01.01 a35 > cannot burn data/data images with cdda correctly
    
    'is there more than one track in a session?
    If UBound(strCDIInfo, 2) > 0 Then
        'yes
        'is the current image a cdda data/data image?
        If strCDIInfo(0, 0) = "Mode2/2336" And strCDIInfo(0, 1) = "Audio/2352" Then
            'yes, so warn them about data/data cdda images
            If MsgBox("CDRecord cannot burn data/data DiscJuggler images with CDDA correctly," & vbCrLf & "but you may continue. Note that if the image has little-to-no space left you" & vbCrLf & "MAY burn a coaster. Do you still want to continue?", vbYesNo + vbExclamation + vbDefaultButton2, "Warning") = vbNo Then
                Exit Sub
            End If
        End If
    End If
    
    If FileExists(AppPath$ & "temp") = False Then MkDir AppPath$ & "temp"
    Call ShellWait("""" & AppPath$ & "tools\cdirip.exe"" """ & frmMain.txtFilename.text & """ """ & AppPath$ & "temp"" -iso" & IIf(strCDIInfo(0, 0) <> "Audio/2352", " -cut -cutall", ""), vbNormalFocus)
        
    '*****************************
    'SORT RIPPED DISCJUGGLER IMAGE
    '*****************************
    If FileExists(AppPath$ & "temp\tdisc.cue") = True Then Call Kill(AppPath$ & "temp\tdisc.cue")
    If FileExists(AppPath$ & "temp\tdisc2.cue") = True Then Call Kill(AppPath$ & "temp\tdisc2.cue")
    
    '*****************************
    'BURN RIPPED DISCJUGGLER IMAGE
    '*****************************
    'reinitialize variable
    temp = 0
    
    'loop the sessions
    For i = LBound(strCDIInfo, 1) To UBound(strCDIInfo, 1)
        
        'clear for each session
        strSession = ""
    
        'loop the tracks
        For j = LBound(strCDIInfo, 2) To UBound(strCDIInfo, 2)
        
            'configure cdrecord command line for current track
            Select Case strCDIInfo(i, j)
            
                Case "Mode1/2048"
                    strSession = strSession & "-data """ & AppPath$ & "temp\s" & Format$(i + 1, "00") & "t" & Format$(temp + 1, "00") & ".iso"" "
            
                Case "Mode2/2336"
                    strSession = strSession & "-xa """ & AppPath$ & "temp\s" & Format$(i + 1, "00") & "t" & Format$(temp + 1, "00") & ".iso"" "
                
                Case "Audio/2352"
                    strSession = strSession & "-audio """ & AppPath$ & "temp\s" & Format$(i + 1, "00") & "t" & Format$(temp + 1, "00") & ".wav"" "
            
            End Select
            
            'current global track
            temp = temp + 1
        
        Next
    
        'write the current session to disc
        Call ShellWait("""" & AppPath$ & "tools\cdrecord.exe"" -dev=" & strDrvID & " gracetime=2 -v driveropts=burnfree speed=" & frmMain.cboBurnSpeed.text & " " & IIf(i = UBound(strCDIInfo, 1), "-eject ", "-multi ") & IIf(InStr(1, strSession, "-xa") > 0 Or InStr(1, strSession, "-data") > 0, "-tao ", "-dao ") & strSession, vbNormalFocus)
    
    Next
            
    '********
    'FINISHED
    '********
    If MsgBox("The DiscJuggler image was successfully written." & vbCrLf & "Do you want to delete the temporary files?", vbYesNo + vbInformation, "Information") = vbYes Then
        Call Kill(AppPath$ & "temp\*.*")
        Call RmDir(AppPath$ & "temp")
    End If
    
    Exit Sub

ErrorHandler:
    MsgBox "Burn_DiscJuggler_Image - modCDburn" & vbCrLf & vbCrLf & Err & " - " & Err.Description, vbCritical, "Error"

End Sub

Public Sub Burn_ISO_Image()

On Error GoTo ErrorHandler

    Dim minfo As String
        
    '***********************
    'MAKE SURE CDRECORD RUNS
    '***********************
    If DoesCDRecordStart = False Then
        MsgBox "Could not start cdrecord, make sure Cygwin is not open.", vbCritical, "Error"
        Exit Sub
    End If
    
    '****************
    'GET CD BURNER ID
    '****************
    Load frmSelectCDBurner
    
    If blnASPI = False Then
        Unload frmSelectCDBurner
        Exit Sub
    End If
    
    If blnDrives = False Then
        Unload frmSelectCDBurner
        Exit Sub
    End If
    
    frmSelectCDBurner.Show vbModal
    If frmSelectCDBurner.Canceled = True Then Exit Sub
    
    'output example: _
    /cygdrive/d/source code/bootdreams_105a_src/tools/cdrecord: Input/Output error. test unit ready: scsi sendcmd: no error _
    CDB:  00 00 00 00 00 00 _
    status: 0x2 (CHECK CONDITION) _
    Sense Bytes: 70 00 02 00 00 00 00 0A 00 00 00 00 3A 00 00 00 00 00 _
    Sense Key: 0x2 Not Ready, Segment 0 _
    Sense Code: 0x3A Qual 0x00 (medium not present) Fru 0x0 _
    Sense flags: Blk 0 (not valid) _
    cmd finished after 0.000s timeout 40s _
    /cygdrive/d/source code/bootdreams_105a_src/tools/cdrecord: No disk / Wrong disk!
    
Retry:
    minfo = ExecuteApp("""" & AppPath$ & "tools\cdrecord.exe"" dev=" & strDrvID & " -minfo")
    
    'is there a disk in the drive?
    If InStr(1, minfo, "No disk") > 0 Then
    
        'no, did the user insert a blank CD-R?
        If MsgBox("Insert a CD-R in the drive and click retry.", vbCritical + vbRetryCancel, "Error") = vbRetry Then
            'yes
            GoTo Retry
        Else
            'no
            Exit Sub
        End If
    
    End If
        
    '********
    'BURN ISO
    '********
    Call ShellWait("""" & AppPath$ & "tools\cdrecord.exe"" dev=" & strDrvID & " gracetime=2 -v driveropts=burnfree speed=" & frmMain.cboBurnSpeed.text & " -eject " & IIf(frmMain.chkMultisession.value = 1, "-multi ", "") & " -tao " & IIf(frmMain.mnuExtrasISOSettingsImgRecordMode1.Checked = True, "-data ", "-xa ") & """" & frmMain.txtFilename.text & """", vbNormalFocus)
    
    '********
    'FINISHED
    '********
    MsgBox "The ISO image was successfully written.", vbInformation, "Information"
    
    Exit Sub

ErrorHandler:
    MsgBox "Burn_ISO_Image - modCDburn" & vbCrLf & vbCrLf & Err & " - " & Err.Description, vbCritical, "Error"
        
End Sub

Public Sub Burn_CUE_sheet()

On Error GoTo ErrorHandler

    Dim minfo As String
        
    '***********************
    'MAKE SURE CDRECORD RUNS
    '***********************
    If DoesCDRecordStart = False Then
        MsgBox "Could not start cdrecord, make sure Cygwin is not open.", vbCritical, "Error"
        Exit Sub
    End If
    
    '****************
    'GET CD BURNER ID
    '****************
    Load frmSelectCDBurner
    
    If blnASPI = False Then
        Unload frmSelectCDBurner
        Exit Sub
    End If
    
    If blnDrives = False Then
        Unload frmSelectCDBurner
        Exit Sub
    End If
    
    frmSelectCDBurner.Show vbModal
    If frmSelectCDBurner.Canceled = True Then Exit Sub
    
    'output example: _
    /cygdrive/d/source code/bootdreams_105a_src/tools/cdrecord: Input/Output error. test unit ready: scsi sendcmd: no error _
    CDB:  00 00 00 00 00 00 _
    status: 0x2 (CHECK CONDITION) _
    Sense Bytes: 70 00 02 00 00 00 00 0A 00 00 00 00 3A 00 00 00 00 00 _
    Sense Key: 0x2 Not Ready, Segment 0 _
    Sense Code: 0x3A Qual 0x00 (medium not present) Fru 0x0 _
    Sense flags: Blk 0 (not valid) _
    cmd finished after 0.000s timeout 40s _
    /cygdrive/d/source code/bootdreams_105a_src/tools/cdrecord: No disk / Wrong disk!
    
Retry:
    minfo = ExecuteApp("""" & AppPath$ & "tools\cdrecord.exe"" dev=" & strDrvID & " -minfo")
    
    'is there a disc in the drive?
    If InStr(1, minfo, "No disk") > 0 Then
    
        'no, did the user click retry?
        If MsgBox("Insert a CD-R in the drive and click retry.", vbCritical + vbRetryCancel, "Error") = vbRetry Then
            'yes
            GoTo Retry
        Else
            'no
            Exit Sub
        End If
    
    End If
        
    '**************
    'BURN CUE SHEET
    '**************
    Call ShellWait("""" & AppPath$ & "tools\cdrecord.exe"" dev=" & strDrvID & " gracetime=2 -v driveropts=burnfree speed=" & frmMain.cboBurnSpeed.text & " -eject -dao " & IIf(frmMain.chkMultisession.value = 1, "-multi ", "") & "cuefile=""" & frmMain.txtFilename.text & """", vbNormalFocus)
    
    '********
    'FINISHED
    '********
    MsgBox "The CUE image was successfully written.", vbInformation, "Information"
        
    Exit Sub

ErrorHandler:
    MsgBox "Burn_CUE_sheet - modCDburn" & vbCrLf & vbCrLf & Err & " - " & Err.Description, vbCritical, "Error"
        
End Sub

'this function will make sure cdrecord runs by
'checking if an error about cygwin1.dll is generated
Public Function DoesCDRecordStart() As Boolean

    Dim sOutput As String

    sOutput = ExecuteApp("""" & AppPath$ & "tools\cdrecord.exe"" -version")

    If InStr(1, sOutput, "Schilling") Then
        DoesCDRecordStart = True
    Else
        DoesCDRecordStart = False
    End If

End Function
