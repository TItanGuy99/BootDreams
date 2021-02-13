VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H0041416D&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "BootDreams"
   ClientHeight    =   4590
   ClientLeft      =   9630
   ClientTop       =   5595
   ClientWidth     =   4215
   ClipControls    =   0   'False
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   4590
   ScaleWidth      =   4215
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtOtherInstance 
      Height          =   285
      Left            =   120
      TabIndex        =   15
      Text            =   "txtOtherInstance"
      Top             =   3840
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Frame Frame 
      BackColor       =   &H0041416D&
      Caption         =   "DiscJuggler"
      ForeColor       =   &H00FFFFFF&
      Height          =   2895
      Left            =   120
      OLEDropMode     =   1  'Manual
      TabIndex        =   0
      Top             =   720
      Width           =   3975
      Begin VB.ComboBox cboImageFormat 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "frmMain.frx":0452
         Left            =   2400
         List            =   "frmMain.frx":045C
         OLEDropMode     =   1  'Manual
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   2280
         Width           =   975
      End
      Begin VB.CheckBox cbMerge 
         BackColor       =   &H0041416D&
         Caption         =   "Merge previous session"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   2400
         OLEDropMode     =   1  'Manual
         TabIndex        =   7
         Top             =   2280
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.ComboBox cboBurnSpeed 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "frmMain.frx":046A
         Left            =   2400
         List            =   "frmMain.frx":0492
         OLEDropMode     =   1  'Manual
         TabIndex        =   5
         Text            =   "8"
         Top             =   1440
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox txtCDlabel 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   120
         MaxLength       =   32
         OLEDropMode     =   1  'Manual
         TabIndex        =   4
         Top             =   1440
         Width           =   2175
      End
      Begin VB.CheckBox chkMultisession 
         BackColor       =   &H0041416D&
         Caption         =   "Multisession"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   2400
         OLEDropMode     =   1  'Manual
         TabIndex        =   12
         Top             =   1530
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.ComboBox cboDiscFormat 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "frmMain.frx":04C1
         Left            =   120
         List            =   "frmMain.frx":04CB
         OLEDropMode     =   1  'Manual
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   2280
         Width           =   2175
      End
      Begin VB.TextBox txtFoldername 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   120
         OLEDropMode     =   1  'Manual
         TabIndex        =   3
         Top             =   600
         Width           =   2175
      End
      Begin VB.TextBox txtFilename 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   120
         OLEDropMode     =   1  'Manual
         TabIndex        =   2
         Top             =   600
         Width           =   2175
      End
      Begin VB.Label lblImageFormat 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Image format:"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   2400
         OLEDropMode     =   1  'Manual
         TabIndex        =   14
         Top             =   2040
         Width           =   960
      End
      Begin VB.Label lblBurnSpeedX 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "x"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   3435
         OLEDropMode     =   1  'Manual
         TabIndex        =   10
         Top             =   1560
         Visible         =   0   'False
         Width           =   75
      End
      Begin VB.Label lblDiscFormat 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Disc format:"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   120
         OLEDropMode     =   1  'Manual
         TabIndex        =   11
         Top             =   2040
         Width           =   840
      End
      Begin VB.Label lblBurnSpeed 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Burn speed:"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   2400
         OLEDropMode     =   1  'Manual
         TabIndex        =   9
         Top             =   1200
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label lblCDlabel 
         AutoSize        =   -1  'True
         BackColor       =   &H0041416D&
         BackStyle       =   0  'Transparent
         Caption         =   "CD label:"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   120
         OLEDropMode     =   1  'Manual
         TabIndex        =   8
         Top             =   1200
         Width           =   645
      End
      Begin VB.Image imgBrowse 
         Height          =   345
         Left            =   2640
         OLEDropMode     =   1  'Manual
         Picture         =   "frmMain.frx":04E6
         Top             =   600
         Width           =   1020
      End
      Begin VB.Label lblFilename 
         AutoSize        =   -1  'True
         BackColor       =   &H0041416D&
         BackStyle       =   0  'Transparent
         Caption         =   "Selfboot folder:"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   120
         OLEDropMode     =   1  'Manual
         TabIndex        =   1
         Top             =   360
         Width           =   1065
      End
   End
   Begin VB.Image imgProcess 
      Height          =   525
      Left            =   2640
      OLEDropMode     =   1  'Manual
      Picture         =   "frmMain.frx":0AA6
      Top             =   3840
      Width           =   1350
   End
   Begin VB.Image imgNero 
      Height          =   480
      Left            =   990
      OLEDropMode     =   1  'Manual
      Picture         =   "frmMain.frx":10DF
      Top             =   120
      Width           =   480
   End
   Begin VB.Image imgAlcohol120 
      Height          =   480
      Left            =   1860
      OLEDropMode     =   1  'Manual
      Picture         =   "frmMain.frx":19B9
      Top             =   120
      Width           =   480
   End
   Begin VB.Image imgCDRecord 
      Height          =   480
      Left            =   2729
      OLEDropMode     =   1  'Manual
      Picture         =   "frmMain.frx":2293
      Top             =   120
      Width           =   480
   End
   Begin VB.Image imgImgRecord 
      Height          =   480
      Left            =   3600
      OLEDropMode     =   1  'Manual
      Picture         =   "frmMain.frx":259D
      Top             =   120
      Width           =   480
   End
   Begin VB.Image imgImgRecordGrey 
      Height          =   480
      Left            =   3600
      Picture         =   "frmMain.frx":28B7
      Top             =   120
      Width           =   480
   End
   Begin VB.Image imgCDRecordGrey 
      Height          =   480
      Left            =   2730
      Picture         =   "frmMain.frx":2BC1
      Top             =   120
      Width           =   480
   End
   Begin VB.Image imgAlcohol120Grey 
      Height          =   480
      Left            =   1860
      Picture         =   "frmMain.frx":2ECB
      Top             =   120
      Width           =   480
   End
   Begin VB.Image imgNeroGrey 
      Height          =   480
      Left            =   990
      Picture         =   "frmMain.frx":3795
      Top             =   120
      Width           =   480
   End
   Begin VB.Image imgDiscJuggler 
      Height          =   480
      Left            =   120
      OLEDropMode     =   1  'Manual
      Picture         =   "frmMain.frx":405F
      Top             =   120
      Width           =   480
   End
   Begin VB.Image imgDiscJugglerGrey 
      Height          =   480
      Left            =   120
      Picture         =   "frmMain.frx":4369
      Top             =   120
      Width           =   480
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuExtras 
      Caption         =   "&Extras"
      Begin VB.Menu mnuExtrasISOSettings 
         Caption         =   "ISO settings"
         Begin VB.Menu mnuExtrasISOSettingsImgRecord 
            Caption         =   "Write Mode"
            Begin VB.Menu mnuExtrasISOSettingsImgRecordMode1 
               Caption         =   "Mode &1 (-data)"
               Checked         =   -1  'True
            End
            Begin VB.Menu mnuExtrasISOSettingsImgRecordMode2Form1 
               Caption         =   "Mode &2 Form 1 (-xa)"
            End
         End
         Begin VB.Menu mnuExtrasISOSettingsFullFilenames 
            Caption         =   "&Full filenames (-l)"
            Checked         =   -1  'True
            Shortcut        =   ^L
         End
         Begin VB.Menu mnuExtrasISOSettingsJoliet 
            Caption         =   "&Joliet (-J)"
            Checked         =   -1  'True
            Shortcut        =   ^J
         End
         Begin VB.Menu mnuExtrasISOSettingsRockRidge 
            Caption         =   "&Rock Ridge (-r)"
            Checked         =   -1  'True
            Shortcut        =   ^R
         End
      End
      Begin VB.Menu barrr 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExtrasAddCDDATracks 
         Caption         =   "Add CDDA tracks"
         Shortcut        =   ^T
      End
      Begin VB.Menu mnuExtrasDummyFile 
         Caption         =   "Dummy file"
         Begin VB.Menu mnuExtrasDummyFile650MB 
            Caption         =   "650 MB CD-R"
         End
         Begin VB.Menu mnuExtrasDummyFile700MB 
            Caption         =   "700 MB CD-R"
         End
         Begin VB.Menu mnuExtrasDummyFileCustom 
            Caption         =   "Custom MB"
         End
         Begin VB.Menu mnuExtrasDummyFileNone 
            Caption         =   "None"
         End
      End
      Begin VB.Menu mnuExtrasInsertMRlogo 
         Caption         =   "Insert MR logo"
         Shortcut        =   ^I
      End
      Begin VB.Menu mnuExtrasBootsectorOnly 
         Caption         =   "Bootsector only (IP.BIN)"
         Shortcut        =   ^B
      End
      Begin VB.Menu mnuBar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExtrasAssociate 
         Caption         =   "Associate Extensions"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpHelp 
         Caption         =   "Help"
         Shortcut        =   {F1}
      End
      Begin VB.Menu barrrr 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileCheckforUpdate 
         Caption         =   "Check for Update"
      End
      Begin VB.Menu barr 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()

'On Error GoTo ErrorHandler

    Dim fn As String
        
    Init
    
    If Command$ <> "" Then
        fn = Replace$(Command$, """", "")
        Call openFile(fn)
    End If

    Exit Sub

ErrorHandler:
    MsgBox "Form_Load - frmMain" & vbCrLf & vbCrLf & Err & " - " & Err.Description, vbCritical, "Error"

End Sub

Private Sub Form_Unload(Cancel As Integer)

    Dim tempFrm As Form

    '//
    '// save settings to INI
    '//
    Call WriteINI("Main", "Task", Frame.Caption)
    Call WriteINI("Main", "Disc format", cboDiscFormat.text)
    Call WriteINI("Main", "Burn speed", cboBurnSpeed.text)
    
    'save current ISO burn task
    If mnuExtrasISOSettingsImgRecordMode1.Checked = True Then
        Call WriteINI("Menu", "ISO", "-data")
    ElseIf mnuExtrasISOSettingsImgRecordMode2Form1.Checked = True Then
        Call WriteINI("Menu", "ISO", "-xa")
    End If
    
    Call WriteINI("Menu", "RockRidge", mnuExtrasISOSettingsRockRidge.Checked)
    Call WriteINI("Menu", "Joliet", mnuExtrasISOSettingsJoliet.Checked)
    Call WriteINI("Menu", "Full filenames", mnuExtrasISOSettingsFullFilenames.Checked)
    Call WriteINI("Menu", "Add CDDA tracks", mnuExtrasAddCDDATracks.Checked)
    Call WriteINI("Menu", "Dummy file", CStr(lngDummySize))
    
    Call WriteINI("Menu", "Insert MR logo", mnuExtrasInsertMRlogo.Checked)
    Call WriteINI("Menu", "Bootsector only (IP.BIN)", mnuExtrasBootsectorOnly.Checked)
    Call WriteINI("Nero", "Image format", cboImageFormat.text)
    Call WriteINI("CDRecord", "Merge previous session", cbMerge.value)
    Call WriteINI("ImgRecord", "Multisession", chkMultisession.value)
    
    'close forms
    For Each tempFrm In Forms
        Unload tempFrm
    Next tempFrm

End Sub

Private Sub imgBrowse_Click()

On Error GoTo ErrorHandler
    
    Dim fn As String
    
    If Frame.Caption = "ImgRecord" Then
    
        fn = ShowOpen("All supported (*.cdi; *.iso; *.cue)|*.cdi; *.iso; *.cue|DiscJuggler images (*.cdi)|*.cdi|ISO images (*.iso)|*.iso|" & "CUE sheets (*.cue)|*.cue", FILEMUSTEXIST)
        If fn = "" Then Exit Sub
        
        txtFilename.text = fn
        
        'Multisession checkbox
        Select Case GetExtension$(fn)
            Case "iso": chkMultisession.Visible = True
            Case "cue": chkMultisession.Visible = True
            Case Else: chkMultisession.Visible = False
        End Select
        
    Else
    
        fn = BrowseForFolder
        If fn = "" Then Exit Sub
        
        txtFoldername.text = fn
        
        'set and highlight the CD label
        txtCDlabel.text = JustTitle$(fn)
        txtCDlabel.SetFocus
        txtCDlabel.SelStart = 0
        txtCDlabel.SelLength = Len(txtCDlabel.text)
        
    End If

    Exit Sub

ErrorHandler:
    MsgBox "imgBrowse_Click - frmMain" & vbCrLf & vbCrLf & Err & " - " & Err.Description, vbCritical, "Error"

End Sub

Private Sub imgProcess_Click()
        
    '//
    '// DiscJuggler
    '//
    If Frame.Caption = "DiscJuggler" Then
    
        If FileExists(txtFoldername.text) = False Then
            MsgBox "The current folder does not exist.", vbCritical, "Error"
            Exit Sub
        End If
        
        If txtCDlabel.text = "" Then
            MsgBox "Please type in a CD label before continuing.", vbCritical, "Error"
            Exit Sub
        End If
        
        If MsgBox("Are you sure you want to create a DiscJuggler image?", vbYesNo + vbQuestion, "Question") = vbNo Then Exit Sub
        
        Create_DiscJuggler_Image
        
    '//
    '// Nero
    '//
    ElseIf Frame.Caption = "Nero" Then

        If FileExists(txtFoldername.text) = False Then
            MsgBox "The current folder does not exist.", vbCritical, "Error"
            Exit Sub
        End If
        
        If txtCDlabel.text = "" Then
            MsgBox "Please type in the CD label before continuing.", vbCritical, "Error"
            Exit Sub
        End If
        
        If MsgBox("Are you sure you want to create a Nero image?", vbYesNo + vbQuestion, "Question") = vbNo Then Exit Sub
        
        If frmMain.cboImageFormat.text = "DAO" Then
            Create_Nero_Image DAO
        Else
            Create_Nero_Image TAO
        End If
        
    '//
    '// Alcohol 120%
    '//
    ElseIf Frame.Caption = "Alcohol 120%" Then

        If FileExists(txtFoldername.text) = False Then
            MsgBox "The current folder does not exist.", vbCritical, "Error"
            Exit Sub
        End If
        
        If txtCDlabel.text = "" Then
            MsgBox "Please type in the CD label before continuing.", vbCritical, "Error"
            Exit Sub
        End If
        
        If mnuExtrasAddCDDATracks.Checked = True Then
            If MsgBox("Are you sure you want to create a Alcohol 120% image with CDDA?", vbYesNo + vbQuestion, "Question") = vbNo Then Exit Sub
        Else
            If MsgBox("Are you sure you want to create a Alcohol 120% image?", vbYesNo + vbQuestion, "Question") = vbNo Then Exit Sub
        End If
        
        Create_Alcohol120_Image
    
    '//
    '// CDRecord
    '//
    ElseIf Frame.Caption = "CDRecord" Then

        If FileExists(txtFoldername.text) = False Then
            MsgBox "The current folder does not exist.", vbCritical, "Error"
            Exit Sub
        End If
        
        If txtCDlabel.text = "" Then
            MsgBox "Please type in the CD label before continuing.", vbCritical, "Error"
            Exit Sub
        End If
        
        If cboDiscFormat.text = "Audio\Data" Then
        
            If mnuExtrasAddCDDATracks.Checked = True Then
                If MsgBox("Are you sure you want to burn a selfbooting CD with CDDA?", vbYesNo + vbQuestion, "Question") = vbNo Then
                    Exit Sub
                End If
            Else
                If MsgBox("Are you sure you want to burn a selfbooting CD?", vbYesNo + vbQuestion, "Question") = vbNo Then
                    Exit Sub
                End If
            End If
            
            Burn_AudioData_CD
            
        ElseIf cboDiscFormat.text = "Data\Data" Then
        
            If mnuExtrasAddCDDATracks.Checked = True Then
                If MsgBox("Are you sure you want to burn a selfbooting CD with CDDA?", vbYesNo + vbQuestion, "Question") = vbNo Then
                    Exit Sub
                End If
            Else
                If MsgBox("Are you sure you want to burn a selfbooting CD?", vbYesNo + vbQuestion, "Question") = vbNo Then
                    Exit Sub
                End If
            End If
            
            Burn_DataData_CD
            
        ElseIf cboDiscFormat.text = "Multisession" Then
        
            If MsgBox("Are you sure you want to burn a multisession CD?", vbYesNo + vbQuestion, "Question") = vbYes Then
                Burn_Multisession_CD
            End If
            
        End If
        
    '//
    '// ImgRecord
    '//
    ElseIf Frame.Caption = "ImgRecord" Then
    
        If FileExists(txtFilename.text) = False Then
            MsgBox "The current file does not exist.", vbCritical, "Error"
            Exit Sub
        End If
        
        Select Case GetExtension$(txtFilename.text)
            
            Case "cdi"
                If MsgBox("Are you sure you want to burn a DiscJuggler image?", vbYesNo + vbQuestion, "Question") = vbYes Then
                    Burn_DiscJuggler_Image
                End If
            
            Case "iso"
                If MsgBox("Are you sure you want to burn a ISO image?", vbYesNo + vbQuestion, "Question") = vbYes Then
                    Burn_ISO_Image
                End If
            
            Case "cue"
                If MsgBox("Are you sure you want to burn a CUE image?", vbYesNo + vbQuestion, "Question") = vbYes Then
                    Burn_CUE_sheet
                End If
            
            Case ""
                MsgBox "Extensionless files are not supported.", vbCritical, "Error"
                Exit Sub
            
            Case Else
                MsgBox UCase$(GetExtension$(txtFilename.text)) & " extensions currently are not supported.", vbCritical, "Error"
                Exit Sub
            
        End Select
        
    End If

End Sub

Private Sub imgDiscJugglerGrey_Click()

    Call imgDiscJuggler_Click

End Sub

Private Sub imgDiscJuggler_Click()

    'labels
    Frame.Caption = "DiscJuggler"
    lblFilename.Caption = "Seflboot folder:"
    
    'resize
    Frame.Height = 2895
    imgProcess.Top = 3840
    Me.Height = 5280
    
    'objects
    imgDiscJuggler.Visible = True
    imgDiscJugglerGrey.Visible = False
    imgNero.Visible = False
    imgNeroGrey.Visible = True
    imgAlcohol120.Visible = False
    imgAlcohol120Grey.Visible = True
    imgCDRecord.Visible = False
    imgCDRecordGrey.Visible = True
    imgImgRecord.Visible = False
    imgImgRecordGrey.Visible = True
    
    txtFoldername.Visible = True
    lblCDlabel.Visible = True
    txtCDlabel.Visible = True
    mnuExtrasISOSettings.Enabled = True
    mnuExtrasISOSettingsImgRecord.Enabled = False
    mnuExtrasDummyFile.Enabled = True
    mnuExtrasInsertMRlogo.Enabled = True
    mnuExtrasBootsectorOnly.Enabled = True
    mnuExtrasAddCDDATracks.Enabled = False
    lblDiscFormat.Visible = True
    cboDiscFormat.Visible = True
    cboDiscFormat.Enabled = True
    lblBurnSpeed.Visible = False
    cboBurnSpeed.Visible = False
    lblBurnSpeedX.Visible = False
    cbMerge.Visible = False
    txtFilename.Visible = False
    chkMultisession.Visible = False
    cboImageFormat.Visible = False
    lblImageFormat.Visible = False
    
    If cboDiscFormat.text = "Multisession" Then
        cboDiscFormat.text = "Audio\Data"
    End If
    
    If cboDiscFormat.ListCount = 3 Then
        cboDiscFormat.RemoveItem (2)
    End If

End Sub

Private Sub imgNeroGrey_Click()

    Call imgNero_Click

End Sub

Private Sub imgNero_Click()
    
    'labels
    Frame.Caption = "Nero"
    lblFilename.Caption = "Seflboot folder:"
    
    'resize
    Frame.Height = 2895
    imgProcess.Top = 3840
    Me.Height = 5280
    
    'objects
    imgDiscJuggler.Visible = False
    imgDiscJugglerGrey.Visible = True
    imgNero.Visible = True
    imgNeroGrey.Visible = False
    imgAlcohol120.Visible = False
    imgAlcohol120Grey.Visible = True
    imgCDRecord.Visible = False
    imgCDRecordGrey.Visible = True
    imgImgRecord.Visible = False
    imgImgRecordGrey.Visible = True
    
    txtFoldername.Visible = True
    lblCDlabel.Visible = True
    txtCDlabel.Visible = True
    mnuExtrasISOSettings.Enabled = True
    mnuExtrasISOSettingsImgRecord.Enabled = False
    mnuExtrasDummyFile.Enabled = True
    mnuExtrasInsertMRlogo.Enabled = True
    mnuExtrasBootsectorOnly.Enabled = True
    mnuExtrasAddCDDATracks.Enabled = False
    lblDiscFormat.Visible = True
    cboDiscFormat.Visible = True
    cboDiscFormat.Enabled = True
    lblBurnSpeed.Visible = False
    cboBurnSpeed.Visible = False
    lblBurnSpeedX.Visible = False
    cbMerge.Visible = False
    txtFilename.Visible = False
    chkMultisession.Visible = False
    cboImageFormat.Visible = True
    lblImageFormat.Visible = True
    
    If cboDiscFormat.text = "Multisession" Then
        cboDiscFormat.text = "Audio\Data"
    End If
    
    If cboDiscFormat.ListCount = 3 Then
        cboDiscFormat.RemoveItem (2)
    End If

End Sub

Private Sub imgAlcohol120Grey_Click()

    Call imgAlcohol120_Click

End Sub

Private Sub imgAlcohol120_Click()

    'labels
    Frame.Caption = "Alcohol 120%"
    lblFilename.Caption = "Seflboot folder:"
    
    'resize
    Frame.Height = 2895
    imgProcess.Top = 3840
    Me.Height = 5280
    
    'objects
    imgDiscJuggler.Visible = False
    imgDiscJugglerGrey.Visible = True
    imgNero.Visible = False
    imgNeroGrey.Visible = True
    imgAlcohol120.Visible = True
    imgAlcohol120Grey.Visible = False
    imgCDRecord.Visible = False
    imgCDRecordGrey.Visible = True
    imgImgRecord.Visible = False
    imgImgRecordGrey.Visible = True
    
    txtFoldername.Visible = True
    lblCDlabel.Visible = True
    txtCDlabel.Visible = True
    mnuExtrasISOSettings.Enabled = True
    mnuExtrasISOSettingsImgRecord.Enabled = False
    mnuExtrasAddCDDATracks.Enabled = True
    mnuExtrasDummyFile.Enabled = True
    mnuExtrasInsertMRlogo.Enabled = True
    mnuExtrasBootsectorOnly.Enabled = True
    lblDiscFormat.Visible = True
    cboDiscFormat.Visible = True
    lblBurnSpeed.Visible = False
    cboBurnSpeed.Visible = False
    lblBurnSpeedX.Visible = False
    cbMerge.Visible = False
    txtFilename.Visible = False
    chkMultisession.Visible = False
    cboImageFormat.Visible = False
    lblImageFormat.Visible = False
    
    If cboDiscFormat.text = "Multisession" Then
        cboDiscFormat.text = "Audio\Data"
    End If
    
    If cboDiscFormat.ListCount = 3 Then
        cboDiscFormat.RemoveItem (2)
    End If
    
    If mnuExtrasAddCDDATracks.Checked = True Then
        cboDiscFormat.text = "Audio\Data"
        cboDiscFormat.Enabled = False
    Else
        cboDiscFormat.Enabled = True
    End If

End Sub

Private Sub imgCDRecordGrey_Click()

    Call imgCDRecord_Click

End Sub

Private Sub imgCDRecord_Click()

    'labels
    Frame.Caption = "CDRecord"
    lblFilename.Caption = "Seflboot folder:"
    
    'resize
    Frame.Height = 2895
    imgProcess.Top = 3840
    Me.Height = 5280
    lblBurnSpeed.Left = 2400
    cboBurnSpeed.Left = 2400
    lblBurnSpeedX.Left = 3440
    
    'objects
    imgDiscJuggler.Visible = False
    imgDiscJugglerGrey.Visible = True
    imgNero.Visible = False
    imgNeroGrey.Visible = True
    imgAlcohol120.Visible = False
    imgAlcohol120Grey.Visible = True
    imgCDRecord.Visible = True
    imgCDRecordGrey.Visible = False
    imgImgRecord.Visible = False
    imgImgRecordGrey.Visible = True
    
    txtFoldername.Visible = True
    lblCDlabel.Visible = True
    txtCDlabel.Visible = True
    lblBurnSpeed.Visible = True
    cboBurnSpeed.Visible = True
    lblBurnSpeedX.Visible = True
    lblDiscFormat.Visible = True
    cboDiscFormat.Enabled = True
    cboDiscFormat.Visible = True
    mnuExtrasISOSettings.Enabled = True
    txtFilename.Visible = False
    cboImageFormat.Visible = False
    lblImageFormat.Visible = False
    chkMultisession.Visible = False
    
    If cboDiscFormat.ListCount = 2 Then
        cboDiscFormat.AddItem "Multisession"
    End If
    
    If cboDiscFormat.text = "Multisession" Then
        cbMerge.Visible = True
        mnuExtrasAddCDDATracks.Enabled = False
        mnuExtrasISOSettingsImgRecord.Enabled = True
        mnuExtrasDummyFile.Enabled = False
        mnuExtrasInsertMRlogo.Enabled = False
        mnuExtrasBootsectorOnly.Enabled = False
    Else
        cbMerge.Visible = False
        mnuExtrasAddCDDATracks.Enabled = True
        mnuExtrasISOSettingsImgRecord.Enabled = False
        mnuExtrasDummyFile.Enabled = True
        mnuExtrasInsertMRlogo.Enabled = True
        mnuExtrasBootsectorOnly.Enabled = True
    End If
    
End Sub

Private Sub imgImgRecordGrey_Click()

    Call imgImgRecord_Click

End Sub

Private Sub imgImgRecord_Click()

    'labels
    Frame.Caption = "ImgRecord"
    lblFilename.Caption = "CD image:"
    
    'resize
    Frame.Height = 2055
    imgProcess.Top = 3000
    Me.Height = 4425
    cboBurnSpeed.Left = 120
    lblBurnSpeed.Left = 120
    lblBurnSpeedX.Left = 1160
    
    'objects
    imgDiscJuggler.Visible = False
    imgDiscJugglerGrey.Visible = True
    imgNero.Visible = False
    imgNeroGrey.Visible = True
    imgAlcohol120.Visible = False
    imgAlcohol120Grey.Visible = True
    imgCDRecord.Visible = False
    imgCDRecordGrey.Visible = True
    imgImgRecord.Visible = True
    imgImgRecordGrey.Visible = False
    
    txtFilename.Visible = True
    lblBurnSpeed.Visible = True
    cboBurnSpeed.Visible = True
    lblBurnSpeedX.Visible = True
    lblDiscFormat.Visible = False
    cboDiscFormat.Visible = False
    cbMerge.Visible = False
    txtFoldername.Visible = False
    lblCDlabel.Visible = False
    txtCDlabel.Visible = False
    mnuExtrasAddCDDATracks.Enabled = False
    mnuExtrasDummyFile.Enabled = False
    mnuExtrasInsertMRlogo.Enabled = False
    mnuExtrasBootsectorOnly.Enabled = False
    cboImageFormat.Visible = False
    lblImageFormat.Visible = False

    Select Case GetExtension$(txtFilename.text)
        Case "iso"
            chkMultisession.Visible = True
            mnuExtrasISOSettings.Enabled = True
        Case "cue"
            chkMultisession.Visible = True
            mnuExtrasISOSettings.Enabled = False
        Case "cdi"
            chkMultisession.Visible = False
            mnuExtrasISOSettings.Enabled = False
    End Select

End Sub

Private Sub cboDiscFormat_Click()

    If cboDiscFormat.text = "Audio\Data" Then
    
        If Frame.Caption = "CDRecord" Or Frame.Caption = "Alcohol 120%" Then
            mnuExtrasAddCDDATracks.Enabled = True
        Else
            mnuExtrasAddCDDATracks.Enabled = False
        End If
        
        cbMerge.Visible = False
        mnuExtrasISOSettingsImgRecord.Enabled = False
        mnuExtrasISOSettings.Enabled = True
        mnuExtrasDummyFile.Enabled = True
        mnuExtrasInsertMRlogo.Enabled = True
        mnuExtrasBootsectorOnly.Enabled = True
        
    ElseIf cboDiscFormat.text = "Data\Data" Then
    
        If Frame.Caption = "CDRecord" Or Frame.Caption = "Alcohol 120%" Then
            mnuExtrasAddCDDATracks.Enabled = True
        Else
            mnuExtrasAddCDDATracks.Enabled = False
        End If
        
        cbMerge.Visible = False
        mnuExtrasISOSettingsImgRecord.Enabled = False
        mnuExtrasISOSettings.Enabled = True
        mnuExtrasDummyFile.Enabled = True
        mnuExtrasInsertMRlogo.Enabled = True
        mnuExtrasBootsectorOnly.Enabled = True
        
    ElseIf cboDiscFormat.text = "Multisession" Then
    
        cbMerge.Visible = True
        mnuExtrasISOSettingsImgRecord.Enabled = True
        mnuExtrasISOSettings.Enabled = True
        mnuExtrasAddCDDATracks.Enabled = False
        mnuExtrasDummyFile.Enabled = False
        mnuExtrasInsertMRlogo.Enabled = False
        mnuExtrasBootsectorOnly.Enabled = False
        
    End If

End Sub

Private Sub mnuExtrasAddCDDATracks_Click()

    mnuExtrasAddCDDATracks.Checked = Not mnuExtrasAddCDDATracks.Checked

    'current task alcohol 120?
    If Frame.Caption = "Alcohol 120%" Then
        
        'yes, is cdda tracks checked?
        If mnuExtrasAddCDDATracks.Checked = True Then
        
            '//NOTE: mds4dc currently only supports audio\data cdda images
        
            'yes, set and disable disc format to audio\data
            cboDiscFormat.text = "Audio\Data"
            cboDiscFormat.Enabled = False
            
        Else
        
            'no, enable disc format
            cboDiscFormat.Enabled = True
            
        End If
    
    End If

End Sub

Private Sub mnuExtrasDummyFile650MB_Click()

    lngDummySize = 650

    mnuExtrasDummyFile650MB.Checked = True
    mnuExtrasDummyFile700MB.Checked = False
    mnuExtrasDummyFileCustom.Checked = False
    mnuExtrasDummyFileNone.Checked = False

End Sub

Private Sub mnuExtrasDummyFile700MB_Click()

    lngDummySize = 700

    mnuExtrasDummyFile650MB.Checked = False
    mnuExtrasDummyFile700MB.Checked = True
    mnuExtrasDummyFileCustom.Checked = False
    mnuExtrasDummyFileNone.Checked = False

End Sub

Private Sub mnuExtrasDummyFileCustom_Click()

    'check the dummy size so we don't show the dummy form when first running BootDreams
    If blnStartup = False Then
        frmDummySize.Show vbModal
        If frmDummySize.Canceled = True Then Exit Sub
    
        'default to the already available dummy sizes
        Select Case lngDummySize
            Case 650
                mnuExtrasDummyFile650MB_Click
                Exit Sub
            Case 700
                mnuExtrasDummyFile700MB_Click
                Exit Sub
        End Select
    End If
    
    
    mnuExtrasDummyFile650MB.Checked = False
    mnuExtrasDummyFile700MB.Checked = False
    mnuExtrasDummyFileCustom.Checked = True
    mnuExtrasDummyFileNone.Checked = False

End Sub

Private Sub mnuExtrasDummyFileNone_Click()

    lngDummySize = 0

    mnuExtrasDummyFile650MB.Checked = False
    mnuExtrasDummyFile700MB.Checked = False
    mnuExtrasDummyFileCustom.Checked = False
    mnuExtrasDummyFileNone.Checked = True

End Sub

Private Sub mnuExtrasISOSettingsImgRecordMode1_Click()

    mnuExtrasISOSettingsImgRecordMode1.Checked = True
    mnuExtrasISOSettingsImgRecordMode2Form1.Checked = False

End Sub

Private Sub mnuExtrasISOSettingsImgRecordMode2Form1_Click()

    mnuExtrasISOSettingsImgRecordMode1.Checked = False
    mnuExtrasISOSettingsImgRecordMode2Form1.Checked = True

End Sub

Private Sub mnuExtrasISOSettingsRockRidge_Click()

    mnuExtrasISOSettingsRockRidge.Checked = Not mnuExtrasISOSettingsRockRidge.Checked

End Sub

Private Sub mnuExtrasISOSettingsJoliet_Click()

    mnuExtrasISOSettingsJoliet.Checked = Not mnuExtrasISOSettingsJoliet.Checked

End Sub

Private Sub mnuExtrasISOSettingsfullfilenames_Click()

    mnuExtrasISOSettingsFullFilenames.Checked = Not mnuExtrasISOSettingsFullFilenames.Checked

End Sub

Private Sub mnuExtrasInsertMRLogo_Click()

    mnuExtrasInsertMRlogo.Checked = Not mnuExtrasInsertMRlogo.Checked

End Sub

Private Sub mnuExtrasBootsectorOnly_Click()

    mnuExtrasBootsectorOnly.Checked = Not mnuExtrasBootsectorOnly.Checked

End Sub

Private Sub mnuExtrasAssociate_Click()

    frmAssociation.Show vbModal

End Sub

Private Sub mnuFileExit_Click()

    Unload Me

End Sub

Private Sub mnuHelpHelp_Click()

    If FileExists(AppPath$ & "BootDreams.chm") = True Then
    
        Call ShellExecute(0, vbNullString, AppPath$ & "BootDreams.chm", vbNullString, vbNullString, vbNormalFocus)
        
    Else
    
        MsgBox "The BootDreams help file is not available.", vbCritical, "Error"
        
    End If

End Sub

Private Sub mnuHelpAbout_Click()

    Call MsgBox("A Dreamcast selfboot frontend for Windows." _
      & vbCrLf & vbCrLf & "See the readme for credit information." _
      & vbCrLf & "" _
      & vbCrLf & "Created by fackue" _
      & vbCrLf & "http://dchelp.dcemulation.com/" _
      , , "About")

End Sub

Private Sub mnuFileCheckforUpdate_Click()

    CheckVersion

End Sub

Private Sub Form_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Call HandleDragDrop(Data.Files(1))
    
End Sub

Private Sub Frame_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Call HandleDragDrop(Data.Files(1))
    
End Sub

Private Sub imgDiscJuggler_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Call HandleDragDrop(Data.Files(1))
    
End Sub

Private Sub imgNero_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Call HandleDragDrop(Data.Files(1))
    
End Sub

Private Sub imgAlcohol120_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Call HandleDragDrop(Data.Files(1))

End Sub

Private Sub imgBrowse_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Call HandleDragDrop(Data.Files(1))

End Sub

Private Sub imgProcess_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Call HandleDragDrop(Data.Files(1))

End Sub

Private Sub lblFilename_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Call HandleDragDrop(Data.Files(1))

End Sub

Private Sub txtFilename_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Call HandleDragDrop(Data.Files(1))

End Sub

Private Sub txtFoldername_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Call HandleDragDrop(Data.Files(1))

End Sub

Private Sub lblCDlabel_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Call HandleDragDrop(Data.Files(1))

End Sub

Private Sub txtCDlabel_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Call HandleDragDrop(Data.Files(1))

End Sub

Private Sub lblBurnSpeed_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Call HandleDragDrop(Data.Files(1))

End Sub

Private Sub lblBurnSpeedX_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Call HandleDragDrop(Data.Files(1))

End Sub

Private Sub cboBurnSpeed_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Call HandleDragDrop(Data.Files(1))

End Sub

Private Sub chkMultisession_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Call HandleDragDrop(Data.Files(1))

End Sub

Private Sub lblDiscFormat_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Call HandleDragDrop(Data.Files(1))

End Sub

Private Sub cboDiscFormat_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Call HandleDragDrop(Data.Files(1))

End Sub

Private Sub lblImageFormat_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Call HandleDragDrop(Data.Files(1))

End Sub

Private Sub cboImageFormat_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Call HandleDragDrop(Data.Files(1))

End Sub

Private Sub cbMerge_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Call HandleDragDrop(Data.Files(1))

End Sub

'handle drag 'n drop
Public Sub HandleDragDrop(sFile As String)

    'bring program to the front
    pSetForegroundWindow Me.hwnd
    
    Call openFile(sFile)
    
End Sub

'handle second instance command line
Private Sub txtOtherInstance_Change()
    
    Dim sCmdLine2ndInst As String
    
    If Len(txtOtherInstance.text) <> 0 Then
    
        'bring program to the front
        pSetForegroundWindow Me.hwnd
        
        'second instance command line
        sCmdLine2ndInst = Trim$(Replace$(txtOtherInstance.text, """", ""))
        
        'did the second instance passed a command line?
        If Len(sCmdLine2ndInst) <> 0 Then
        
            Call openFile(sCmdLine2ndInst)
    
        End If
    
        'clear otherwise this even my not fire next time
        txtOtherInstance.text = vbNullString
        
    End If
    
End Sub

Public Sub Init()

    Dim myUniqueID As String
    Dim hPrevInst As Long
    Dim lPropValue As Long

    blnStartup = True
    
    'check and reuse instance
    myUniqueID = "10928347lsdflijsf07124" & App.title
    lPropValue = txtOtherInstance.hwnd
    hPrevInst = IsPrevInstance(Me.hwnd, myUniqueID, lPropValue, True)
    If hPrevInst <> 0 Then End

    If FileExists(AppPath$ & "tools\audio.raw") = False Then
        MsgBox "tools\audio.raw is missing", vbCritical, "Error"
        End
    ElseIf FileExists(AppPath$ & "tools\cdi2nero.exe") = False Then
        MsgBox "tools\cdi2nero.exe is missing", vbCritical, "Error"
        End
    ElseIf FileExists(AppPath$ & "tools\cdi4dc.exe") = False Then
        MsgBox "tools\cdi4dc.exe is missing", vbCritical, "Error"
        End
    ElseIf FileExists(AppPath$ & "tools\cdirip.exe") = False Then
        MsgBox "tools\cdirip.exe is missing", vbCritical, "Error"
        End
    ElseIf FileExists(AppPath$ & "tools\cdrecord.exe") = False Then
        MsgBox "tools\cdrecord.exe is missing", vbCritical, "Error"
        End
    ElseIf FileExists(AppPath$ & "tools\IP.TMPL") = False Then
        MsgBox "tools\IP.TMPL is missing", vbCritical, "Error"
        End
    ElseIf FileExists(AppPath$ & "tools\lbacalc.exe") = False Then
        MsgBox "tools\lbacalc.exe is missing", vbCritical, "Error"
        End
    ElseIf FileExists(AppPath$ & "tools\libmp3lame-0.dll") = False Then
        MsgBox "tools\libmp3lame-0.dll is missing", vbCritical, "Error"
        End
    ElseIf FileExists(AppPath$ & "tools\libogg-0.dll") = False Then
        MsgBox "tools\libogg-0.dll is missing", vbCritical, "Error"
        End
    ElseIf FileExists(AppPath$ & "tools\libvorbis-0.dll") = False Then
        MsgBox "tools\libvorbis-0.dll is missing", vbCritical, "Error"
        End
    ElseIf FileExists(AppPath$ & "tools\libvorbisenc-2.dll") = False Then
        MsgBox "tools\libvorbisenc-2.dll is missing", vbCritical, "Error"
        End
    ElseIf FileExists(AppPath$ & "tools\libvorbisfile-3.dll") = False Then
        MsgBox "tools\libvorbisfile-3.dll is missing", vbCritical, "Error"
        End
    ElseIf FileExists(AppPath$ & "tools\mds4dc.exe") = False Then
        MsgBox "tools\mds4dc.exe is missing", vbCritical, "Error"
        End
    ElseIf FileExists(AppPath$ & "tools\mkisofs.exe") = False Then
        MsgBox "tools\mkisofs.exe is missing", vbCritical, "Error"
        End
    ElseIf FileExists(AppPath$ & "tools\newfile.exe") = False Then
        MsgBox "tools\newfile.exe is missing", vbCritical, "Error"
        End
    ElseIf FileExists(AppPath$ & "tools\scramble.exe") = False Then
        MsgBox "tools\scramble.exe is missing", vbCritical, "Error"
        End
    ElseIf FileExists(AppPath$ & "tools\sh-elf-objcopy.exe") = False Then
        MsgBox "tools\sh-elf-objcopy.exe is missing", vbCritical, "Error"
        End
    ElseIf FileExists(AppPath$ & "tools\sox.exe") = False Then
        MsgBox "tools\sox.exe is missing", vbCritical, "Error"
        End
    End If
    
    Me.Caption = Me.Caption & " v" & strCurVersion
    
    '//
    '// load settings from INI
    '//
    If ReadINI("Main", "Burn speed") <> "" Then cboBurnSpeed.text = ReadINI("Main", "Burn speed") Else cboBurnSpeed.text = "8"
    
    'ISO burning format in ImgRecord
    Select Case ReadINI("Menu", "ISO")
        Case "-data": Call mnuExtrasISOSettingsImgRecordMode1_Click
        Case "-xa": Call mnuExtrasISOSettingsImgRecordMode2Form1_Click
    End Select
    
    If ReadINI("Menu", "RockRidge") <> "" Then mnuExtrasISOSettingsRockRidge.Checked = ReadINI("Menu", "RockRidge") Else mnuExtrasISOSettingsRockRidge.Checked = True
    If ReadINI("Menu", "Joliet") <> "" Then mnuExtrasISOSettingsJoliet.Checked = ReadINI("Menu", "Joliet") Else mnuExtrasISOSettingsJoliet.Checked = True
    If ReadINI("Menu", "Full filenames") <> "" Then mnuExtrasISOSettingsFullFilenames.Checked = ReadINI("Menu", "Full filenames") Else mnuExtrasISOSettingsFullFilenames.Checked = True
    If ReadINI("Menu", "Add CDDA tracks") <> "" Then mnuExtrasAddCDDATracks.Checked = ReadINI("Menu", "Add CDDA tracks") Else mnuExtrasAddCDDATracks.Checked = False
    
    'load previous dummy setting
    If ReadINI("Menu", "Dummy file") = "650" Then
        mnuExtrasDummyFile650MB_Click
    ElseIf ReadINI("Menu", "Dummy file") = "700" Then
        Call mnuExtrasDummyFile700MB_Click
    ElseIf ReadINI("Menu", "Dummy file") <> "" And ReadINI("Menu", "Dummy file") <> "0" And ReadINI("Menu", "Dummy file") <> "None" Then
        lngDummySize = CLng(ReadINI("Menu", "Dummy file"))
        Call mnuExtrasDummyFileCustom_Click
    Else
        Call mnuExtrasDummyFileNone_Click
    End If
    
    If ReadINI("Menu", "Insert MR logo") <> "" Then mnuExtrasInsertMRlogo.Checked = ReadINI("Menu", "Insert MR logo") Else mnuExtrasInsertMRlogo.Checked = False
    If ReadINI("Menu", "Bootsector only (IP.BIN)") <> "" Then mnuExtrasBootsectorOnly.Checked = ReadINI("Menu", "Bootsector only (IP.BIN)") Else mnuExtrasBootsectorOnly.Checked = False
    If ReadINI("Nero", "Image format") <> "" Then cboImageFormat.text = ReadINI("Nero", "Image format") Else cboImageFormat.text = "DAO"
    If ReadINI("CDRecord", "Merge previous session") <> "" Then cbMerge.value = ReadINI("CDRecord", "Merge previous session") Else cbMerge.value = 1
    If ReadINI("ImgRecord", "Multisession") <> "" Then chkMultisession.value = ReadINI("ImgRecord", "Multisession") Else chkMultisession.value = 0
    
    'load last task
    Select Case ReadINI("Main", "Task")
        Case "DiscJuggler": Call imgDiscJuggler_Click
        Case "Nero": Call imgNero_Click
        Case "Alcohol 120%": Call imgAlcohol120_Click
        Case "CDRecord": Call imgCDRecord_Click
        Case "ImgRecord": Call imgImgRecord_Click
        Case Else: Call imgDiscJuggler_Click
    End Select
    
    If ReadINI("Main", "Disc format") <> "" Then cboDiscFormat.text = ReadINI("Main", "Disc format") Else cboDiscFormat.text = "Audio\Data"
    
    'crashes when selecting CD label text if this isnot done
    Me.Show
    
    blnStartup = False

End Sub

Public Sub openFile(sFile As String)
    
    If CBool(PathIsDirectory(sFile)) = True Then
        
        If Frame.Caption = "ImgRecord" Then
            Call imgDiscJuggler_Click
        End If
        
        txtFoldername.text = sFile
        
        'set and highlight the CD label text
        txtCDlabel.text = JustTitle$(sFile)
        txtCDlabel.SetFocus
        txtCDlabel.SelStart = 0
        txtCDlabel.SelLength = Len(txtCDlabel.text)
        
    Else
    
        If Frame.Caption <> "ImgRecord" Then
            Call imgImgRecord_Click
        End If
    
        txtFilename.text = sFile
        
        Call imgImgRecord_Click
    
    End If

End Sub
