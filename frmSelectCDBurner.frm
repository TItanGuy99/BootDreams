VERSION 5.00
Begin VB.Form frmSelectCDBurner 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Select a CD burner..."
   ClientHeight    =   2565
   ClientLeft      =   4950
   ClientTop       =   5610
   ClientWidth     =   4245
   ControlBox      =   0   'False
   Icon            =   "frmSelectCDBurner.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2565
   ScaleWidth      =   4245
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3000
      TabIndex        =   2
      Top             =   2040
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   375
      Left            =   1800
      TabIndex        =   1
      Top             =   2040
      Width           =   1095
   End
   Begin VB.ListBox lsDrives 
      Height          =   1620
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3975
   End
End
Attribute VB_Name = "frmSelectCDBurner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Canceled As Boolean

Private Sub Form_Load()
    
On Error GoTo ErrorHandler

    Dim strScanbus As String
    Dim strScanbusAr() As String
    Dim strTempDrive As String
    Dim strDriveAr() As String
    Dim lngTempDrive As Long
    Dim strDrive As String
    Dim i As Integer
    
    strScanbus = ExecuteApp("""" & AppPath$ & "tools\cdrecord.exe"" -scanbus")
    
    'does CDRecord need a ASPI driver?
    If InStr(1, strScanbus, "Can not load ASPI driver!") > 0 Then
    
        'yes
        MsgBox "CDRecord: Can not load ASPI driver!", vbCritical, "Error"
        blnASPI = False
        Exit Sub
        
    Else
    
        'no
        blnASPI = True
        
    End If
    
    strScanbusAr = Split(strScanbus, vbLf)
    
    'line example: _
            0,2,0     2) '_NEC    ' 'DVD_RW ND-3540A ' '1.WB' Removable CD-ROM
        
    'loop each line of the -scanbus output
    For i = LBound(strScanbusAr) To UBound(strScanbusAr)
        
        'does the current line have "CD-ROM" in it?
        If InStr(1, strScanbusAr(i), "CD-ROM") > 0 Then
            
            'yes, so parse the drive name & id
            'remove first 2 tabs
            strTempDrive = Mid$(strScanbusAr(i), 2)
            
            'split the drive ID and drive name into an array
            strDriveAr = Split(strTempDrive, vbTab)
            
            'parse the drive brand/model/firmware
            lngTempDrive = InStr(1, strDriveAr(1), "'")
            strDrive = Mid$(strDriveAr(1), lngTempDrive, 37)
            strDrive = Replace$(strDrive, "'", "")
            
            'add the drive to the list
            lsDrives.AddItem strDrive & Space$(100) & strDriveAr(0)
            
        End If
        
    Next
    
    'did CDRecord return any CD-ROM drives?
    If lsDrives.ListCount = 0 Then
    
        'no
        MsgBox "CDRecord did not return any CD-ROM drives.", vbCritical, "Error"
        blnDrives = False
        Exit Sub
        
    Else
    
        'yes
        blnDrives = True
        
    End If
    
    Exit Sub

ErrorHandler:
    MsgBox "Form_Load - frmSelectCDBurner" & vbCrLf & vbCrLf & Err & " - " & Err.Description, vbCritical, "Error"

End Sub

'** OK BUTTON **
Private Sub cmdOK_Click()

    strDrvID = Trim$(Mid$(lsDrives.text, InStrRev(lsDrives.text, " ")))
    
    Canceled = False
    Unload Me

End Sub

'** LST CLK **
Private Sub lsDrives_Click()

    'enable/disable the OK button
    If lsDrives.text <> "" Then
        cmdOK.Enabled = True
    Else
        cmdOK.Enabled = False
    End If

End Sub

'** LST DBL CLK **
Private Sub lsDrives_DblClick()

    Call cmdOK_Click
        
End Sub

'** CANCEL BUTTON **
Private Sub cmdCancel_Click()
    
    Canceled = True
    Unload Me

End Sub
