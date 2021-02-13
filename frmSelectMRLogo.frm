VERSION 5.00
Begin VB.Form frmSelectMRLogo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Select a logo..."
   ClientHeight    =   4935
   ClientLeft      =   1425
   ClientTop       =   3570
   ClientWidth     =   5295
   ControlBox      =   0   'False
   Icon            =   "frmSelectMRLogo.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4935
   ScaleWidth      =   5295
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   120
      Top             =   4440
   End
   Begin VB.Frame Frame1 
      Height          =   1695
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   5055
      Begin VB.PictureBox picMR 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   1335
         Left            =   120
         ScaleHeight     =   89
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   320
         TabIndex        =   4
         Top             =   240
         Width           =   4800
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   375
      Left            =   2880
      TabIndex        =   1
      Top             =   4440
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4080
      TabIndex        =   2
      Top             =   4440
      Width           =   1095
   End
   Begin VB.FileListBox flbFiles 
      Height          =   2235
      Left            =   120
      OLEDropMode     =   1  'Manual
      Pattern         =   "*.mr"
      TabIndex        =   0
      Top             =   2040
      Width           =   5055
   End
End
Attribute VB_Name = "frmSelectMRLogo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Canceled As Boolean

Private Sub flbFiles_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Dim i As Long
    
    'loop the dropped files
    For i = 1 To Data.Files.Count
        Call HandleDragDrop(Data.Files(i))
    Next
    
End Sub

Private Sub Form_Load()

    flbFiles.Path = AppPath$ & "iplogos"
    
End Sub

'** OK BUTTON **
Private Sub cmdOK_Click()
    
    'will the MR fit in a normal IP.BIN?
    If FileLen(AppPath$ & "iplogos\" & flbFiles.Filename) > 8192 Then
    
        'no, but continue?
        If MsgBox("This MR will not fit in a normal IP.BIN. Do you want to continue?", vbYesNo + vbExclamation + vbDefaultButton2, "Warning") = vbNo Then
        
            'no
            Exit Sub
            
        End If
    
    End If
    
    strMRFilename = AppPath$ & "iplogos\" & flbFiles.Filename
    
    Canceled = False
    Unload Me
    
End Sub

'** FL CLK **
Private Sub flbFiles_Click()

    'enable/diable the OK button
    If flbFiles.Filename <> "" Then
    
        cmdOK.Enabled = True
        
        Call ViewMRLogo(AppPath$ & "iplogos\" & flbFiles.Filename)
        
    Else
    
        cmdOK.Enabled = False
        
    End If

End Sub

'** FL DBL CLK **
Private Sub flbFiles_DblClick()

    Call cmdOK_Click

End Sub

'** CANCEL BUTTON **
Private Sub cmdCancel_Click()

    Canceled = True
    Unload Me
    
End Sub

Public Sub HandleDragDrop(sFile As String)

    'bring program to the front
    pSetForegroundWindow Me.hwnd
    
    'does the dragged item exist?
    If FileExists(sFile) = True Then
    
        'yes, is the .mr extension there?
        If GetExtension$(sFile) = "mr" Then
        
            'yes, does the .mr already exist?
            If FileExists(AppPath$ & "iplogos\" & JustTitle$(sFile)) = True Then
                
                'yes, replace it?
                If MsgBox(AppPath$ & "iplogos\" & JustTitle$(sFile) & " already exists." & vbCrLf & "Do you want to replace it?", vbExclamation + vbYesNo, "Replace") = vbYes Then
                    
                    'yes, delete the current mr
                    Call SetAttr(AppPath$ & "iplogos\" & JustTitle$(sFile), vbNormal)
                    Call Kill(AppPath$ & "iplogos\" & JustTitle$(sFile))
                    
                Else
                
                    'no, nothing to delete
                    Exit Sub
                    
                End If
                
            End If
            
            'move the mr to iplogos
            Name sFile As AppPath$ & "iplogos\" & JustTitle$(sFile)
            
            'show the new file
            flbFiles.Refresh
            
            'default the window
            Set picMR.Picture = Nothing
            cmdOK.Enabled = False
            
            'loop the file list
            'For i = 0 To flbFiles.ListCount
            
                'is current item our dragged item?
                'If flbFiles.List(i) = JustTitle$(sFile) Then
                
                    'yes, so select it
                    'flbFiles.Selected(i) = True
                    'Exit For
                    
                'End If
                
            'Next
            
        End If
    
    End If
    
End Sub

Private Sub Timer1_Timer()

    Dim strFileName As String
    Dim lngFilesInFolder As Long
    Dim i As Long
    
    'initialize files in folder count
    lngFilesInFolder = 0
    
    'get the first file in the folder
    strFileName = Dir$(AppPath$ & "iplogos\", vbNormal + vbReadOnly + vbArchive)
    
    'loop the directory
    Do Until strFileName = ""
    
        'add to file in folder count
        If GetExtension$(strFileName) = "mr" Then
            lngFilesInFolder = lngFilesInFolder + 1
        End If
        
        'get the next file
        strFileName = Dir
        
    Loop
    
    'is the folder count different then the number of files listed in the file list?
    If flbFiles.ListCount <> lngFilesInFolder Then
    
        'yes, so refresh the list
        flbFiles.Refresh
        Exit Sub
        
    Else
    
        'no, so were going to check each file in the list with the files in the folder.
        'get the first file in the folder
        strFileName = Dir$(AppPath$ & "iplogos\", vbNormal + vbReadOnly + vbArchive)
        
        'loop the directory
        Do Until strFileName = ""
        
            'loop the file list
            For i = 1 To flbFiles.ListCount
            
                'not of MR extension
                If GetExtension$(strFileName) <> "mr" Then
                    GoTo GetOut
                End If
            
                'current file is in the list
                If strFileName = flbFiles.List(i - 1) Then
                    GoTo GetOut
                End If
            
            Next
            
            'should get here only if GoTo was not executed
            flbFiles.Refresh
            Exit Sub
            
GetOut:
            'get the next file
            strFileName = Dir
            
        Loop
        
    End If
    
End Sub
