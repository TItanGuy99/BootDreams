VERSION 5.00
Begin VB.Form frmSelectMainBinary 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Select the Main binary..."
   ClientHeight    =   3030
   ClientLeft      =   5265
   ClientTop       =   8880
   ClientWidth     =   3135
   ControlBox      =   0   'False
   Icon            =   "frmSelectMainBinary.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3030
   ScaleWidth      =   3135
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   120
      Top             =   2520
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   375
      Left            =   720
      TabIndex        =   1
      Top             =   2520
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1920
      TabIndex        =   2
      Top             =   2520
      Width           =   1095
   End
   Begin VB.FileListBox flbFiles 
      Height          =   2235
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2895
   End
End
Attribute VB_Name = "frmSelectMainBinary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Canceled As Boolean

Private Sub flbFiles_Click()

    'enables/disables the OK button
    If flbFiles.Filename <> "" And LCase$(flbFiles.Filename) <> "ip.bin" Then
        cmdOK.Enabled = True
    Else
        cmdOK.Enabled = False
    End If

End Sub

Private Sub Form_Load()

    '''this form is loaded twice
    'the first time the form is loaded it's loaded into memory without getting shown _
    at an attempt to autodetect the main binary if there's one or two files in the list
    'the second time the form is loaded it's shown and awaits the user to select the _
    main binary

    flbFiles.Path = frmMain.txtFoldername.text
    
    Canceled = False

    If flbFiles.ListCount = 0 Then
        MsgBox "The current folder has no files in it.", vbCritical, "Error"
        Canceled = True
        Exit Sub
    End If
    
    '//
    '// One file in folder
    '//
    If flbFiles.ListCount = 1 Then
        
        'is the file a IP.BIN?
        If LCase$(flbFiles.List(0)) <> "ip.bin" Then
        
            'no, is it longer then 16 characters?
            If Len(flbFiles.List(0)) > 16 Then
            
                'yes
                MsgBox "The detected main binary is " & Len(flbFiles.List(0)) - 16 & " characters too long.", vbCritical, "Error"
                Canceled = True
                
            Else
            
                'no, so lets set it as our main binary
                strMainBinaryFilename = AddTrailingSlash$(frmMain.txtFoldername.text) & flbFiles.List(0)
                
            End If
            Exit Sub
            
        Else
        
            'yes, it's a IP.BIN
            MsgBox "The current folder only has a IP.BIN in it.", vbCritical, "Error"
            Canceled = True
            Exit Sub
            
        End If
        
    End If
    
    '//
    '// Two files in folder
    '//
    If flbFiles.ListCount = 2 Then
    
        'check for the IP.BIN in the fl, if it's found
        'then the other file is our main binary
        
        If LCase$(flbFiles.List(0)) = "ip.bin" Then
            
            'is the second file longer than 16 characters?
            If Len(flbFiles.List(1)) > 16 Then
            
                'yes
                MsgBox "The detected main binary is " & Len(flbFiles.List(1)) - 16 & " characters too long.", vbCritical, "Error"
                Canceled = True
                
            Else
            
                'no, so set it as our main binary
                strMainBinaryFilename = AddTrailingSlash$(frmMain.txtFoldername.text) & flbFiles.List(1)
                
            End If
            Exit Sub
            
        ElseIf LCase$(flbFiles.List(1)) = "ip.bin" Then
        
            'is the first file longer than 16 characters?
            If Len(flbFiles.List(0)) > 16 Then
            
                'yes
                MsgBox "The detected main binary is " & Len(flbFiles.List(0)) - 16 & " characters too long.", vbCritical, "Error"
                Canceled = True
                
            Else
            
                'no, so set it as our main binary
                strMainBinaryFilename = AddTrailingSlash$(frmMain.txtFoldername.text) & flbFiles.List(0)
                
            End If
            Exit Sub
            
        Else
        
            'the main binary isn't found yet
            Exit Sub
            
        End If
        
    End If

End Sub

'** OK BUTTON **
Private Sub cmdOK_Click()
    
    'is the file longer than 16 characters?
    If Len(flbFiles.Filename) > 16 Then
        
        'yes
        MsgBox "This file is " & Len(flbFiles.Filename) - 16 & " characters too long.", vbCritical, "Error"
        Exit Sub
        
    Else
    
        'no, so set it as our main binary
        strMainBinaryFilename = AddTrailingSlash$(frmMain.txtFoldername.text) & flbFiles.Filename
        Canceled = False
        Unload Me
    
    End If

End Sub

'** FL DBL CLICK **
Private Sub flbFiles_DblClick()

    'is the OK button disabled? (IP.BIN or no file selected)
    If cmdOK.Enabled <> False Then
    
        'no, so call OK_Click
        Call cmdOK_Click
        
    End If

End Sub

'** CANCEL BUTTON **
Private Sub cmdCancel_Click()

    Canceled = True
    Unload Me

End Sub

Private Sub Timer1_Timer()

    Dim strFileName As String
    Dim lngFilesInFolder As Long
    Dim i As Long
    
    'initialize files in folder count
    lngFilesInFolder = 0
    
    'get the first file in the folder
    strFileName = Dir$(AddTrailingSlash$(frmMain.txtFoldername.text), vbNormal + vbReadOnly + vbArchive)
    
    'loop the directory
    Do Until strFileName = ""
    
        'add to file in folder count
        lngFilesInFolder = lngFilesInFolder + 1
        
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
        strFileName = Dir$(AddTrailingSlash$(frmMain.txtFoldername.text), vbNormal + vbReadOnly + vbArchive)
        
        'loop the directory
        Do Until strFileName = ""
        
            'loop the file list
            For i = 1 To flbFiles.ListCount
            
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
