VERSION 5.00
Begin VB.Form frmSelectCDDATracks 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Select the CDDA tracks..."
   ClientHeight    =   3255
   ClientLeft      =   75
   ClientTop       =   300
   ClientWidth     =   6855
   ControlBox      =   0   'False
   Icon            =   "frmSelectCDDATracks.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3255
   ScaleWidth      =   6855
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   120
      Top             =   2760
   End
   Begin VB.FileListBox flbFiles 
      Height          =   2235
      Left            =   120
      MultiSelect     =   2  'Extended
      OLEDropMode     =   1  'Manual
      Pattern         =   "*.voc;*.aif;*.aiff;*.au;*.mp2;*.mp3;*.ogg;*.vorbis;*.raw;*.wav"
      TabIndex        =   7
      Top             =   360
      Width           =   2895
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   375
      Left            =   4080
      TabIndex        =   6
      Top             =   2760
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   5280
      TabIndex        =   5
      Top             =   2760
      Width           =   1095
   End
   Begin VB.CommandButton cmdMoveDown 
      Height          =   255
      Left            =   6480
      Picture         =   "frmSelectCDDATracks.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1560
      Width           =   255
   End
   Begin VB.CommandButton cmdMoveUp 
      Height          =   255
      Left            =   6480
      Picture         =   "frmSelectCDDATracks.frx":0062
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1200
      Width           =   255
   End
   Begin VB.CommandButton cmdMoveLeft 
      Height          =   255
      Left            =   3120
      Picture         =   "frmSelectCDDATracks.frx":00B7
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1200
      Width           =   255
   End
   Begin VB.CommandButton cmdMoveRight 
      Height          =   255
      Left            =   3120
      Picture         =   "frmSelectCDDATracks.frx":010D
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1560
      Width           =   255
   End
   Begin VB.ListBox lbFiles 
      Height          =   2205
      ItemData        =   "frmSelectCDDATracks.frx":0163
      Left            =   3480
      List            =   "frmSelectCDDATracks.frx":0165
      TabIndex        =   0
      Top             =   360
      Width           =   2895
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Tracks to burn:"
      Height          =   195
      Left            =   3480
      TabIndex        =   9
      Top             =   120
      Width           =   1080
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Tracks to choose:"
      Height          =   195
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   1290
   End
End
Attribute VB_Name = "frmSelectCDDATracks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Canceled As Boolean

Private Sub flbFiles_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    flbFiles.ToolTipText = flbFiles.List((Y \ (flbFiles.FontSize * 24.4)) + flbFiles.TopIndex)

End Sub

Private Sub lbFiles_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    lbFiles.ToolTipText = lbFiles.List((Y \ (lbFiles.FontSize * 24.4)) + lbFiles.TopIndex)

End Sub

Private Sub Form_Load()

    flbFiles.Path = AppPath$ & "cdda"
    cmdOK.Enabled = False

End Sub

Private Sub cmdOK_Click()

    Dim i As Integer
    Dim strNewTrack As String
    
    If FileExists(AppPath$ & "cdda\temp") = False Then MkDir AppPath$ & "cdda\temp"
    
    strCDDAFilenames = ""
    
    'loop the list
    For i = 0 To lbFiles.ListCount - 1
    
        'track name w\ path (no extension)
        strNewTrack = AppPath$ & "cdda\temp\track" & Format(i, "00")
        
        'convert audio file
        Select Case LCase$(GetExtension$(lbFiles.List(i)))
        
            Case "mp2", "mp3", "ogg", "vorbis"
                Call ShellWait("""" & AppPath$ & "tools\sox.exe"" -S """ & AppPath$ & "cdda\" & lbFiles.List(i) & """ """ & strNewTrack & ".wav""", vbNormalFocus)
                Call ShellWait("""" & AppPath$ & "tools\sox.exe"" -S """ & strNewTrack & ".wav"" """ & strNewTrack & ".raw""", vbNormalFocus)
                Call Kill(strNewTrack & ".wav")
                
            Case "aif", "aiff", "au", "wav", "snd", "voc"
                Call ShellWait("""" & AppPath$ & "tools\sox.exe"" -S """ & AppPath$ & "cdda\" & lbFiles.List(i) & """ """ & strNewTrack & ".raw""", vbNormalFocus)
            
            Case "raw"
                Call FileCopy(AppPath$ & "cdda\audio.raw", strNewTrack & ".raw")
        
        End Select
        
        'add new raw track
        strCDDAFilenames = strCDDAFilenames & """" & strNewTrack & ".raw"" "
        
    Next
    
    Canceled = False
    Unload Me

End Sub

'** CANCEL BUTTON **
Private Sub cmdCancel_Click()

    Canceled = True
    Unload Me

End Sub

Private Sub cmdMoveUp_Click()

    Dim tempEntry As String
    Dim tempIndex As Byte
    
    'are there files in the listbox?
    If lbFiles.ListIndex <> -1 Then
    
        'yes, get the current index and text
        tempEntry = lbFiles.List(lbFiles.ListIndex)
        tempIndex = lbFiles.ListIndex
        
        'make sure we're not at the top already
        If tempIndex = 0 Then Exit Sub
        
        'remove current item
        lbFiles.RemoveItem lbFiles.ListIndex
        
        'add temp item up one entry
        lbFiles.AddItem tempEntry, tempIndex - 1
        
        'highlight moved entry
        lbFiles.ListIndex = tempIndex - 1
        
    End If

End Sub

Private Sub cmdMoveDown_Click()

    Dim tempEntry As String
    Dim tempIndex As Byte
    
    'are there files in the listbox?
    If lbFiles.ListIndex <> -1 Then
    
        'yes, get the current index and text
        tempEntry = lbFiles.List(lbFiles.ListIndex)
        tempIndex = lbFiles.ListIndex
        
        'make sure we're not at the bottom already
        If tempIndex = lbFiles.ListCount - 1 Then Exit Sub
        
        'remove current item
        lbFiles.RemoveItem lbFiles.ListIndex
        
        'add temp item down one entry
        lbFiles.AddItem tempEntry, tempIndex + 1
        
        'highlight moved entry
        lbFiles.ListIndex = tempIndex + 1
        
    End If

End Sub

Private Sub cmdMoveLeft_Click()

    Dim i As Integer
    
    'does the listbox have any files in it?
    If lbFiles.ListCount <> 0 Then
    
        'yes, so loop them backwards to prevent crashing if we remove more than one file
        For i = (lbFiles.ListCount - 1) To 0 Step -1
        
            'is the current file selected?
            If lbFiles.Selected(i) = True Then
            
                'yes, so remove it
                lbFiles.RemoveItem i
                
                If i > lbFiles.ListCount - 1 Then
                    'lbFiles.Selected(i - 1) = True 'todo: fix
                Else
                    lbFiles.Selected(i) = True
                End If
                
            End If
            
        Next
        
    End If
    
    'is the listbox empty?
    If lbFiles.ListCount = 0 Then
        cmdOK.Enabled = False
    End If

End Sub

Private Sub cmdMoveRight_Click()

    Dim i As Integer
    
    'does the file list have any files in it?
    If flbFiles.ListCount <> 0 Then
    
        'yes, so loop them
        For i = 0 To flbFiles.ListCount - 1
        
            'is the current file selected?
            If flbFiles.Selected(i) = True Then
                
                'yes, so add it to the listbox
                lbFiles.AddItem (flbFiles.List(i))
                
            End If
            
        Next
        
    End If
    
    cmdOK.Enabled = True

End Sub

Private Sub lbFiles_DblClick()

    Call cmdMoveLeft_Click

End Sub

Private Sub flbFiles_DblClick()

    Call cmdMoveRight_Click

End Sub

Private Sub flbFiles_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Dim i As Long
    
    'loop the dropped files
    For i = 1 To Data.Files.Count
        Call HandleDragDrop(Data.Files(i))
    Next
    
End Sub

'moves known dropped files to \cdda
Public Sub HandleDragDrop(sFile As String)

    'bring program to the front
    pSetForegroundWindow Me.hwnd
    
    'does the dragged item exist?
    If FileExists(sFile) = True Then
    
        'yes, a supported extension there?
        Select Case GetExtension$(sFile)
            Case "voc", "aif", "aiff", "au", "mp2", "mp3", "vorbis", "ogg", "raw", "wav"
            
                'yes, does the file already exist?
                If FileExists(AppPath$ & "cdda\" & JustTitle$(sFile)) = True Then
                    
                    'yes, replace it?
                    If MsgBox(AppPath$ & "cdda\" & JustTitle$(sFile) & " already exists." & vbCrLf & "Do you want to replace it?", vbExclamation + vbYesNo, "Replace") = vbYes Then
                        
                        'yes, delete the current file
                        Call SetAttr(AppPath$ & "cdda\" & JustTitle$(sFile), vbNormal)
                        Call Kill(AppPath$ & "cdda\" & JustTitle$(sFile))
                        
                    Else
                    
                        'no
                        Exit Sub
                        
                    End If
                    
                End If
                
                'move to cdda
                Name sFile As AppPath$ & "cdda\" & JustTitle$(sFile)
                
                'show the new file
                flbFiles.Refresh
            
        End Select
    
    End If
    
End Sub

'refreshs list for new found files
Private Sub Timer1_Timer()

    Dim strFileName As String
    Dim lngFilesInFolder As Long
    Dim i As Long
    
    'initialize files in folder count
    lngFilesInFolder = 0
    
    'get the first file in the folder
    strFileName = Dir$(AppPath$ & "cdda\", vbNormal + vbReadOnly + vbArchive)
    
    'loop the directory
    Do Until strFileName = ""
    
        'add to file in folder count if an audio file
        Select Case GetExtension$(strFileName)
            Case "voc", "aif", "aiff", "au", "mp2", "mp3", "ogg", "vorbis", "raw", "wav"
                lngFilesInFolder = lngFilesInFolder + 1
        End Select
        
        'get the next file
        strFileName = Dir
        
    Loop
    
    'is the current file in folder count different then the number of files listed _
    in the file list box?
    If flbFiles.ListCount <> lngFilesInFolder Then
    
        'yes, so refresh the list
        flbFiles.Refresh
        Exit Sub
        
    Else
    
        'no, so were going to check each file in the list with the files in the folder
        'get the first file in the folder
        strFileName = Dir$(AppPath$ & "cdda\", vbNormal + vbReadOnly + vbArchive)
        
        'loop the directory
        Do Until strFileName = ""
        
            'check next file if not an audio file
            Select Case GetExtension$(strFileName)
                Case "voc", "aif", "aiff", "au", "mp2", "mp3", "ogg", "vorbis", "raw", "wav"
                Case Else
                    GoTo NextFile
            End Select
        
            'loop the file list
            For i = 1 To flbFiles.ListCount
            
                If strFileName = flbFiles.List(i - 1) Then
                    'current file is in the list
                    GoTo NextFile
                End If
            
            Next
            
            'should get here only if GoTo was not executed
            flbFiles.Refresh
            Exit Sub
            
NextFile:
            'get the next file
            strFileName = Dir
            
        Loop
        
    End If
    
End Sub

