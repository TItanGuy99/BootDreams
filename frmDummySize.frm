VERSION 5.00
Begin VB.Form frmDummySize 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Dummy Size"
   ClientHeight    =   1230
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   2790
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1230
   ScaleWidth      =   2790
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   1440
      TabIndex        =   2
      Top             =   720
      Width           =   1215
   End
   Begin VB.TextBox txtDummySize 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   2535
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Size (in MB):"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   885
   End
End
Attribute VB_Name = "frmDummySize"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Canceled As Boolean

'** CANCEL BUTTON **
Private Sub cmdCancel_Click()

    Canceled = True
    Unload Me
    
End Sub

Private Sub cmdOK_Click()

    If txtDummySize < 3 Then
        MsgBox "Please set the dummy size to 3 MB or bigger.", vbCritical, "Error"
        Exit Sub
    End If

    lngDummySize = txtDummySize.text
    Canceled = False
    Unload Me

End Sub

Private Sub Form_Load()

    txtDummySize.text = lngDummySize

End Sub

Private Sub txtDummySize_KeyPress(KeyAscii As Integer)

    If KeyAscii > 31 And (KeyAscii < 48 Or KeyAscii > 57) Then
        KeyAscii = 0
        Beep
    End If
    
End Sub
