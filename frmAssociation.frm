VERSION 5.00
Begin VB.Form frmAssociation 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Association"
   ClientHeight    =   2310
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4590
   ControlBox      =   0   'False
   Icon            =   "frmAssociation.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2310
   ScaleWidth      =   4590
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdDone 
      Caption         =   "Done"
      Default         =   -1  'True
      Height          =   375
      Left            =   3360
      TabIndex        =   4
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Height          =   1575
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4335
      Begin VB.CheckBox chkCUE 
         Caption         =   "CUE sheets"
         Height          =   195
         Left            =   240
         TabIndex        =   3
         Top             =   1080
         Width           =   1215
      End
      Begin VB.CheckBox chkISO 
         Caption         =   "ISO images"
         Height          =   195
         Left            =   240
         TabIndex        =   2
         Top             =   720
         Width           =   1215
      End
      Begin VB.CheckBox chkDiscJuggler 
         Caption         =   "DiscJuggler images"
         Height          =   195
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   1695
      End
   End
End
Attribute VB_Name = "frmAssociation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()

    ReadAssociation
    
End Sub

Private Sub chkCUE_Click()

    If chkCUE.value = 0 Then Call DisAssociateExtension(".cue", "CUE sheet")
    If chkCUE.value = 1 Then Call AssociateExtension(".cue", "CUE sheet")

End Sub

Private Sub chkDiscJuggler_Click()

    If chkDiscJuggler.value = 0 Then Call DisAssociateExtension(".cdi", "DiscJuggler CD image")
    If chkDiscJuggler.value = 1 Then Call AssociateExtension(".cdi", "DiscJuggler CD image")

End Sub

Private Sub chkISO_Click()

    If chkISO.value = 0 Then Call DisAssociateExtension(".iso", "ISO image")
    If chkISO.value = 1 Then Call AssociateExtension(".iso", "ISO image")

End Sub

Private Sub cmdDone_Click()

    Unload Me

End Sub
