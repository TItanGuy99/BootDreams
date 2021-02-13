Attribute VB_Name = "modPubDeclares"
Option Explicit

Public Const strCurVersion As String = "1.0.6c"
Public Const strOnlineVersion As String = "http://dchelp.dcemulation.org/downloads/bootdreams.txt"

Public blnStartup               As Boolean
Public strDrvID                 As String   'ex: 0,2,0
Public strCDDAFilenames         As String   'full path with quotes
Public strMainBinaryFilename    As String   'full path
Public strMRFilename            As String   'full path
Public lngDummySize             As Long     'dummy size in MB
Public blnIP                    As Boolean  'IP.BIN routine completed successfully
Public blnASPI                  As Boolean  'cdrecord found ASPI
Public blnDrives                As Boolean  'cdrecord returned CD drives
Public blnISOSuccess            As Boolean  'ISO header conversion
Public blnConvSuccess           As Boolean  'Main binary conversion
