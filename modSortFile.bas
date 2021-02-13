Attribute VB_Name = "modSortFile"
Option Explicit

'insertion sort method adapted from xtremevbtalk

Private FileList As String
Private RootFolder As String
Private RootFolder1 As String
Private GetRoot As Integer

Public Sub CreateSortFile(sPath As String)
    
'On Error GoTo ErrorHandler

    Dim SortFile() As String
    Dim i As Long, j As Long
    Dim FilesizeEnds As Integer, Filesize1 As Long, Filesize2 As Long
    Dim Weight As Integer
    Dim sTemp As String
    
    'get files
    GetFiles (sPath)
    FileList = Left(FileList, Len(FileList) - 2)
    SortFile = Split(FileList, vbCrLf)
    
    'highest weight size
    Weight = UBound(SortFile) + 1
    
    For i = LBound(SortFile) + 1 To UBound(SortFile)
        'Get the value to be inserted
        FilesizeEnds = InStr(1, SortFile(i), "|") - 1
        Filesize1 = Left(SortFile(i), FilesizeEnds)
        sTemp = SortFile(i)
        'Move along the already sorted values shifting along
        For j = i - 1 To LBound(SortFile) Step -1
            FilesizeEnds = InStr(1, SortFile(j), "|") - 1
            Filesize2 = Left(SortFile(j), FilesizeEnds)
            'No more shifting needed, we found the right spot!
            If Filesize2 <= Filesize1 Then Exit For
            SortFile(j + 1) = SortFile(j)
        Next j
        'Insert value in the slot
        SortFile(j + 1) = sTemp
    Next i
    
    'change to unix folder separaters, remove filesizes and add weight
    Open AppPath$ & "sort.txt" For Output As #1
        For i = LBound(SortFile) To UBound(SortFile)
            FilesizeEnds = InStr(1, SortFile(i), "|") + 1
            SortFile(i) = Mid(SortFile(i), FilesizeEnds)
            SortFile(i) = Replace(SortFile(i), "\", "/")
            SortFile(i) = SortFile(i) & " " & Weight
            Weight = Weight - 1
            Print #1, SortFile(i)
        Next i
    Close #1

    Exit Sub

ErrorHandler:
    MsgBox "CreateSortFile - modSortFile" & vbCrLf & vbCrLf & Err & " - " & Err.Description, vbCritical, "Error"
        
End Sub
        
Public Sub GetFiles(sPath As String)

On Error GoTo ErrorHandler

    Dim sCurPath, sCurPath2 As String, sName As String
    Dim sLastDir As String, RootFolderStarts As Integer
    Dim RootFolderStarts1 As Integer

    sCurPath = sPath & "\"
    sCurPath2 = sPath & "\"
    sName = vbNullString
    sName = Dir(sCurPath, vbDirectory)
    
    If GetRoot <> 1 Then
        GetRoot = 0
    End If
        
    Do While sName <> vbNullString
        If (GetAttr(sCurPath & sName) And vbDirectory) = vbDirectory Then
            If sName <> "." And sName <> ".." Then
                sLastDir = sName
                GetFiles sCurPath & sName
                sName = Dir(sCurPath, vbDirectory)
                While sName <> sLastDir
                    sName = Dir
                Wend
            End If
        Else
            If GetRoot = 0 Then
                RootFolderStarts = InStrRev(sPath, "\") + 1
                RootFolder = Mid(sPath, RootFolderStarts) & "\"
                GetRoot = 1
            End If
            If sName = "000DUMMY.DAT" Then
                FileList = FileList & "0" & "|" & RootFolder & sName & vbCrLf
            Else
                RootFolderStarts = InStrRev(sPath, "\") + 1
                RootFolder1 = Mid(sPath, RootFolderStarts) & "\"
                If RootFolder1 = RootFolder Then
                    FileList = FileList & FileLen(sCurPath & sName) & "|" & RootFolder & sName & vbCrLf
                Else
                    FileList = FileList & FileLen(sCurPath & sName) & "|" & RootFolder & RootFolder1 & sName & vbCrLf
                End If
            End If
        End If
        sName = Dir
    Loop

    Exit Sub

ErrorHandler:
    MsgBox "GetFiles - modSortFile" & vbCrLf & vbCrLf & Err & " - " & Err.Description, vbCritical, "Error"
    
End Sub

