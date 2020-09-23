VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   8640
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10230
   LinkTopic       =   "Form1"
   ScaleHeight     =   8640
   ScaleWidth      =   10230
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdRenameFiles 
      Caption         =   "Rename files"
      Height          =   255
      Left            =   8760
      TabIndex        =   8
      Top             =   8280
      Width           =   1215
   End
   Begin VB.ListBox lstRenamedFile 
      Height          =   1425
      Left            =   120
      TabIndex        =   7
      Top             =   6720
      Width           =   9975
   End
   Begin VB.ListBox lstOriginalName 
      Height          =   1425
      Left            =   120
      TabIndex        =   6
      Top             =   5160
      Width           =   9975
   End
   Begin VB.TextBox TxtPath 
      Height          =   285
      Left            =   1320
      TabIndex        =   4
      Text            =   "C:\Documents and Settings\Drew\Desktop\TestZip\"
      Top             =   120
      Width           =   7455
   End
   Begin VB.Timer tmrUnZip 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   9360
      Top             =   1560
   End
   Begin VB.CommandButton cmdGetZipFiles 
      Caption         =   "Get Zip Files"
      Height          =   255
      Left            =   8880
      TabIndex        =   3
      Top             =   120
      Width           =   1215
   End
   Begin VB.ListBox lstTextfiles 
      Height          =   1425
      Left            =   120
      TabIndex        =   2
      Top             =   3600
      Width           =   9975
   End
   Begin VB.ListBox lstZipsPath 
      Height          =   1425
      Left            =   120
      TabIndex        =   1
      Top             =   2040
      Width           =   9975
   End
   Begin VB.ListBox lstZips 
      Height          =   1425
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   9975
   End
   Begin VB.Label lblPathOfZipFiles 
      Caption         =   "Path of Zip files"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private Const DDL_READWRITE      As Long = &H0
Private Const LB_DIR             As Long = &H18D
Private FSO                      As New FileSystemObject
Private CounterZipFile           As Integer
Private CounterFullPath          As Integer
Private CounterRename            As Integer
Private nextline                 As String
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, _
ByVal wMsg As Long, _
ByVal wParam As Integer, _
ByVal lParam As Any) As Long
Private Sub cmdGetZipFiles_Click()

    If CounterZipFile > lstZips.ListCount Then
        MsgBox "Finished"
        'Temp fix for the extra file added
    lstOriginalName.RemoveItem (lstOriginalName.ListCount - 1)
    lstRenamedFile.RemoveItem (lstRenamedFile.ListCount - 1)
        'End Fix
        Exit Sub
    End If
    
    If CounterZipFile = 0 Then
        Call Getfiles(lstZips, TxtPath.Text, "*.zip")
        TxtPath.Enabled = False
        Do Until CounterFullPath = lstZips.ListCount
            lstZipsPath.AddItem TxtPath.Text & lstZips.List(CounterFullPath)
            CounterFullPath = CounterFullPath + 1
        Loop
    End If
    
    Call WinZipit(lstZipsPath.List(CounterZipFile), App.Path & "\;", UNZIP)
    If CounterZipFile > lstZipsPath.ListCount Then
        Exit Sub
     Else
        tmrUnZip.Enabled = True
    End If
    
End Sub

Private Sub cmdRemove_Click()

End Sub

Private Sub cmdRenameFiles_Click()
    Do Until CounterRename = lstRenamedFile.ListCount
        Name lstZipsPath.List(CounterRename) As lstRenamedFile.List(CounterRename)
        CounterRename = CounterRename + 1
    Loop
End Sub
Private Sub Form_Load()
    CounterZipFile = 0
    CounterFullPath = 0
    CounterRename = 0
End Sub
'Requires: Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Any) As Long
Public Sub Getfiles(Lst As ListBox, _
                    ByVal Path As String, _
                    ByVal Filetype As String)
  
  Dim sPattern As String
    
    sPattern = Path & "\" & Filetype
    SendMessage Lst.hwnd, LB_DIR, DDL_READWRITE, sPattern
End Sub


Private Sub tmrUnZip_Timer()
    On Error Resume Next
    tmrUnZip.Enabled = False
    Call Getfiles(lstTextfiles, App.Path & "\;", "*.txt")
    '    If (List3.ListCount - 1) > CounterZipFile Then
    '    Do Until List3.ListCount = CounterZipFile
    '    List3.RemoveItem (0)
    '    Loop
    '    GoTo Morethan1
    '    End If
    Open App.Path & "\;\" & lstTextfiles.List(CounterZipFile) For Input As #1
    Input #1, nextline
    Close #1
    lstOriginalName.AddItem Mid$(nextline, 8)
    Call StripAll
    lstRenamedFile.AddItem TxtPath.Text & lstOriginalName.List(CounterZipFile) & ".zip"
    'Name List2.List(CounterZipFile) As List5.List(CounterZipFile)
    'Morethan1:
    'Del the directory
    Call FSO.DeleteFolder(App.Path & "\;", True)
    'goto the next zip file
    CounterZipFile = CounterZipFile + 1
    'start the process again
    cmdGetZipFiles_Click
End Sub

Public Sub StripAll()
    'The left side
    lstOriginalName.List(CounterZipFile) = LStrip(lstOriginalName.List(CounterZipFile), "_")
    lstOriginalName.List(CounterZipFile) = LStrip(lstOriginalName.List(CounterZipFile), "*")
    lstOriginalName.List(CounterZipFile) = LStrip(lstOriginalName.List(CounterZipFile), "[")
    lstOriginalName.List(CounterZipFile) = LStrip(lstOriginalName.List(CounterZipFile), "\")
    lstOriginalName.List(CounterZipFile) = LStrip(lstOriginalName.List(CounterZipFile), "/")
    lstOriginalName.List(CounterZipFile) = LStrip(lstOriginalName.List(CounterZipFile), "]")
    lstOriginalName.List(CounterZipFile) = LStrip(lstOriginalName.List(CounterZipFile), "-")
    lstOriginalName.List(CounterZipFile) = LStrip(lstOriginalName.List(CounterZipFile), "=")
    lstOriginalName.List(CounterZipFile) = LStrip(lstOriginalName.List(CounterZipFile), "+")
    lstOriginalName.List(CounterZipFile) = LStrip(lstOriginalName.List(CounterZipFile), "|")
    lstOriginalName.List(CounterZipFile) = LStrip(lstOriginalName.List(CounterZipFile), "?")
    lstOriginalName.List(CounterZipFile) = LStrip(lstOriginalName.List(CounterZipFile), " ")
    'The right side
    lstOriginalName.List(CounterZipFile) = RStrip(lstOriginalName.List(CounterZipFile), "_")
    lstOriginalName.List(CounterZipFile) = RStrip(lstOriginalName.List(CounterZipFile), "*")
    lstOriginalName.List(CounterZipFile) = RStrip(lstOriginalName.List(CounterZipFile), "[")
    lstOriginalName.List(CounterZipFile) = RStrip(lstOriginalName.List(CounterZipFile), "\")
    lstOriginalName.List(CounterZipFile) = RStrip(lstOriginalName.List(CounterZipFile), "/")
    lstOriginalName.List(CounterZipFile) = RStrip(lstOriginalName.List(CounterZipFile), "]")
    lstOriginalName.List(CounterZipFile) = RStrip(lstOriginalName.List(CounterZipFile), "-")
    lstOriginalName.List(CounterZipFile) = RStrip(lstOriginalName.List(CounterZipFile), "=")
    lstOriginalName.List(CounterZipFile) = RStrip(lstOriginalName.List(CounterZipFile), "+")
    lstOriginalName.List(CounterZipFile) = RStrip(lstOriginalName.List(CounterZipFile), "|")
    lstOriginalName.List(CounterZipFile) = RStrip(lstOriginalName.List(CounterZipFile), "?")
    lstOriginalName.List(CounterZipFile) = RStrip(lstOriginalName.List(CounterZipFile), " ")
    'Replacing invalid characters ie \ / : * ? " < > | from the middle parts
    lstOriginalName.List(CounterZipFile) = Replace(lstOriginalName.List(CounterZipFile), "\", "-")
    lstOriginalName.List(CounterZipFile) = Replace(lstOriginalName.List(CounterZipFile), "/", "-")
    lstOriginalName.List(CounterZipFile) = Replace(lstOriginalName.List(CounterZipFile), ":", "-")
    lstOriginalName.List(CounterZipFile) = Replace(lstOriginalName.List(CounterZipFile), "*", "-")
    lstOriginalName.List(CounterZipFile) = Replace(lstOriginalName.List(CounterZipFile), "?", "-")
    lstOriginalName.List(CounterZipFile) = Replace(lstOriginalName.List(CounterZipFile), Chr(34), "-")
    lstOriginalName.List(CounterZipFile) = Replace(lstOriginalName.List(CounterZipFile), "<", "-")
    lstOriginalName.List(CounterZipFile) = Replace(lstOriginalName.List(CounterZipFile), ">", "-")
    lstOriginalName.List(CounterZipFile) = Replace(lstOriginalName.List(CounterZipFile), "|", "-")
End Sub


