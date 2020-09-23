Attribute VB_Name = "Module1"
Option Explicit
Enum Winzip
UNZIP = 1
ZIP = 0
End Enum

    'Name: WinZipit
    'V1.0 By: renyi [ace]
    'V1.1 By: JosGroen@hotmail.com
    'example: WinZipit "D:\pscDownload\achoopump vs "ladybug 2playergame.zip","C:\unziphippie1",UnZip
Public Sub WinZipit(ByVal strSource As String, _
                    ByVal strTarget As String, _
                    Mode As Winzip)
  
  Dim strWinZip         As String 'string for winzip
  Dim strWinZiplocation As String 'location of Winzip
  Dim RetVal
    strWinZiplocation = "C:\Program Files\WinZip\WINZIP32.EXE"
    
    Select Case Mode
     Case Winzip.ZIP
        strWinZip = strWinZiplocation & " -a " & Chr$(34) & strTarget & Chr$(34) & "; " & Chr$(34) & strSource & Chr$(34)
     Case Winzip.UNZIP
        strWinZip = strWinZiplocation & " -e " & Chr$(34) & strSource & Chr$(34) & "; " & Chr$(34) & strTarget & Chr$(34)
     Case Else
    End Select
    
    RetVal = Shell(strWinZip, vbHide)
End Sub
'Strip strInput specified character from start of string
'From Code Fixer 1.1.48, by Roger Gilchrist
Public Function LStrip(ByVal strInput As String, ByVal strStrip As String) As String
If Left$(strInput, 1) = strStrip Then
Do
strInput = Mid$(strInput, 2)
Loop While Left$(strInput, 1) = strStrip
End If
LStrip = strInput
End Function
'Strip strInput specified character from end of string
'From Code Fixer 1.1.48, by Roger Gilchrist
Public Function RStrip(ByVal strInput As String, _
                       ByVal strStrip As String) As String

  If Right$(strInput, 1) = strStrip Then
    Do
      strInput = Left$(strInput, Len(strInput) - 1)
    Loop While Right$(strInput, 1) = strStrip
  End If
  RStrip = strInput
End Function

