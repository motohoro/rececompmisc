'The MIT License
'Copyright (c) 2012 Mitsuhiro Matsumoto

'Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:
'The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.
'THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
'========================================================================
Option Explicit
Dim objShell,objFSO,objRegExp
Dim tgtdir3,tgtfile3,tgtfile3mpa,sendsetting,objsettingfile

Set objShell = WScript.CreateObject("WScript.Shell")
Set objFSO = WScript.CreateObject("Scripting.FileSystemObject")
Set objRegExp = New RegExp
objRegExp.Pattern = "^OUT=(.+)$"

tgtdir3 = ""
sendsetting = "C:\Program Files\CASIO\本社送信業務\Setting.ini"

Sub BackupCall()
  Dim predir
  predir = objShell.CurrentDirectory
  objShell.CurrentDirectory="C:\Unike\P-POS\App"
  objShell.Run("C:\Unike\P-POS\App\PPBackup.exe" )
  objShell.CurrentDirectory=predir
End Sub

If objFSO.FileExists(sendsetting)=True Then
  Set objsettingfile = objFSO.OpenTextFile(sendsetting,1)
  If Err.Number=0 Then
    Dim linee,objMatches
    Do Until objsettingfile.AtEndOfStream = true
      linee =objsettingfile.ReadLine
      Set objMatches = objRegExp.Execute(linee)
      If objMatches.Count >0 Then
'        WScript.Echo objMatches(0).SubMatches(0)
        tgtdir3 = objMatches(0).SubMatches(0)
      End If
    Loop
    objsettingfile.Close
    tgtfile3 =  mid(Now,3,2) & mid(Now,6,2) & mid(Now,9,2) & ".lzh"
    tgtfile3 = tgtdir3 & tgtfile3
    tgtfile3mpa = tgtdir3 & "日次薬品使用量_" & mid(Now,1,4) & mid(Now,6,2) & mid(Now,9,2) & ".csv"
'============================================
    If objFSO.FileExists(tgtfile3) = False Or objFSO.FileExists(tgtfile3mpa) = False Then
'      retval = MsgBox ("データ作成、送信業務が完了してません"&Chr(13)&Chr(10)&"このまま電源ＯＦＦしますか？",4+16+256)
      retval = MsgBox ("データ作成、送信業務が完了してません"&Chr(13)&Chr(10)&"このままバックアップしますか？",4+16+256)
      Dim retval
      If retval=6 Then
'        objShell.Run("shutdown -s -t 0")
        BackupCall
'MsgBox(objShell.CurrentDirectory)
      End If
    Else
'      objShell.Run("shutdown -s -t 0")
      BackupCall
    End If
  Else
    MsgBox("Error1 Bye")
  End If
Else
  MsgBox("Error2 Bye")
End If