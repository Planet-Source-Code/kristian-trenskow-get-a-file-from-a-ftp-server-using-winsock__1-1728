<div align="center">

## Get a file from a FTP server using winsock\.


</div>

### Description

This function shows how to get a file from an FTP site.
 
### More Info
 
Make a form (form1) and insert two winsock controls (winsock1 and winsock2). Then insert a command button (command1) and three labels (label1, label2 and label3). Then you need to add a module (module1), and last you need a timer (timer1). That's it.

A file


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Kristian Trenskow](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/kristian-trenskow.md)
**Level**          |Unknown
**User Rating**    |5.0 (15 globes from 3 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Internet/ HTML](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/internet-html__1-34.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/kristian-trenskow-get-a-file-from-a-ftp-server-using-winsock__1-1728/archive/master.zip)

### API Declarations

```
Type Com ' the type to the array, that tells the proggie witch command should be send after a certain reply from the server.
 BackCode As String
 Command As String
End Type
```


### Source Code

```
Dim States(4) As Com ' initialize the command/reply array
Dim State As Integer ' tells where in the commucation process we are
Dim Total As Long ' Total data to recieve
Dim Current As Long ' Current data recieved
Dim Old As Long ' a timer1 data value
Dim server as string
dim Username as String
dim password as string
dim LocalFile as String
dim remotefile as string
Private Sub Command1_Click()
Server = "ftp.microsoft.com"
Username = "anonymous"
Password = "guest"
LocalFile = "c:\vbrun60.exe"
Remotefile = "/Softlib/MSLFILES/VBRUN60.EXE"
States(0).BackCode = "220" ' this is the welcome message from server
States(0).Command = "USER " + username ' logges in.
States(1).BackCode = "331" ' "Username ok. Need password" from server
States(1).Command = "PASS " + password ' send the password
States(2).BackCode = "230" ' "Access allowed" massage from server
States(2).Command = "TYPE I" ' Sets the type
States(3).BackCode = "200" ' "TYPE I OK" from server
States(3).Command = "PORT " ' Port command (enhanced features command button click."
States(4).BackCode = "200" ' On port OK
States(4).Command = "RETR " + remotefile ' send request for file
Winsock1.Close
Winsock2.Close
Do Until Winsock1.State = 0 And Winsock2.State = 0
DoEvents
Loop
Winsock1.RemoteHost = Server
Winsock1.RemotePort = 21
Dim nr1 As Long
Dim nr2 As Long
Randomize Timer
nr1 = Int(Rnd * 126) + 1
nr2 = Int(Rnd * 255) + 1
Winsock2.LocalPort = (nr1 * 256) + nr2
Dim IP As String
IP = Winsock2.LocalIP
Do Until InStr(IP, ".") = 0
IP = Left(IP, InStr(IP, ".") - 1) + "," + Right(IP, Len(IP) - InStr(IP, "."))
Loop
States(3).Command = "PORT " + IP + "," + Trim(Str(nr1)) + "," + Trim(Str(nr2))
Winsock2.Listen
Winsock1.Connect
Open localfile For Output As #1
End Sub
Private Sub Timer1_Timer() ' status timer (calculates speed and elabsed time.)
Dim Left As Long
Label2 = Trim(Str((Current - Old) / 512)) + " KB/s"
If (Current - Old) > 0 Then
Left = Total - Current
Label3 = Trim(Str(Left / (Current - Old))) + " Sec left."
Else
Label3 = "dunno"
End If
Old = Current
End Sub
Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long) ' handles the ftp connection
Dim tmpS As String
Winsock1.GetData tmpS, , bytesTotal
If State < 5 Then
If Left(tmpS, 3) = States(State).BackCode Then
Winsock1.SendData States(State).Command + Chr(13) + Chr(10)
Debug.Print States(State).Command + Chr(13) + Chr(10)
State = State + 1
Else
MsgBox "Error! " + Left(tmpS, Len(tmpS) - 2), vbOKOnly + vbCritical, "FTPget"
End If
ElseIf State = 6 Then
Timer1.Enabled = False
MsgBox "Done!", vbOKOnly + vbInformation, "FTPget"
Else
If Left(tmpS, 4) = "150 " Then
Total = Val(Right(tmpS, Len(tmpS) - InStr(tmpS, "(")))
Timer1.Enabled = True
End If
State = State + 1
End If
End Sub
Private Sub Winsock2_Close() ' handles the data connection
Close #1
Winsock1.Close
End Sub
Private Sub Winsock2_ConnectionRequest(ByVal requestID As Long)
Winsock2.Close
Do Until Winsock2.State = 0
DoEvents
Loop
Winsock2.Accept requestID
End Sub
Private Sub Winsock2_DataArrival(ByVal bytesTotal As Long)
Dim tmpS As String
Winsock2.GetData tmpS, , bytesTotal
Print #1, tmpS;
Current = Current + Len(tmpS)
Label1 = Trim(Str(Current)) + " / " + Trim(Str(Total))
End Sub
Private Sub Form_Load()
Timer1.Enabled = False
Timer1.Interval = 500
End Sub
```

