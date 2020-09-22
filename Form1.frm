VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Port Sniffer by TheLemon"
   ClientHeight    =   1935
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   4335
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1935
   ScaleWidth      =   4335
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   1935
      Left            =   2160
      TabIndex        =   7
      Top             =   0
      Width           =   2175
      Begin MSWinsockLib.Winsock Winsock8 
         Left            =   1560
         Top             =   0
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin MSWinsockLib.Winsock Winsock7 
         Left            =   1080
         Top             =   0
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin MSWinsockLib.Winsock Winsock6 
         Left            =   600
         Top             =   0
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin MSWinsockLib.Winsock Winsock5 
         Left            =   120
         Top             =   0
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Pause"
         Enabled         =   0   'False
         Height          =   255
         Left            =   1200
         TabIndex        =   15
         Top             =   600
         Width           =   855
      End
      Begin VB.ListBox List1 
         Height          =   840
         Left            =   120
         TabIndex        =   11
         Top             =   960
         Width           =   1935
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label5 
         Caption         =   "Open ports:"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "Scanning port:"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1935
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2055
      Begin MSWinsockLib.Winsock Winsock10 
         Left            =   480
         Top             =   480
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin MSWinsockLib.Winsock Winsock9 
         Left            =   0
         Top             =   480
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin MSWinsockLib.Winsock Winsock4 
         Left            =   1440
         Top             =   0
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin MSWinsockLib.Winsock Winsock3 
         Left            =   960
         Top             =   0
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin MSWinsockLib.Winsock Winsock2 
         Left            =   480
         Top             =   0
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   1440
         TabIndex        =   13
         Text            =   "1"
         Top             =   1560
         Width           =   495
      End
      Begin MSWinsockLib.Winsock Winsock1 
         Left            =   0
         Top             =   0
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   500
         Left            =   120
         Top             =   840
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Start"
         Enabled         =   0   'False
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   1560
         Width           =   615
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   1200
         TabIndex        =   6
         Text            =   "32000"
         Top             =   1200
         Width           =   735
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   120
         TabIndex        =   4
         Text            =   "1"
         Top             =   1200
         Width           =   735
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Caption         =   "Timeout:"
         Height          =   255
         Left            =   720
         TabIndex        =   14
         Top             =   1605
         Width           =   735
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "To"
         Height          =   255
         Left            =   840
         TabIndex        =   5
         Top             =   1245
         Width           =   375
      End
      Begin VB.Label Label2 
         Caption         =   "Range of ports to scan:"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   960
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "IP Address to scan:"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuLogToFile 
         Caption         =   "Log to file"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim PortToScan As Integer
Dim PortToStopOn As Integer
Dim Port1 As Integer
Dim Port2 As Integer
Dim Port3 As Integer
Dim Port4 As Integer
Dim Port5 As Integer
Dim Port6 As Integer
Dim Port7 As Integer
Dim Port8 As Integer
Dim Port9 As Integer
Dim Port10 As Integer
Dim TimeOut1
Dim TimeOut2
Dim TimeOut3
Dim TimeOut4
Dim TimeOut5
Dim TimeOut6
Dim TimeOut7
Dim TimeOut8
Dim TimeOut9
Dim TimeOut10
Dim Paused
Dim FileText
Dim TimerThing

Private Sub Command1_Click()
TimerThing = 0
Winsock1.Close
Winsock2.Close
Winsock3.Close
Winsock4.Close
Winsock5.Close
Winsock6.Close
Winsock7.Close
Winsock8.Close
Winsock9.Close
Winsock10.Close
If Text2.Text = "" Then
Text2.Text = "1"
End If
If Text3.Text = "" Then
Text3.Text = "32000"
End If
Winsock1.RemoteHost = Text1.Text
Winsock2.RemoteHost = Text1.Text
Winsock3.RemoteHost = Text1.Text
Winsock4.RemoteHost = Text1.Text
Winsock5.RemoteHost = Text1.Text
Winsock6.RemoteHost = Text1.Text
Winsock7.RemoteHost = Text1.Text
Winsock8.RemoteHost = Text1.Text
Winsock9.RemoteHost = Text1.Text
Winsock10.RemoteHost = Text1.Text
PortToScan = Text2.Text
PortToStopOn = Text3.Text
Text1.Enabled = False
Text2.Enabled = False
Text3.Enabled = False
Text5.Enabled = False
If Text5.Text = "" Then
Text5.Text = "1"
End If
Timer1.Enabled = True
Command1.Enabled = False
If Paused <> 1 Then
List1.Clear
Text4.Text = "Starting..."
End If
Command2.Enabled = True
Paused = 0
If mnuLogToFile.Checked = True Then
Open "PortLog.txt" For Output As #1
Close #1
CopyFileText
Open "PortLog.txt" For Output As #2
Write #2, FileText & vbCrLf & vbCrLf & Winsock1.RemoteHost & vbCrLf
Close #2
End If
mnuLogToFile.Enabled = False
End Sub

Private Sub Command2_Click()
Winsock1.Close
Text1.Enabled = True
Text2.Enabled = True
Text2.Text = PortToScan - 1
Text3.Enabled = True
Text5.Enabled = True
Command1.Enabled = True
Text4.Text = "Paused"
Timer1.Enabled = False
Command1.Enabled = True
Command2.Enabled = False
Paused = 1
End Sub

Private Sub Form_Load()
Form1.Caption = "Port Sniffer " & App.Major & "." & App.Minor & App.Revision & " by TheLemon"
Timer1.Interval = 1
End Sub

Private Sub mnuAbout_Click()
MsgBox "This program is pretty self explanitory.  It scans any computer you" & vbCrLf & _
       "want for any and all open ports." & vbCrLf & vbCrLf & _
       "Timeout:  This is the number of milliseconds the program should wait for" & vbCrLf & _
       "a connection on each port before passing it and trying the next one." & vbCrLf & _
       "I would suggest pinging the destination address first to find out how" & vbCrLf & _
       "clean your connection is with their computer.  If the ping time averages" & vbCrLf & _
       "between 0-300ms, you can use 1-300 as the timeout.  Anything above 300ms" & vbCrLf & _
       "though, I would suggest using a higher timeout time than that.  The result" & vbCrLf & _
       "of a timeout that's too fast is that this port sniffer may sniff ports" & vbCrLf & _
       "too fast for it to know if they're open.  If the timeout is set too slow," & vbCrLf & _
       "you'll have to wait a lot longer than you have to for the program to" & vbCrLf & _
       "finish it's scan." & vbCrLf & _
       "Log to file:  Currently doesn't work very well.  The log file, though," & vbCrLf & _
       "is PortLog.txt, and is in the same directory as this program." & vbCrLf & _
       "You can expect a scan of all 32000 ports to take about 15 hours at a" & vbCrLf & _
       "timeout of .001 seconds, assuming Winsock doesn't overload.  Have fun!" & vbCrLf, _
       vbOKOnly, "About Port Sniffer " & App.Major & "." & App.Minor & App.Revision
End Sub

Private Sub mnuExit_Click()
End
End Sub

Private Sub mnuLogToFile_Click()
If mnuLogToFile.Checked = True Then
mnuLogToFile.Checked = False
Else
mnuLogToFile.Checked = True
End If
End Sub

Private Sub Text1_Change()
If Text1.Text <> "" Then
Command1.Enabled = True
Else
Command1.Enabled = False
End If
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
Command1_Click
End If
End Sub

Private Sub Text2_Click()
Text2.SelStart = 0
Text2.SelLength = Len(Text2.Text)
End Sub

Private Sub Text3_Click()
Text3.SelStart = 0
Text3.SelLength = Len(Text3.Text)
End Sub

Private Sub Text5_Change()
If Text5.Text <> "" Then
Timer1.Interval = Text5.Text
End If
End Sub

Private Sub Text5_Click()
Text5.SelStart = 0
Text5.SelLength = Len(Text5.Text)
End Sub

Private Sub Timer1_Timer()
Text4.Text = PortToScan
TimerThing = TimerThing + 1

If TimerThing = 10 Then
 If TimeOut1 = 1 Then
  If Winsock1.State = sckConnected Then
  Winsock1.Close
  List1.AddItem Port1
   If mnuLogToFile.Checked = True Then
   CopyFileText
   Open "PortLog.txt" For Output As #3
   Write #3, FileText & Port1 & vbCrLf
   Close #3
   End If
  Beep
  TimeOut1 = 0
  PortToScan = PortToScan + 1
  Else
  Winsock1.Close
  TimeOut1 = 0
  PortToScan = PortToScan + 1
  End If
 End If
End If

If TimerThing = 1 Then
 If TimeOut1 <> 1 Then
  If PortToScan <= PortToStopOn Then
  Winsock1.RemotePort = PortToScan
  Port1 = PortToScan
  Winsock1.Connect
  TimeOut1 = 1
  Else
  Text1.Enabled = True
  Text2.Enabled = True
  Text3.Enabled = True
  Text5.Enabled = True
  Command1.Enabled = True
  Text4.Text = "Done"
  Command2.Enabled = False
  Timer1.Enabled = False
  mnuLogToFile.Enabled = True
  End If
 End If
End If

'

If TimerThing = 1 Then
 If TimeOut2 = 1 Then
  If Winsock2.State = sckConnected Then
  Winsock2.Close
  List1.AddItem Port2
   If mnuLogToFile.Checked = True Then
   CopyFileText
   Open "PortLog.txt" For Output As #3
   Write #3, FileText & Port2 & vbCrLf
   Close #3
   End If
  Beep
  TimeOut2 = 0
  PortToScan = PortToScan + 1
  Else
  Winsock2.Close
  TimeOut2 = 0
  PortToScan = PortToScan + 1
  End If
 End If
End If

If TimerThing = 2 Then
 If TimeOut2 <> 1 Then
  If PortToScan <= PortToStopOn Then
  Winsock2.RemotePort = PortToScan
  Port2 = PortToScan
  Winsock2.Connect
  TimeOut2 = 1
  Else
  Text1.Enabled = True
  Text2.Enabled = True
  Text3.Enabled = True
  Text5.Enabled = True
  Command1.Enabled = True
  Text4.Text = "Done"
  Command2.Enabled = False
  Timer1.Enabled = False
  mnuLogToFile.Enabled = True
  End If
 End If
End If

'

If TimerThing = 2 Then
 If TimeOut3 = 1 Then
  If Winsock3.State = sckConnected Then
  Winsock3.Close
  List1.AddItem Port3
   If mnuLogToFile.Checked = True Then
   CopyFileText
   Open "PortLog.txt" For Output As #3
   Write #3, FileText & Port3 & vbCrLf
   Close #3
   End If
  Beep
  TimeOut3 = 0
  PortToScan = PortToScan + 1
  Else
  Winsock3.Close
  TimeOut3 = 0
  PortToScan = PortToScan + 1
  End If
 End If
End If

If TimerThing = 3 Then
 If TimeOut3 <> 1 Then
  If PortToScan <= PortToStopOn Then
  Winsock3.RemotePort = PortToScan
  Port3 = PortToScan
  Winsock3.Connect
  TimeOut3 = 1
  Else
  Text1.Enabled = True
  Text2.Enabled = True
  Text3.Enabled = True
  Text5.Enabled = True
  Command1.Enabled = True
  Text4.Text = "Done"
  Command2.Enabled = False
  Timer1.Enabled = False
  mnuLogToFile.Enabled = True
  End If
 End If
End If

'

If TimerThing = 3 Then
 If TimeOut4 = 1 Then
  If Winsock4.State = sckConnected Then
  Winsock4.Close
  List1.AddItem Port4
   If mnuLogToFile.Checked = True Then
   CopyFileText
   Open "PortLog.txt" For Output As #3
   Write #3, FileText & Port4 & vbCrLf
   Close #3
   End If
  Beep
  TimeOut4 = 0
  PortToScan = PortToScan + 1
  Else
  Winsock4.Close
  TimeOut4 = 0
  PortToScan = PortToScan + 1
  End If
 End If
End If

If TimerThing = 4 Then
 If TimeOut4 <> 1 Then
  If PortToScan <= PortToStopOn Then
  Winsock4.RemotePort = PortToScan
  Port4 = PortToScan
  Winsock4.Connect
  TimeOut4 = 1
  Else
  Text1.Enabled = True
  Text2.Enabled = True
  Text3.Enabled = True
  Text5.Enabled = True
  Command1.Enabled = True
  Text4.Text = "Done"
  Command2.Enabled = False
  Timer1.Enabled = False
  mnuLogToFile.Enabled = True
  End If
 End If
End If

'

If TimerThing = 4 Then
 If TimeOut5 = 1 Then
  If Winsock5.State = sckConnected Then
  Winsock5.Close
  List1.AddItem Port5
   If mnuLogToFile.Checked = True Then
   CopyFileText
   Open "PortLog.txt" For Output As #3
   Write #3, FileText & Port5 & vbCrLf
   Close #3
   End If
  Beep
  TimeOut5 = 0
  PortToScan = PortToScan + 1
  Else
  Winsock5.Close
  TimeOut5 = 0
  PortToScan = PortToScan + 1
  End If
 End If
End If

If TimerThing = 5 Then
 If TimeOut5 <> 1 Then
  If PortToScan <= PortToStopOn Then
  Winsock5.RemotePort = PortToScan
  Port5 = PortToScan
  Winsock5.Connect
  TimeOut5 = 1
  Else
  Text1.Enabled = True
  Text2.Enabled = True
  Text3.Enabled = True
  Text5.Enabled = True
  Command1.Enabled = True
  Text4.Text = "Done"
  Command2.Enabled = False
  Timer1.Enabled = False
  mnuLogToFile.Enabled = True
  End If
 End If
End If

'

If TimerThing = 5 Then
 If TimeOut6 = 1 Then
  If Winsock6.State = sckConnected Then
  Winsock6.Close
  List1.AddItem Port6
   If mnuLogToFile.Checked = True Then
   CopyFileText
   Open "PortLog.txt" For Output As #3
   Write #3, FileText & Port6 & vbCrLf
   Close #3
   End If
  Beep
  TimeOut6 = 0
  PortToScan = PortToScan + 1
  Else
  Winsock6.Close
  TimeOut6 = 0
  PortToScan = PortToScan + 1
  End If
 End If
End If

If TimerThing = 6 Then
 If TimeOut6 <> 1 Then
  If PortToScan <= PortToStopOn Then
  Winsock6.RemotePort = PortToScan
  Port6 = PortToScan
  Winsock6.Connect
  TimeOut6 = 1
  Else
  Text1.Enabled = True
  Text2.Enabled = True
  Text3.Enabled = True
  Text5.Enabled = True
  Command1.Enabled = True
  Text4.Text = "Done"
  Command2.Enabled = False
  Timer1.Enabled = False
  mnuLogToFile.Enabled = True
  End If
 End If
End If

'

If TimerThing = 6 Then
 If TimeOut7 = 1 Then
  If Winsock7.State = sckConnected Then
  Winsock7.Close
  List1.AddItem Port7
   If mnuLogToFile.Checked = True Then
   CopyFileText
   Open "PortLog.txt" For Output As #3
   Write #3, FileText & Port7 & vbCrLf
   Close #3
   End If
  Beep
  TimeOut7 = 0
  PortToScan = PortToScan + 1
  Else
  Winsock7.Close
  TimeOut7 = 0
  PortToScan = PortToScan + 1
  End If
 End If
End If

If TimerThing = 7 Then
 If TimeOut7 <> 1 Then
  If PortToScan <= PortToStopOn Then
  Winsock7.RemotePort = PortToScan
  Port7 = PortToScan
  Winsock7.Connect
  TimeOut7 = 1
  Else
  Text1.Enabled = True
  Text2.Enabled = True
  Text3.Enabled = True
  Text5.Enabled = True
  Command1.Enabled = True
  Text4.Text = "Done"
  Command2.Enabled = False
  Timer1.Enabled = False
  mnuLogToFile.Enabled = True
  End If
 End If
End If

'

If TimerThing = 7 Then
 If TimeOut8 = 1 Then
  If Winsock8.State = sckConnected Then
  Winsock8.Close
  List1.AddItem Port8
   If mnuLogToFile.Checked = True Then
   CopyFileText
   Open "PortLog.txt" For Output As #3
   Write #3, FileText & Port8 & vbCrLf
   Close #3
   End If
  Beep
  TimeOut8 = 0
  PortToScan = PortToScan + 1
  Else
  Winsock8.Close
  TimeOut8 = 0
  PortToScan = PortToScan + 1
  End If
 End If
End If

If TimerThing = 8 Then
 If TimeOut8 <> 1 Then
  If PortToScan <= PortToStopOn Then
  Winsock8.RemotePort = PortToScan
  Port8 = PortToScan
  Winsock8.Connect
  TimeOut8 = 1
  Else
  Text1.Enabled = True
  Text2.Enabled = True
  Text3.Enabled = True
  Text5.Enabled = True
  Command1.Enabled = True
  Text4.Text = "Done"
  Command2.Enabled = False
  Timer1.Enabled = False
  mnuLogToFile.Enabled = True
  End If
 End If
End If

'

If TimerThing = 8 Then
 If TimeOut9 = 1 Then
  If Winsock9.State = sckConnected Then
  Winsock9.Close
  List1.AddItem Port9
   If mnuLogToFile.Checked = True Then
   CopyFileText
   Open "PortLog.txt" For Output As #3
   Write #3, FileText & Port9 & vbCrLf
   Close #3
   End If
  Beep
  TimeOut9 = 0
  PortToScan = PortToScan + 1
  Else
  Winsock9.Close
  TimeOut9 = 0
  PortToScan = PortToScan + 1
  End If
 End If
End If

If TimerThing = 9 Then
 If TimeOut9 <> 1 Then
  If PortToScan <= PortToStopOn Then
  Winsock9.RemotePort = PortToScan
  Port9 = PortToScan
  Winsock9.Connect
  TimeOut9 = 1
  Else
  Text1.Enabled = True
  Text2.Enabled = True
  Text3.Enabled = True
  Text5.Enabled = True
  Command1.Enabled = True
  Text4.Text = "Done"
  Command2.Enabled = False
  Timer1.Enabled = False
  mnuLogToFile.Enabled = True
  End If
 End If
End If

'

If TimerThing = 9 Then
 If TimeOut10 = 1 Then
  If Winsock10.State = sckConnected Then
  Winsock10.Close
  List1.AddItem Port10
   If mnuLogToFile.Checked = True Then
   CopyFileText
   Open "PortLog.txt" For Output As #3
   Write #3, FileText & Port10 & vbCrLf
   Close #3
   End If
  Beep
  TimeOut10 = 0
  PortToScan = PortToScan + 1
  Else
  Winsock10.Close
  TimeOut10 = 0
  PortToScan = PortToScan + 1
  End If
 End If
End If

If TimerThing = 10 Then
 If TimeOut10 <> 1 Then
  If PortToScan <= PortToStopOn Then
  Winsock10.RemotePort = PortToScan
  Port10 = PortToScan
  Winsock10.Connect
  TimeOut10 = 1
  TimerThing = 0
  Else
  Text1.Enabled = True
  Text2.Enabled = True
  Text3.Enabled = True
  Text5.Enabled = True
  Command1.Enabled = True
  Text4.Text = "Done"
  Command2.Enabled = False
  Timer1.Enabled = False
  mnuLogToFile.Enabled = True
  End If
 End If
End If
End Sub

Public Sub CopyFileText()
Open "PortLog.txt" For Binary As #4
Close #4
Open "PortLog.txt" For Input As #5
 If Not EOF(5) Then
 Input #5, FileText
 End If
Close #5
End Sub
