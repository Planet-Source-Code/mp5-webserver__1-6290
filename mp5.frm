VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   Caption         =   "MP5-Webserver"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton res 
      Caption         =   "Rese t"
      Height          =   1095
      Left            =   4320
      TabIndex        =   9
      Top             =   1320
      Width           =   255
   End
   Begin VB.Frame Frame1 
      Caption         =   "Status"
      Height          =   615
      Left            =   0
      TabIndex        =   7
      Top             =   2520
      Width           =   4575
      Begin VB.Label status 
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   4335
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   0
      Top             =   2040
   End
   Begin VB.CommandButton Noli 
      Caption         =   "No Listening"
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   600
      Width           =   1335
   End
   Begin VB.CommandButton Terminate 
      Caption         =   "Terminate"
      Height          =   375
      Left            =   3000
      TabIndex        =   3
      Top             =   360
      Width           =   1455
   End
   Begin VB.CommandButton Li 
      Caption         =   "Listen"
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Width           =   1335
   End
   Begin MSWinsockLib.Winsock ws 
      Left            =   0
      Top             =   1320
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   80
   End
   Begin VB.Label bytes 
      Height          =   255
      Left            =   1920
      TabIndex        =   6
      Top             =   2040
      Width           =   2175
   End
   Begin VB.Label Label2 
      Caption         =   "Bytes sent:"
      Height          =   255
      Left            =   600
      TabIndex        =   5
      Top             =   2040
      Width           =   975
   End
   Begin VB.Label r 
      Height          =   255
      Left            =   1920
      TabIndex        =   1
      Top             =   1440
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "Hits:"
      Height          =   255
      Left            =   600
      TabIndex        =   0
      Top             =   1440
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Ihr braucht zum rekompilieren eine Lizenz fuer die mswinsck.ocx!
'Durch einen Fehler kann die manchmal nicht verfügbar sein!
'Fehlerbehebung:
'Sucht Auf der InstallationsCD nach *.srg Dateien.
'Es sollte eine mit eine mit zu mswinsck ähnlichen Name dabei sein!
'Kopiert sie, öffnet sie im Notepad und fügt ihr als erste Zeile "REGEDIT4" (Ohne ")
'hinzu. Benennt sie in <name>.reg um und klickt doppel auf die datei!
'----------------------------------MP5 Webserver------------------------------------------
'Dies ist ein kleiner Webserver, der nur HTML unterstützt, keine Grafiken!
'Ich werde in ein paar Monaten eine neue Version mit mehr Features herausbringen!
'Ich übernehme KEINE Verantwortung für eventuelle Schäden an Hard- oder Software!
'By Martin Vielsmaier (visual.basic@gmx.de)
'Freeware!
'Wenn ihr Fehler findet, mailt mir!
'////////////////////////////////////////////////////////////////////////////////
'This is a little Webserver, that only HTML supports, but no grafics!
'I will program a new version with more features in some months!
'I´m NOT responsible for damages in your Hardware and software!
'By Martin Vielsmaier (visual.basic@gmx.de)
'This is FREEWARE!!!!!!!!!
'If you find some bugs, email me!

Option Explicit

Private Declare Function WSAGetLastError Lib "WSOCK32.DLL" () _
        As Long
        
Private Declare Function WSAStartup Lib "WSOCK32.DLL" (ByVal _
        wVersionRequired&, lpWSAData As WinSocketDataType) _
        As Long
        
Private Declare Function WSACleanup Lib "WSOCK32.DLL" () _
        As Long
        
Private Declare Function gethostname Lib "WSOCK32.DLL" (ByVal _
        HostName$, ByVal HostLen%) As Long
        
Private Declare Function gethostbyname Lib "WSOCK32.DLL" _
        (ByVal HostName$) As Long
        
Private Declare Function gethostbyaddr Lib "WSOCK32.DLL" _
        (ByVal addr$, ByVal laenge%, ByVal typ%) As Long
        
Private Declare Sub RtlMoveMemory Lib "kernel32" (hpvDest As _
        Any, ByVal hpvSource&, ByVal cbCopy&)

Const WS_VERSION_REQD = &H101
Const WS_VERSION_MAJOR = WS_VERSION_REQD \ &H100 And &HFF&
Const WS_VERSION_MINOR = WS_VERSION_REQD And &HFF&

Const MIN_SOCKETS_REQD = 1
Const SOCKET_ERROR = -1
Const WSADescription_Len = 256
Const WSASYS_Status_Len = 128


Private Type HostDeType
  hName As Long
  hAliases As Long
  hAddrType As Integer
  hLength As Integer
  hAddrList As Long
End Type

Private Type WinSocketDataType
   wversion As Integer
   wHighVersion As Integer
   szDescription(0 To WSADescription_Len) As Byte
   szSystemStatus(0 To WSASYS_Status_Len) As Byte
   iMaxSockets As Integer
   iMaxUdpDg As Integer
   lpszVendorInfo As Long
End Type
'"Normale" Deklarationen, die Darüber sind uninteressant!
Dim ownIPs$(10)
Dim httpbase$
Dim logdic$
Dim dat$
Dim er%
Dim Data$
Dim Htmldata$
Dim Header$
Dim fil$
Dim file$
Dim counter%
Dim i%
Dim fileineed$
Dim inp$
Dim e$
Dim length$
Dim bytessent As Long '"As long" is very important! :-)


Private Sub GetOptions()
Dim Base$
On Error GoTo OptionsNotfound
Open "C:\mp5Options.cfg" For Input As #1
Input #1, Base$
httpbase$ = Base$
Close #1
Exit Sub
OptionsNotfound:
inp$ = InputBox("´mp5Options.cfg´ nicht gefunden! ´mp5Options.cfg´ not found!" + vbNewLine + "HTML  Verzeichnis eingeben: / Input HTML directory", "Error!")
If inp$ = "" Then
httpbase$ = "C:\"
Else
httpbase$ = inp$
End If
inp$ = ""
Open "C:\mp5Options.cfg" For Output As #1
Print #1, httpbase$
Close #1
End Sub

Private Sub GetIPs()
  'Nicht von mir, von www.goetz-reinecke.de
  'This code is from www.goetz-reinecke.de:
  
  Dim X%
  Dim IP$, HOST$
  '--Start Init Sock--
  Dim Result%
  Dim LoBy%, HiBy%
  Dim SocketData As WinSocketDataType
  
    Result = WSAStartup(WS_VERSION_REQD, SocketData)
    If Result <> 0 Then
      MsgBox ("'winsock.dll' antwortet nicht !")
      End
    End If
    
    LoBy = SocketData.wversion And &HFF&
    HiBy = SocketData.wversion \ &H100 And &HFF&
    
    If LoBy < WS_VERSION_MAJOR Or LoBy = WS_VERSION_MAJOR And _
       HiBy < WS_VERSION_MINOR Then
        MsgBox ("Die Windows-Sockets Version " & Trim$(Str$(LoBy)) & _
                "." & Trim$(Str$(HiBy)) & " wird nicht von der '" & _
                "winsock.dll' unterstützt !")
        End
    End If
    
    If SocketData.iMaxSockets < MIN_SOCKETS_REQD Then
      MsgBox ("Diese Anwendung verlangt mindestens " & _
              Trim$(Str$(MIN_SOCKETS_REQD)) & " Sockets !")
      End
    End If
    '--End Init Sock--
    '--Start Get IPs--
     HOST = MyHostName$()
     For X = 1 To 10
        IP = HostByName$(HOST, X - 1)
        If Len(IP) = 0 Then GoTo goon
        ownIPs$(X) = IP
     Next X
goon:
    '--End Get IPs--
    '--Start Clean Sock--
    Result = WSACleanup()
    If Result <> 0 Then
      MsgBox ("Socket Error " & Trim$(Str$(Result)) & _
              " in Prozedur 'CleanSockets' aufgetreten !")
      End
    End If
    '--End Clean Sock--
    End Sub




Private Function HostByName(Name$, Optional X% = 0) As String
  Dim MemIp() As Byte
  Dim Y%
  Dim HostDeAddress&, HostIp&
  Dim IpAddress$
  Dim HOST As HostDeType
  
    HostDeAddress = gethostbyname(Name)
    If HostDeAddress = 0 Then
      HostByName = ""
      Exit Function
    End If
    
    Call RtlMoveMemory(HOST, HostDeAddress, LenB(HOST))
    
    For Y = 0 To X
      Call RtlMoveMemory(HostIp, HOST.hAddrList + 4 * Y, 4)
      If HostIp = 0 Then
        HostByName = ""
        Exit Function
      End If
    Next Y
    
    ReDim MemIp(1 To HOST.hLength)
    Call RtlMoveMemory(MemIp(1), HostIp, HOST.hLength)
    
    IpAddress = ""
    
    For Y = 1 To HOST.hLength
      IpAddress = IpAddress & MemIp(Y) & "."
    Next Y
    
    IpAddress = Left$(IpAddress, Len(IpAddress) - 1)
    HostByName = IpAddress
End Function




Private Function MyHostName() As String
  Dim HostName As String * 256
  
    If gethostname(HostName, 256) = SOCKET_ERROR Then
      MsgBox "Windows Sockets error " & Str(WSAGetLastError())
      Exit Function
    Else
      MyHostName = NextChar(Trim$(HostName), Chr$(0))
    End If
End Function

Private Function NextChar(Text$, Char$) As String
  Dim POS%
    POS = InStr(1, Text, Char)
    If POS = 0 Then
      NextChar = Text
      Text = ""
    Else
      NextChar = Left$(Text, POS - 1)
      Text = Mid$(Text, POS + Len(Char))
    End If
End Function


Private Sub Command1_Click()

End Sub

Private Sub Form_Load()
Li.Enabled = False
Call GetOptions
Call GetIPs
Call Li_Click
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
ws.Close 'Aufraeumen! / Tidy up!
End
End Sub

Private Sub Li_Click()
ws.Close
ws.Listen
Li.Enabled = False
Noli.Enabled = True
status.Caption = "Listening"
End Sub

Private Sub Noli_Click()
ws.Close
Li.Enabled = True
Noli.Enabled = False
status.Caption = "Sleeping"
End Sub

Private Sub res_Click()
bytessent = 0
bytes.Caption = CStr(bytessent)
counter% = 0
r.Caption = CStr(counter%)
End Sub

Private Sub Terminate_Click()
ws.Close 'Aufraeumen! / Tidy up!
End
End Sub

Private Sub Timer1_Timer()
Timer1.Enabled = False
ws.Close
ws.Listen
status.Caption = "Listening"
End Sub

Private Sub ws_ConnectionRequest(ByVal requestID As Long)
ws.Close
ws.Accept requestID
status.Caption = "Connected"
End Sub

Private Sub ws_DataArrival(ByVal bytesTotal As Long)
ws.GetData dat$, vbString 'Wir holen uns den Muell, den der Client sendet./We get the shit from the client!
If Left$(dat$, 3) <> "GET" Then er% = 1
Open "C:\um.tmp" For Output As #1 ' nicht schoen, aber einfach! / Not the best way, but an easy one!
Print #1, dat$
Close #1
Open "C:\um.tmp" For Input As #1
Input #1, Data$
Close #1
Kill "C:\um.tmp"
status.Caption = "Incoming request"
Call request(Data$, dat$)
End Sub

Private Sub ws_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
status.Caption = "Error"
If MsgBox("An error occured:" + vbNewLine + Description, vbRetryCancel, "Error!") = 4 Then
ws.Close  'Aufraeumen! / Tidy up!
Call Form_Load 'Versuchen neuzustarten / Try to restart
Else
ws.Close
End
End If
End Sub

Private Sub request(requdata$, fulldata$)
Htmldata$ = ""
Header$ = ""
fileineed$ = ""
If er% = 1 Then GoTo err 'wenn das ding keine GET-befehl enthaelt springen wir zu e, wo die verbindung getrennt wird!
fil$ = Left(Right(requdata$, Len(requdata$) - 4), Len(requdata$) - 12)
counter% = counter% + 1
r.Caption = CStr(counter%)
'MsgBox fulldata$ 'Falls ihr das sehen wollt, aktivieren! (For Debuging!)
'Open "C:\mp5requests.log" For Append As #1 'Ich will eure Festplatte nicht füllen / I don´t want to fil your HD
'Print #1, Date$ + ", " + Time$ + " :" + vbNewLine
'Print #1, fulldata$
'Close #1
i% = 0
Do
i% = i% + 1
If InStr(fil$, ownIPs$(i%)) Then fil$ = Right(fil$, Len(fil$) - Len(ownIPs$(i%))) 'falls die dateiangabe absolut, sprich mit IP, dann entferne sie!
Loop Until Len(ownIPs$(i%)) = 0 Or i% = 10
file$ = fil$
For i = 1 To Len(file$)
If Asc(Right(Left(file$, i), 1)) = 47 Then
fileineed$ = fileineed$ + "\"
Else
fileineed$ = fileineed$ + Right(Left(file$, i), 1)
End If
Next i
If fileineed$ = "\ " Then fileineed$ = "\index.html" 'Standartseite ist index.html / Standart page is index.html
On Error GoTo err404
Open httpbase$ + fileineed$ For Input As #1
Do
Input #1, inp$
For i = 1 To Len(inp$)
e$ = Right(Left(inp$, i), 1)
If Asc(e$) <> 34 Then Htmldata$ = Htmldata$ + e$ 'Filtere " heraus!
Next i
Loop Until EOF(1) = True
Close #1
length$ = CStr(Len(Htmldata$))
Header$ = "HTTP/1.0 200 OK" + vbNewLine + "Server: MP5" + vbNewLine + "Content-Length: " + length$ + vbNewLine + "Content-Type: text/html" + vbNewLine + "Connection: Keep -Alive" + vbNewLine + "Keep-Alive: timeout=5, max=25"
GoTo sen
err404: 'FILE NOT FOUND!
Header$ = "HTTP/1.0 404 Not found" + vbNewLine + "Server: MP5" + vbNewLine + "Content-Length: 569" + vbNewLine + "Content-Type: text/html"
Htmldata$ = "<html><title>Error</title><body><h1><font face=Arial,Helvetica>The error 404 - Not found occured!</font></h1><font face=Arial,Helvetica>I don't know why but the URL does not exist, retype the URL, perhaps you made a mistake!</font><br><br><h1><font face=Arial,Helvetica>Der Fehler 404 - Nicht gefunden ist aufgetreten!</font></h1><br><font face=Arial,Helvetica>Ich weiss nicht warum, aber die URl existiert nicht, &uuml;berpr&uuml;fen sie die URL!</font><font face=Arial,Helvetica></font><p><hr WIDTH=100%><br>Servertype: MP5 by Martin V.</body></html>"
sen:
ws.SendData Header$ + vbNewLine + Htmldata$
'Open "C:\mp5responses.log" For Append As #1  'Ich will eure Festplatte nicht füllen / I don´t want to fil your HD
'Print #1, Header$ + vbNewLine + Htmldata$ + vbNewLine
'Close #1
bytessent = bytessent + Len(Header$) + 1 + Len(Htmldata$)
bytes.Caption = CStr(bytessent)
GoTo e
err:
ws.Close
ws.Listen
status.Caption = "Listening"
Exit Sub
e:
Timer1.Enabled = True
End Sub

