VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Adblocker 1.1"
   ClientHeight    =   1125
   ClientLeft      =   0
   ClientTop       =   285
   ClientWidth     =   3195
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1125
   ScaleWidth      =   3195
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSWinsockLib.Winsock wsock 
      Index           =   0
      Left            =   3390
      Top             =   2280
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin RichTextLib.RichTextBox rtb2 
      Height          =   990
      Left            =   3345
      TabIndex        =   6
      Top             =   1185
      Width           =   1440
      _ExtentX        =   2540
      _ExtentY        =   1746
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"Form1.frx":0000
   End
   Begin RichTextLib.RichTextBox rtb1 
      Height          =   1035
      Left            =   3315
      TabIndex        =   5
      Top             =   30
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   1826
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"Form1.frx":00C9
   End
   Begin VB.CommandButton cmdhide 
      Caption         =   "&Hide"
      Height          =   345
      Left            =   1875
      TabIndex        =   2
      Top             =   600
      Width           =   1125
   End
   Begin VB.CommandButton cmdconnect 
      Caption         =   "&Connect"
      Height          =   345
      Left            =   1875
      TabIndex        =   1
      Top             =   135
      Width           =   1125
   End
   Begin VB.CheckBox chkkillcookies 
      Caption         =   "&Kill Cookies"
      Height          =   255
      Left            =   285
      TabIndex        =   0
      Top             =   645
      Width           =   1125
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   6585
      Top             =   0
   End
   Begin VB.Frame Frame1 
      Caption         =   "&Options"
      Height          =   930
      Left            =   135
      TabIndex        =   3
      Top             =   75
      Width           =   1515
      Begin VB.CheckBox chkblockads 
         Caption         =   "&Block Ads"
         Height          =   255
         Left            =   135
         TabIndex        =   4
         Top             =   240
         Width           =   1155
      End
   End
   Begin VB.Menu mnuconfigure 
      Caption         =   "&Configure"
      Visible         =   0   'False
      Begin VB.Menu mnuexit 
         Caption         =   "&Exit"
      End
      Begin VB.Menu mnusep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuabout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' The first submission I made really messed up the
'hosts file on some systems. If you have downloaded
'the previous version before then please check your
'systems hosts file and edit/delete it if required.

' webserver coding by T-3-T-3@gmx.li

Private Type NOTIFYICONDATA
   cbSize As Long
   hWnd As Long
   uId As Long
   uFlags As Long
   uCallBackMessage As Long
   hIcon As Long
   szTip As String * 64
End Type

Private Type OSVERSIONINFO
        dwOSVersionInfoSize As Long
        dwMajorVersion As Long
        dwMinorVersion As Long
        dwBuildNumber As Long
        dwPlatformId As Long
        szCSDVersion As String * 128      '  Maintenance string for PSS usage
End Type

Private Const INTERNET_DIAL_FORCE_PROMPT = &H2000

Private Const NIM_ADD = &H0
Private Const NIM_MODIFY = &H1
Private Const NIM_DELETE = &H2

Private Const WM_MOUSEMOVE = &H200

Private Const NIF_MESSAGE = &H1
Private Const NIF_ICON = &H2
Private Const NIF_TIP = &H4
Private Const WM_LBUTTONDBLCLK = &H203   'Double-click
Private Const WM_LBUTTONDOWN = &H201     'Button down
Private Const WM_LBUTTONUP = &H202       'Button up
Private Const WM_RBUTTONDBLCLK = &H206   'Double-click
Private Const WM_RBUTTONDOWN = &H204     'Button down
Private Const WM_RBUTTONUP = &H205       'Button up

Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean
Private Declare Function SetFileAttributes Lib "kernel32" Alias "SetFileAttributesA" (ByVal lpFileName As String, ByVal dwFileAttributes As Long) As Long
Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Private Declare Function InternetDial Lib "wininet.dll" (ByVal hwndParent As Long, ByVal lpszConnectoid As String, ByVal dwFlags As Long, lpdwConnection As Long, ByVal dwReserved As Long) As Long

Dim nid As NOTIFYICONDATA
Dim strwindir As String
Dim struname As String
Dim strhostsfile As String
Dim connections As Integer
Private Sub chkblockads_Click()
' switch adblocking on (or off)
 blockads (chkblockads.Value)
End Sub

Private Sub chkkillcookies_Click()
' switch cookie killing on (or off)
 killcookies (chkkillcookies.Value)
End Sub

Private Sub cmdhide_Click()
' hide the form
 Me.Hide
End Sub

Private Sub cmdconnect_Click()
On Error GoTo err
'connect to internet using the internet dial function
If InternetDial(0&, vbNull, INTERNET_DIAL_FORCE_PROMPT, i, 0&) = 0 Then
 ' hide the form
 Me.Hide
End If
Exit Sub

err:
 MsgBox Str(err.Number) + " : " + err.Description, vbCritical + vbOKOnly, "Error"
End Sub

Private Sub Form_Load()
' terminate if app is already running
If App.PrevInstance Then
 MsgBox "Application is already running" + vbNewLine + "This instance will terminate", vbInformation + vbOKOnly, "Error"
 End
End If

On Error GoTo err
Dim OSInfo As OSVERSIONINFO
Dim r As Integer
strwindir = Space(255)
struname = Space(255)

rtb1.Text = " "
rtb2.Text = " "

' get windows directory name
r = GetWindowsDirectory(strwindir, Len(strwindir))
If r = 0 Then
 MsgBox "Error retrieveing Windows directory name" + vbNewLine + "C:\WINDOWS will be used as the default windows directory", vbCritical + vbOKOnly, "Error"
 strwindir = "C:\WINDOWS"
Else
 strwindir = Left(strwindir, r)
End If

' get the user name
r = GetUserName(struname, Len(struname))
If r = 0 Then
 MsgBox "Error retrieving User Name", vbCritical + vbOKOnly, "Error"
Else
 struname = Left(struname, r)
End If

' get os information
OSInfo.dwOSVersionInfoSize = Len(OSInfo)

If GetVersionEx(OSInfo) = 0 Then
 MsgBox "Error retrieving Operating System Information", vbCritical + vbOKOnly, "Error"
Else
 Select Case OSInfo.dwPlatformId
  Case 1
  ' if os=win95/98, then hosts file is windowsdir\hosts
     strhostsfile = strwindir + "\hosts"
     'cookies folder=windir\cookies
     strwindir = strwindir + "\cookies"
  Case 2
  ' if os=winnt, hosts file is windir\system32\drivers\etc\hosts
     strhostsfile = strwindir + "\system32\drivers\etc\hosts"
     'cookies folder=windir\username\cookies
     strwindir = strwindir + struname + "\cookies"
 End Select
End If

' if hosts file is already the, load it in rtb2 to restore it later
If Dir(strhostsfile) <> "" Then rtb2.LoadFile strhostsfile, 1

' init systray icon etc. params & add icon to systray
With nid
 .cbSize = Len(nid)
 .hWnd = Form1.hWnd
 .uId = vbNull
 .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
 .uCallBackMessage = WM_MOUSEMOVE
 .hIcon = Form1.Icon
 .szTip = "AdBlocker 1.0" & vbNullChar
End With
Shell_NotifyIcon NIM_ADD, nid

' init web server params
connections = 1
Me.wsock(0).Close
Me.wsock(0).LocalPort = 80
Me.wsock(0).Listen

Exit Sub

err:
If err.Number = 10048 Then
 MsgBox "A webserver is already running on port 80, for best results please stop the server and restart this application", vbOKOnly, "Error"
 End
Else
 MsgBox Str(err.Number) + " : " + err.Description, vbCritical + vbOKOnly, "Error"
End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim msg As Long
  msg = X / Screen.TwipsPerPixelX
    Select Case msg
       Case WM_LBUTTONDOWN
       Case WM_LBUTTONUP
       Case WM_LBUTTONDBLCLK
           Me.Show
       Case WM_RBUTTONDOWN
        PopupMenu Form1.mnuconfigure
       Case WM_RBUTTONUP
       Case WM_RBUTTONDBLCLK
    End Select
End Sub

Private Sub Form_Terminate()
 ' remove icon from systray
 Shell_NotifyIcon NIM_DELETE, nid
End Sub

Private Sub killcookies(state As Boolean)
 ' enable cookies killing
 Timer1.Enabled = state
End Sub

Private Sub Form_Unload(Cancel As Integer)
If MsgBox("Do you want to terminate AdBlocker?", vbYesNo, "Adblocker 1.1") = vbYes Then
 ' disable adblocking on form close
 'this will also restore the contents of the
 'original hosts file from rtb2
  blockads (False)
  Me.Hide
Else
 Cancel = True
End If
End Sub

Private Sub mnuabout_Click()
 ' didnt have time to make a stupid about box, this works great
 MsgBox "Adblocker v 1.0" + vbNewLine + "Created by Abhishek Dutta", vbOKOnly, "About"
End Sub

Private Sub mnuexit_Click()
 Unload Me
End Sub

Private Sub Timer1_Timer()
 On Error Resume Next
 ' delete all cookies (*.txt files) from the cookies folder
 strtemp = Dir(strwindir + "\*.txt")
 Kill (strwindir + "\" + Dir(strwindir + "\*.txt"))
End Sub

Private Sub blockads(state As Boolean)
' if adblocking is enabled, load contents from hosts.svr
'and save it to system's hosts file
If state = True Then
  rtb1.LoadFile "hosts.svr", 1
  rtb1.SaveFile strhostsfile, 1
Else
' restore the contents of the original hosts file frm rtb2
  rtb2.SaveFile strhostsfile, 1
End If
rtb1.Text = ""
End Sub

Private Sub wsock_ConnectionRequest(Index As Integer, ByVal requestID As Long)
  If Index = 0 Then
    connections = connections + 1
    Load wsock(connections)
    wsock(connections).LocalPort = 0
    wsock(connections).Accept requestID
  End If
End Sub

Private Sub wsock_DataArrival(Index As Integer, ByVal bytesTotal As Long)
Dim strdata As String
' if the browser request for any ads, send it the adblock.gif file
wsock(Index).GetData strdata$
If Mid$(strdata$, 1, 3) = "GET" Then
  SendPage Page, Index
End If
End Sub

Public Sub SendPage(Page, Index)
' actual file sending
stradpage = App.Path & "\" & "adblock.gif"
On Error GoTo err
Page = "adblock.gif"
  Nr = FreeFile
  Tx$ = " "
  Lg = FileLen(stradpage)
  Open (stradpage) For Binary As Nr
    Tx1$ = ""
    For m = 1 To Lg
      Get #Nr, , Tx$
      Tx1$ = Tx1$ + Tx$
    Next
  Close Nr
  nid.szTip = Str(nadcount)
  wsock(Index).SendData Tx1$
Exit Sub
err:
If err.Number = 53 Then wsock(Index).SendData "The URL you asked for does not exist on this website "
End Sub

Private Sub wsock_SendComplete(Index As Integer)
 wsock(Index).Close
End Sub
