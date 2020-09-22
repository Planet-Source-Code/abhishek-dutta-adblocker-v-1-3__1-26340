VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About"
   ClientHeight    =   1785
   ClientLeft      =   0
   ClientTop       =   285
   ClientWidth     =   2250
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1785
   ScaleWidth      =   2250
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox lstpops 
      Height          =   1425
      Left            =   3825
      TabIndex        =   7
      Top             =   2175
      Width           =   1050
   End
   Begin VB.ListBox lstFile 
      Height          =   1425
      Left            =   2520
      TabIndex        =   6
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      Height          =   345
      Left            =   585
      TabIndex        =   2
      Top             =   1275
      Width           =   1125
   End
   Begin VB.FileListBox fileCk 
      Height          =   1455
      Left            =   1320
      Pattern         =   "*.txt"
      TabIndex        =   1
      Top             =   2160
      Width           =   1140
   End
   Begin VB.ListBox lstCk 
      Height          =   1425
      Left            =   120
      TabIndex        =   0
      Top             =   2160
      Width           =   1200
   End
   Begin MSWinsockLib.Winsock wsock 
      Index           =   0
      Left            =   795
      Top             =   3690
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Timer Timer 
      Interval        =   100
      Left            =   120
      Top             =   3690
   End
   Begin VB.Image imgicon 
      Height          =   480
      Left            =   105
      Picture         =   "frmMain.frx":0442
      Top             =   90
      Width           =   480
   End
   Begin VB.Label Label4 
      Caption         =   "AdBlocker v 1.3"
      Height          =   180
      Left            =   660
      TabIndex        =   5
      Top             =   240
      Width           =   1230
   End
   Begin VB.Label Label1 
      Caption         =   "Created by Abhishek Dutta"
      Height          =   210
      Left            =   150
      TabIndex        =   4
      Top             =   675
      Width           =   1950
   End
   Begin VB.Label Label3 
      Caption         =   "<abhishekdutta@mail.com>"
      ForeColor       =   &H00FF0000&
      Height          =   180
      Left            =   135
      MouseIcon       =   "frmMain.frx":0884
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   930
      Width           =   1995
   End
   Begin VB.Menu mnuconfigure 
      Caption         =   "&Configure"
      Visible         =   0   'False
      Begin VB.Menu mnuexit 
         Caption         =   "&Exit"
      End
      Begin VB.Menu mnuabout 
         Caption         =   "&About"
      End
      Begin VB.Menu mnusep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnusettings 
         Caption         =   "&Settings"
      End
      Begin VB.Menu mnusep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnublockads 
         Caption         =   "Block &Ads"
      End
      Begin VB.Menu mnukillcookies 
         Caption         =   "Kill &Cookies"
      End
      Begin VB.Menu mnublockpopups 
         Caption         =   "Block &Popups"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' webserver coding by T-3-T-3@gmx.li

Dim nid As NOTIFYICONDATA
Dim strwindir As String
Dim connections As Integer
Dim Status As Boolean
Dim stats(3) As Integer
Dim hnd As Long
Dim hcld As Long
Dim class As String * 16
'

Private Sub cmdOk_Click()
Me.Visible = False
End Sub

Private Sub Form_Load()

Me.Visible = False

' terminate if app is already running
If App.PrevInstance Then
    End
End If

On Error GoTo err

'declare / initialize variables
Dim r As Integer
Dim struname As String * 255

strwindir = Space(255)

For i = 0 To 3
    stats(i) = 0
Next

lstFile.Clear

' get windows directory name
r = GetWindowsDirectory(strwindir, Len(strwindir))

If r = 0 Then
    MsgBox "Error retrieveing Windows directory name" + vbNewLine + "C:\WINDOWS will be used as the default windows directory", vbCritical + vbOKOnly, "Error"
    strwindir = "C:\WINDOWS" 'on error use default
Else
    strwindir = Left(strwindir, r)
End If

' get the user name
r = GetUserName(struname, 255)
struname = Left(struname, InStr(struname, Chr(0)) - 1)

Dim strtemp As String
strtemp = Trim(struname)

'get location of windows hosts file and cookies diretory

If GetVersion() = 0 Then
    'if nt then set appropriate paths for cookie dir and hosts file
    fileCk.Path = (strwindir + "\profiles\" + struname + "\" + "cookies")
    strwindir = (strwindir + "\system32\drivers\etc\hosts")
Else
    'if win95/98 etc then set apt paths
    If Len(strtemp) > 0 Then
    'if user is logged in the cookes dir is \profiles\username\cookies
        If Dir((strwindir + "\profiles\" + strtemp + "\cookies"), vbDirectory Or vbSystem) <> "" Then fileCk.Path = (strwindir + "\profiles\" + strtemp + "\cookies")
    Else
        'else normal cookies dir \cookies
        fileCk.Path = strwindir + "\cookies"
    End If
    strwindir = (strwindir + "\hosts")
End If

' init systray icon etc. params & add icon to systray
With nid
 .cbSize = Len(nid)
 .hwnd = frmMain.hwnd
 .uId = vbNull
 .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
 .uCallBackMessage = WM_MOUSEMOVE
 .hIcon = frmMain.Icon
 .szTip = "AdBlocker 1.3" & vbNullChar
End With

Shell_NotifyIcon NIM_ADD, nid

'set attributes of files to vbnormal
Call SetAttributes

'back up original hosts file, only if it was not restored properly the last time
If GetSetting("AdBlocker", "Data", "Restore", "True") = "True" Then
    CreateBlankFile (strwindir) 'create hosts if it does not exists
    Call LoadFileLst(lstFile, strwindir)
    Call SaveFileLst(lstFile, App.Path + "\data\hosts.bak")
End If
 
'create log files if they dont' exist
Call CreateBlankFile(App.Path + "\ads.log")
Call CreateBlankFile(App.Path + "\cookies.log")

'check and delete large log files
If FileLen(App.Path + "\ads.log") > 16000 Then Call Kill(App.Path + "\ads.log")
If FileLen(App.Path + "\cookies.log") > 16000 Then Call Kill(App.Path + "\cookies.log")

'hosts file changed, store it in registry
Call SaveSetting("AdBlocker", "Data", "Restore", "False")

'get previous settings if any
Call InitSettings

Exit Sub

err:

MsgBox Str(err.Number) + " : " + err.Description, vbCritical + vbOKOnly, "Error"
Call LogError(err)
Unload Me

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

stats(3) = X / Screen.TwipsPerPixelX

    Select Case stats(3)
       Case WM_LBUTTONDOWN
       Case WM_LBUTTONUP
       Case WM_LBUTTONDBLCLK
           Me.Show
       Case WM_RBUTTONDOWN
           'set this form as current window, otherwise popup
           'menu will stay on the screen even after you click
           'outside it
           Call SetForegroundWindow(Me.hwnd)
           PopupMenu frmMain.mnuconfigure
       Case WM_RBUTTONUP
       Case WM_RBUTTONDBLCLK
    End Select
End Sub


Private Sub Form_Unload(Cancel As Integer)
'show stats
'MsgBox "Ads Blocked : " + Str(stats(0)) + vbNewLine + "Cookies Killed : " + Str(stats(1)), vbInformation, "Statistics"

'stop adblocking
blockads (False)

'save current settings
Call SaveSettings

'remove icon from systray
Shell_NotifyIcon NIM_DELETE, nid

End Sub

Private Sub Label3_Click()
'don't forget to send me a mail!
ShellExecute 0, "open", "mailto:abhishekdutta@mail.com", vbNullString, vbNullString, vbshownormal
End Sub

Private Sub mnuabout_Click()
Me.Visible = True
End Sub

Private Sub mnublockads_Click()
mnublockads.Checked = Not mnublockads.Checked
blockads (mnublockads.Checked)
End Sub

Private Sub mnublockpopups_Click()
mnublockpopups.Checked = Not mnublockpopups.Checked
Call LoadFileLst(lstpops, App.Path + "\data\popups.svr")
'block popups
End Sub

Private Sub mnuexit_Click()
Unload Me
End Sub

Private Sub blockads(state As Boolean)
On Error GoTo err:

If state = True Then
    ' if adblocking is enabled, load contents from hosts.svr
    'and save it to system's hosts file
    Call LoadFileLst(lstFile, App.Path + "\data\hosts.svr", "127.0.0.1 ")
    Call SaveFileLst(lstFile, strwindir)
    
    ' if dns caching is enabled, load dns,svr and append to hosts file
    If GetSetting("AdBlocker", "Settings", "UseDNS", "0") = "1" Then
        Call LoadFileLst(lstFile, App.Path + "\data\dns.svr")
        Call AppendFileLst(lstFile, strwindir)
    End If
    
    'start our own ad server!
    Call StartServer
    
Else
    ' restore the contents of the original hosts file
    Call LoadFileLst(lstFile, App.Path + "\data\hosts.bak")
    Call SaveFileLst(lstFile, strwindir)
    Call StopServer
End If

Exit Sub

err:
MsgBox Str(err.Number) + " : " + err.Description, vbCritical + vbOKOnly, "Error"
Call LogError(err)

End Sub

Private Sub mnukillcookies_Click()

On Error GoTo err:

mnukillcookies.Checked = Not mnukillcookies.Checked

'start cookie killing, but before that get names of servers from where
'cookies should not be deleted
If mnukillcookies.Checked = True Then
    If CInt(GetSetting("AdBlocker", "Settings", "All", "0")) = 0 Then
        Call LoadFileLst(lstCk, App.Path + "\data\cookies.svr")
    End If
Else
    'stop cookie killing
    lstCk.Clear
End If

Exit Sub

err:
MsgBox Str(err.Number) + " : " + err.Description, vbCritical + vbOKOnly, "Error"
Call LogError(err)

End Sub

Private Sub mnusettings_Click()
'
Load frmSettings
frmSettings.Show

End Sub

Private Sub Timer_Timer()

class = ""
hnd = 0
hcld = 0
Status = False

On Error GoTo err:

'check active window
If mnublockpopups.Checked = True Then
    hnd = GetForegroundWindow()
    Call GetClassName(hnd, class, 15)
    'if the window is an IE window then
    If Left(class, 7) = "IEFrame" Or Left(class, 13) = "CabinetWClass" Then
        'get its title
        'if parent ie window is visible then
        If IsWindowVisible(hnd) <> False Then
        
            'find first child window
            hcld = FindWindowEx(hnd, GW_CHILD And GW_HWNDFIRST, vbNullString, vbNullString)
            Call GetClassName(hcld, class, 8)
            
            'if WorkerA (ie animation on top right) class is found and it is invisible
            'then the window is a popup
            If Left(class, 7) = "WorkerA" And IsWindowVisible(hcld) = False Then
                Call Sleep(900)
                Dim cap As String * 100
                Call GetWindowText(hnd, cap, 100)
        
                        'check whether should it be closed or not
                If lstpops.ListCount > 0 Then
                    For j = 0 To lstpops.ListCount - 1
                        If InStr(1, cap, lstpops.List(j), vbTextCompare) > 0 Then
                            Status = True 'dont close this window because user doesn't want so
                            Exit For
                        End If
                    Next
                End If
                
                If Status = False Then
                    'close the popup
                    Call PostMessage(hnd, WM_CLOSE, 0, 0)
                    stats(2) = stats(2) + 1
                End If
                
            End If
        End If
    End If
End If

If mnukillcookies.Checked = True Then
    fileCk.Refresh

    'delete first cookie if selective cookie deletion is disabled
    If CInt(GetSetting("AdBlocker", "Settings", "All", "0")) = 1 Then
        Call KillCookie(fileCk.List(0))
    Else
    'find out files names of cookies to be deleted and delete them
        For i = 0 To fileCk.ListCount - 1
            Status = False
    
                For j = 0 To lstCk.ListCount - 1
        
                    If InStr(1, fileCk.List(i), lstCk.List(j), vbTextCompare) > 1 Then
                        Status = True
                        Exit For
                    End If
                Next
    
            If Status = False Then
                Call KillCookie(fileCk.List(i))
            End If
 
        Next
    End If
End If

nid.szTip = "Ads:" + Str(stats(0)) + " Cookies:" + Str(stats(1)) + " Popups:" + Str(stats(2)) + vbNullChar
Shell_NotifyIcon NIM_MODIFY, nid

Exit Sub

err:
Call LogError(err)
Exit Sub

End Sub

Private Sub wsock_ConnectionRequest(Index As Integer, ByVal requestID As Long)
On Error GoTo err

If Index = 0 Then
    connections = connections + 1
    Load wsock(connections)
    wsock(connections).LocalPort = 0
    wsock(connections).Accept requestID
End If
Exit Sub

err:
MsgBox Str(err.Number) + " : " + err.Description, vbCritical + vbOKOnly, "Error"
Call LogError(err)

End Sub

Private Sub wsock_DataArrival(Index As Integer, ByVal bytesTotal As Long)

On Error GoTo err

Dim strdata As String
Dim Page As String

strdata = ""

' if the browser request for any ads, send it an appropriate image
wsock(Index).GetData strdata$

If Mid$(strdata$, 1, 3) = "GET" Then
    
    Page = Trim(GetSetting("AdBlocker", "Settings", "Color", "0"))
    Page = App.Path + "\images\" + Page + ".gif"
    
    'store inforamtion in a log file
    Call CreateBlankFile(App.Path + "\ads.log")
    strdata = Left(strdata, (InStr(1, strdata, "HTTP") - 2))
    strdata = Mid(strdata, 4)
    lstFile.AddItem strdata + " send on " + Str(Now())
    Call AppendFileLst(lstFile, App.Path + "\ads.log")
    
    SendPage Page, Index, StrConv(strdata, vbLowerCase)
    stats(0) = stats(0) + 1
Else
    wsock(Index).Close
End If
Exit Sub

err:

'MsgBox Str(err.Number) + " : " + err.Description, vbCritical + vbOKOnly, "Error"
Call LogError(err)

End Sub

Private Sub wsock_SendComplete(Index As Integer)
wsock(Index).Close
Call Sleep(100)
Call DeleteUrlCacheEntry("http://127.0.0.1/bg.gif")
End Sub

Private Sub StartServer()
On Error GoTo err

connections = 1
Me.wsock(0).Close
Me.wsock(0).LocalPort = 80
Me.wsock(0).Listen
Exit Sub

err:
If err.Number = 10048 Then
    MsgBox "A webserver is already running on port 80, for best results please stop the server and restart this application", vbOKOnly, "Error"
    Form_Unload (False)
    End
Else
    MsgBox Str(err.Number) + " : " + err.Description, vbCritical + vbOKOnly, "Error"
End If
Call LogError(err)
End Sub

Private Sub StopServer()
'not required
End Sub

Private Sub SaveSettings()
On Error GoTo err

Call SaveSetting("AdBlocker", "Data", "AdBlock", mnublockads.Checked)
Call SaveSetting("AdBlocker", "Data", "Cookies", mnukillcookies.Checked)
Call SaveSetting("AdBlocker", "Data", "Popup", mnublockpopups.Checked)
Call SaveSetting("AdBlocker", "Data", "Restore", "True")

Call SetAttr(App.Path + "\data\hosts.svr", vbReadOnly + vbSystem)
Call SetAttr(App.Path + "\data\cookies.svr", vbReadOnly + vbSystem)
Call SetAttr(App.Path + "\data\dns.svr", vbReadOnly + vbSystem)
Call SetAttr(App.Path + "\data\popups.svr", vbReadOnly + vbSystem)
Call SetAttr(strwindir, vbReadOnly)

Exit Sub

err:
MsgBox Str(err.Number) + " : " + err.Description, vbCritical + vbOKOnly, "Error"
Call LogError(err)

End Sub

Private Sub InitSettings()
On Error GoTo err

If GetSetting("AdBlocker", "Data", "AdBlock", False) = True Then Call mnublockads_Click
If GetSetting("AdBlocker", "Data", "Cookies", False) = True Then Call mnukillcookies_Click
If GetSetting("AdBlocker", "Data", "Popup", False) = True Then Call mnublockpopups_Click

Exit Sub

err:
MsgBox Str(err.Number) + " : " + err.Description, vbCritical + vbOKOnly, "Error"
Call LogError(err)

End Sub

Private Sub KillCookie(CookieName)

If DeleteFile(fileCk.Path + "\" + CookieName) = 1 Then
            'remove the cookie cache entry
            Call DeleteUrlCacheEntry("Cookie:" + Left(CookieName, (Len(CookieName) - 7)) + ".com" + "/")
            
            'save info in log file
            Call CreateBlankFile(App.Path + "\cookies.log")
            lstFile.AddItem (fileCk.Path + "\" + CookieName + " deleted on " + Str(Now()))
            Call AppendFileLst(lstFile, App.Path + "\cookies.log")
            
            stats(1) = stats(1) + 1
                        
End If


End Sub

Private Sub SetAttributes()
On Error GoTo err

Call SetAttr(App.Path + "\data\hosts.svr", vbNormal)
Call SetAttr(App.Path + "\data\cookies.svr", vbNormal)
Call SetAttr(App.Path + "\data\dns.svr", vbNormal)
Call SetAttr(App.Path + "\data\popups.svr", vbNormal)
If Dir(strwindir, vbHidden Or vbSystem Or vbReadOnly) <> "" Then Call SetAttr(strwindir, vbNormal)
Exit Sub

err:

MsgBox Str(err.Number) + " : " + err.Description, vbCritical + vbOKOnly, "Error"
Call LogError(err)

End Sub
