VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.DLL"
Begin VB.Form Form1 
   Caption         =   "Power Browser"
   ClientHeight    =   7395
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   10755
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7395
   ScaleWidth      =   10755
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "Refresh"
      Height          =   255
      Left            =   8160
      TabIndex        =   7
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "Stop"
      Height          =   255
      Left            =   7200
      TabIndex        =   6
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton cmdForward 
      Caption         =   "Forward"
      Height          =   255
      Left            =   6240
      TabIndex        =   5
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back"
      Height          =   255
      Left            =   5280
      TabIndex        =   4
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton cmdGo 
      Caption         =   "GO!"
      Default         =   -1  'True
      Height          =   255
      Left            =   4320
      TabIndex        =   1
      Top             =   120
      Width           =   855
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   4335
      Left            =   1320
      TabIndex        =   3
      Top             =   840
      Width           =   6735
      ExtentX         =   11880
      ExtentY         =   7646
      ViewMode        =   1
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   -1  'True
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin VB.TextBox txtAddress 
      Height          =   285
      Left            =   0
      TabIndex        =   0
      Text            =   "Http://"
      Top             =   120
      Width           =   4335
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   9360
      Picture         =   "Form1.frx":0442
      Top             =   0
      Width           =   480
   End
   Begin VB.Label lblStatus 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Power Browser - Blocking Pop-Up windows | Coded by:Dustin Davis - Bootleg Software Inc."
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   480
      Width           =   7815
   End
   Begin VB.Menu nmuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuexit 
         Caption         =   "Exit"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "&Options"
      Begin VB.Menu mnuOptionsAllow 
         Caption         =   "Allow Pop-Up windows"
      End
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "&About Power Browser"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***********************************************************************
'Coded by: Dustin Davis
'Bootleg Software Inc.
'http://www.warpnet.org/bsi
'
'
'Power Browser shows you how to use the web browser control and how to keep pop
'up windows from appearing!!
'If you use this in your program, please give me credit
'
'PLEASE DO NOT STEAL THIS PROGRAM, I'VE ALREADY RELEASED IT TO THE PUBLIC ON
'MY SITE!! http://www.warpnet.org/bsi
'
'***************************************************************************

Public AllowPopup As Boolean 'This is for Pop-up windows

Private Sub cmdBack_Click()
'Go back one page
WebBrowser1.GoBack
End Sub

Private Sub cmdForward_Click()
'go forward one page
WebBrowser1.GoForward
End Sub

Private Sub cmdGo_Click()
'Go to web page
WebBrowser1.Navigate txtAddress.Text
lblStatus.Caption = "Going to: " & txtAddress.Text
End Sub

Private Sub cmdRefresh_Click()
'Refresh page
WebBrowser1.Refresh
End Sub

Private Sub cmdStop_Click()
'Stop loading
WebBrowser1.Stop
End Sub

Private Sub Form_Load()
'Resize and place objects
With WebBrowser1
    .Width = Form1.Width - 200
    .Left = 50
    .Height = Form1.Height - 200
End With
With lblStatus
    .Top = txtAddress.Top + txtAddress.Height + 50
    .Left = WebBrowser1.Left
    .FontBold = True
    .Width = WebBrowser1.Width
End With
txtAddress.Left = 50
cmdGo.Left = (txtAddress.Left + txtAddress.Width) + 20
End Sub

Private Sub Form_Resize()
'Resizes everything to fit to the form
With WebBrowser1
    .Width = Form1.Width - 200
    .Left = 50
    .Height = Form1.Height - 1500
End With
With lblStatus
    .Top = txtAddress.Top + txtAddress.Height + 50
    .Left = WebBrowser1.Left
    .Width = WebBrowser1.Width
End With
End Sub

Private Sub mnuAbout_Click()
MsgBox "Power Browser" & vbCrLf & "Coded by: Dustin Davis" & vbCrLf & "http://www.warpnet.org/bsi", vbOKOnly, "About Power Browser"
WebBrowser1.Navigate "http://www.warpnet.org/bsi"
End Sub

Private Sub mnuexit_Click()
'Exit program
Unload Me
End Sub

Private Sub mnuOptionsAllow_Click()
'Turn on/off pop-up windows
If AllowPopup = True Then
    AllowPopup = False
    mnuOptionsAllow.Checked = False
    lblStatus.Caption = "Power Browser - Blocking Pop-Up windows | Coded by:Dustin Davis - Bootleg Software Inc."
ElseIf AllowPopup = False Then
    AllowPopup = True
    mnuOptionsAllow.Checked = True
    lblStatus.Caption = "Power Browser - Allowing Pop-Up windows | Coded by:Dustin Davis - Bootleg Software Inc."
End If
End Sub

Private Sub WebBrowser1_DocumentComplete(ByVal pDisp As Object, URL As Variant)
'shows done in the status bar
lblStatus.Caption = "Done Loading"
Form1.Caption = "Power Browser - " & WebBrowser1.LocationName
End Sub

Private Sub WebBrowser1_DownloadBegin()
'Starting download
lblStatus.Caption = "Starting Download"
End Sub

Private Sub WebBrowser1_DownloadComplete()
'Done downloading
lblStatus.Caption = "Download Done!"
End Sub

Private Sub WebBrowser1_NavigateComplete2(ByVal pDisp As Object, URL As Variant)
'Loaded page
lblStatus.Caption = "Done Loading!"
Form1.Caption = "Power Browser - " & WebBrowser1.LocationName  'Shows webpage in title bar
End Sub

Private Sub WebBrowser1_NewWindow2(ppDisp As Object, Cancel As Boolean)
'This will allow a pop-up window to load or to be blocked!
If AllowPopup = True Then
    Cancel = False
    DoEvents
ElseIf AllowPopup = False Then
    Cancel = True
End If
End Sub

Private Sub WebBrowser1_ProgressChange(ByVal Progress As Long, ByVal ProgressMax As Long)
'Shows progress in status bar
lblStatus.Caption = "Reading " & Progress & "  of  " & ProgressMax
End Sub

Private Sub WebBrowser1_StatusTextChange(ByVal Text As String)
'shows new text in status bar
lblStatus.Caption = Text
End Sub

Function FileExist(vFile As String) As Boolean
    On Error Resume Next
    FileExist = False
    If Dir$(vFile) <> "" Then: FileExist = True
End Function
