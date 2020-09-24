VERSION 5.00
Begin VB.Form main 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Fast Download"
   ClientHeight    =   2775
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7215
   Icon            =   "main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2775
   ScaleWidth      =   7215
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox af 
      Caption         =   "After download open destination folder"
      Height          =   195
      Left            =   3000
      TabIndex        =   13
      Top             =   2160
      Width           =   3015
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   6720
      Top             =   1920
   End
   Begin VB.CheckBox ac 
      Caption         =   "After download close this window"
      Height          =   195
      Left            =   120
      TabIndex        =   12
      Top             =   2160
      Value           =   1  'Checked
      Width           =   2655
   End
   Begin filedownloader.XP_ProgressBar pb 
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   2400
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   450
   End
   Begin filedownloader.CandyButton cancel 
      Height          =   255
      Left            =   6120
      TabIndex        =   10
      Top             =   2400
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   450
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Cancel"
      Checked         =   0   'False
      ColorButtonHover=   16760976
      ColorButtonUp   =   15309136
      ColorButtonDown =   15309136
      BorderBrightness=   0
      ColorBright     =   16772528
      DisplayHand     =   0   'False
      ColorScheme     =   0
   End
   Begin filedownloader.CandyButton ok 
      Height          =   255
      Left            =   6120
      TabIndex        =   9
      Top             =   120
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   450
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Download"
      Checked         =   0   'False
      ColorButtonHover=   16760976
      ColorButtonUp   =   15309136
      ColorButtonDown =   15309136
      BorderBrightness=   0
      ColorBright     =   16772528
      DisplayHand     =   0   'False
      ColorScheme     =   0
   End
   Begin VB.Frame Frame1 
      Caption         =   "Download Info"
      Height          =   1575
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   6975
      Begin VB.Label starttime2 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   255
         Left            =   3120
         TabIndex        =   8
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label folder 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   615
         Left            =   1800
         TabIndex        =   7
         Top             =   840
         Width           =   5055
      End
      Begin VB.Label starttime 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   255
         Left            =   1800
         TabIndex        =   6
         Top             =   600
         Width           =   975
      End
      Begin VB.Label nume 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   255
         Left            =   1200
         TabIndex        =   5
         Top             =   360
         Width           =   5655
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Destination Folder :"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Download time :"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "File name :"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.TextBox url 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Text            =   "http://replica.wen.ru/vzbss.rar"
      Top             =   120
      Width           =   5895
   End
   Begin filedownloader.dl dl 
      Left            =   2880
      Top             =   2880
      _ExtentX        =   1799
      _ExtentY        =   1667
   End
End
Attribute VB_Name = "main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cancel_Click()
dl.cancel
pb.Value = 0
Me.Caption = "Fast Download"
ok.Enabled = True
url.Enabled = True
cancel.Enabled = False
End Sub

Private Sub dl_DowloadComplete()
Dim shll32 As New Shell
pb.Value = 0
GetFileName (url.Text)
starttime2.Caption = Time
ok.Enabled = True
url.Enabled = True
cancel.Enabled = False
pb.Value = 0
GetFileName (url.Text)
starttime2.Caption = Time
ok.Enabled = True
url.Enabled = True
cancel.Enabled = False
If ac.Value = 1 Then
Timer1.Enabled = True
Else
Timer1.Enabled = False
End If
If af.Value = 1 Then
shll32.Explore folder.Caption
Else
End If
End Sub

Private Sub dl_DownloadErrors(strError As String)
MsgBox "Connection error", vbCritical, "Not downloaded"
ok.Enabled = True
url.Enabled = True
cancel.Enabled = False
End Sub

Private Sub dl_DownloadEvents(strEvent As String)
starttime.Caption = Time
ok.Enabled = False
url.Enabled = False
cancel.Enabled = True
nume.Caption = GetFileName(url.Text)
End Sub

Private Sub dl_DownloadProgress(intPercent As String)
pb.Value = intPercent
Me.Caption = "Fast Download " & pb.Value & " %"
GetFileName (url.Text)
cancel.Enabled = True
url.Text = ""
End Sub

Private Sub Form_Load()
ok.Enabled = True
url.Enabled = True
cancel.Enabled = False
End Sub

Private Sub Form_Unload(cancel As Integer)
dl.cancel
End
End Sub

Private Sub ok_Click()
    Dim shll32 As New Shell
    Dim objFolder As Object
    Dim objItem As FolderItem
    Const BIF_NEWDIALOGSTYLE = &H40
    On Error GoTo errHand
Set objFolder = shll32.BrowseForFolder(Me.hwnd, _
                                         "Select a Folder", _
                                         BIF_NEWDIALOGSTYLE)
Set objItem = objFolder.Self
folder.Caption = objItem.Path
dl.DownloadFile main.url.Text, objItem.Path & "\" & GetFileName(main.url.Text)
On Error GoTo 0                     ' Turn off error handling.

Exit Sub
errHand:
End Sub

Private Sub Timer1_Timer()
dl.cancel
End
End Sub
