VERSION 5.00
Object = "{22D6F304-B0F6-11D0-94AB-0080C74C7E95}#1.0#0"; "msdxm.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmAudioPlayer 
   Appearance      =   0  'Flat
   BackColor       =   &H80000018&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ADMP3 2003"
   ClientHeight    =   6150
   ClientLeft      =   3000
   ClientTop       =   2340
   ClientWidth     =   11505
   Icon            =   "frmAudioPlayer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6150
   ScaleWidth      =   11505
   Begin Project1.lvButtons_H lvButtons_Close 
      Height          =   615
      Left            =   6480
      TabIndex        =   23
      Top             =   840
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   1085
      Caption         =   "Close"
      CapAlign        =   2
      BackStyle       =   2
      Shape           =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cBhover         =   12640511
      LockHover       =   1
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      ImgAlign        =   5
      Image           =   "frmAudioPlayer.frx":0442
      cBack           =   16761024
   End
   Begin Project1.lvButtons_H lvButtons_Pause 
      Height          =   615
      Left            =   6000
      TabIndex        =   22
      Top             =   840
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   1085
      Caption         =   "Pause"
      CapAlign        =   2
      BackStyle       =   2
      Shape           =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cBhover         =   12640511
      LockHover       =   1
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      ImgAlign        =   4
      Image           =   "frmAudioPlayer.frx":0894
      cBack           =   16761024
   End
   Begin Project1.lvButtons_H lvButtons_Stop 
      Height          =   615
      Left            =   6480
      TabIndex        =   21
      Top             =   120
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   1085
      Caption         =   "Stop"
      CapAlign        =   2
      BackStyle       =   2
      Shape           =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cBhover         =   12640511
      LockHover       =   1
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      ImgAlign        =   5
      Image           =   "frmAudioPlayer.frx":0CE6
      cBack           =   16761024
   End
   Begin Project1.lvButtons_H lvButtons_Play 
      Height          =   615
      Left            =   6000
      TabIndex        =   20
      Top             =   120
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   1085
      Caption         =   "Play"
      CapAlign        =   2
      BackStyle       =   2
      Shape           =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFHover         =   0
      cBhover         =   12640511
      LockHover       =   3
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      ImgAlign        =   4
      Image           =   "frmAudioPlayer.frx":1138
      cBack           =   16761024
   End
   Begin VB.Frame Frame4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      Caption         =   "Volume"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1215
      Left            =   2160
      TabIndex        =   17
      Top             =   120
      Width           =   3735
      Begin Project1.lvButtons_H lvButtons_VolDecrease 
         Height          =   375
         Left            =   120
         TabIndex        =   25
         Top             =   720
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cBhover         =   12640511
         LockHover       =   1
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         Image           =   "frmAudioPlayer.frx":158A
         cBack           =   16761024
      End
      Begin Project1.lvButtons_H lvButtons_VolIncrease 
         Height          =   375
         Left            =   3240
         TabIndex        =   24
         Top             =   720
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cBhover         =   12640511
         LockHover       =   1
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         Image           =   "frmAudioPlayer.frx":19DC
         cBack           =   16761024
      End
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   375
         Left            =   600
         TabIndex        =   18
         Top             =   720
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   661
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.Image Image2 
         Height          =   360
         Left            =   3240
         Picture         =   "frmAudioPlayer.frx":1E2E
         Stretch         =   -1  'True
         Top             =   240
         Width           =   360
      End
      Begin VB.Image Image1 
         Height          =   360
         Left            =   120
         Picture         =   "frmAudioPlayer.frx":2278
         Stretch         =   -1  'True
         Top             =   240
         Width           =   360
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00404000&
         BeginProperty Font 
            Name            =   "Moderne"
            Size            =   21.75
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   600
         TabIndex        =   19
         Top             =   240
         Width           =   2535
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   6000
      Top             =   240
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   240
      TabIndex        =   6
      Top             =   4080
      Width           =   1455
   End
   Begin VB.FileListBox File1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1710
      Left            =   4440
      Pattern         =   "*.mp3;*.cda;*.mid;*.wav;*.avi;*.dat"
      TabIndex        =   8
      Top             =   4080
      Width           =   3015
   End
   Begin VB.DirListBox Dir1 
      Appearance      =   0  'Flat
      Height          =   1665
      Left            =   1800
      TabIndex        =   7
      Top             =   4080
      Width           =   2535
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1320
      TabIndex        =   1
      Top             =   1680
      Width           =   6135
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      Height          =   1200
      Left            =   240
      Sorted          =   -1  'True
      TabIndex        =   2
      Top             =   2040
      Width           =   7215
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      Caption         =   "File List"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2175
      Left            =   120
      TabIndex        =   9
      Top             =   3840
      Width           =   7455
      Begin Project1.lvButtons_H lvButtons_Command2 
         Height          =   615
         Left            =   120
         TabIndex        =   27
         Top             =   1440
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   1085
         Caption         =   "File To Play List"
         CapAlign        =   2
         BackStyle       =   7
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cBhover         =   12640511
         LockHover       =   1
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         ImgAlign        =   2
         Image           =   "frmAudioPlayer.frx":26C2
         ImgSize         =   32
         cBack           =   8438015
      End
      Begin Project1.lvButtons_H lvButtons_Command1 
         Height          =   615
         Left            =   120
         TabIndex        =   26
         Top             =   720
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   1085
         Caption         =   "Play File"
         CapAlign        =   2
         BackStyle       =   7
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cBhover         =   12640511
         LockHover       =   1
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         ImgAlign        =   2
         Image           =   "frmAudioPlayer.frx":2B14
         ImgSize         =   32
         cBack           =   8438015
      End
      Begin VB.Line Line1 
         BorderColor     =   &H000000FF&
         X1              =   1680
         X2              =   4200
         Y1              =   2040
         Y2              =   2040
      End
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      Caption         =   "Play List"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2295
      Left            =   120
      TabIndex        =   10
      Top             =   1440
      Width           =   7455
      Begin Project1.lvButtons_H lvButtons_Remove 
         Height          =   255
         Left            =   5760
         TabIndex        =   29
         ToolTipText     =   "Remove From Play List"
         Top             =   1920
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   450
         Caption         =   "REM"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cBhover         =   12640511
         LockHover       =   1
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         cBack           =   16761024
      End
      Begin Project1.lvButtons_H lvButtons_Empty 
         Height          =   255
         Left            =   6600
         TabIndex        =   28
         ToolTipText     =   "Empty Play List"
         Top             =   1920
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   450
         Caption         =   "EMPTY"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cBhover         =   12640511
         LockHover       =   1
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         cBack           =   16761024
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H80000018&
         Caption         =   "Single Play"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   1920
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.OptionButton opCont 
         BackColor       =   &H80000018&
         Caption         =   "Continuous Play"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3360
         TabIndex        =   5
         Top             =   1920
         Width           =   1695
      End
      Begin VB.OptionButton opRepeat 
         BackColor       =   &H80000018&
         Caption         =   "Repeat Play"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1680
         TabIndex        =   4
         Top             =   1920
         Width           =   1455
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000018&
         Caption         =   "Selected:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame Frame3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      Caption         =   "Now Playing"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1215
      Left            =   120
      TabIndex        =   12
      Top             =   120
      Width           =   1935
      Begin VB.Label lblMin 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Moderne"
            Size            =   48
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   1095
         Left            =   120
         TabIndex        =   16
         Top             =   120
         Width           =   495
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "Moderne"
            Size            =   48
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   1095
         Left            =   600
         TabIndex        =   15
         Top             =   0
         Width           =   255
      End
      Begin VB.Label lblSec2 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Moderne"
            Size            =   48
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   1095
         Left            =   840
         TabIndex        =   14
         Top             =   120
         Width           =   495
      End
      Begin VB.Label lblSec 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Moderne"
            Size            =   48
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   1095
         Left            =   1320
         TabIndex        =   13
         Top             =   120
         Width           =   495
      End
   End
   Begin MediaPlayerCtl.MediaPlayer MediaPlayer1 
      CausesValidation=   0   'False
      Height          =   5895
      Left            =   7680
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   120
      Visible         =   0   'False
      Width           =   3690
      AudioStream     =   -1
      AutoSize        =   0   'False
      AutoStart       =   -1  'True
      AnimationAtStart=   0   'False
      AllowScan       =   0   'False
      AllowChangeDisplaySize=   0   'False
      AutoRewind      =   0   'False
      Balance         =   0
      BaseURL         =   ""
      BufferingTime   =   5
      CaptioningID    =   ""
      ClickToPlay     =   0   'False
      CursorType      =   0
      CurrentPosition =   -1
      CurrentMarker   =   0
      DefaultFrame    =   ""
      DisplayBackColor=   0
      DisplayForeColor=   16777215
      DisplayMode     =   0
      DisplaySize     =   4
      Enabled         =   -1  'True
      EnableContextMenu=   0   'False
      EnablePositionControls=   0   'False
      EnableFullScreenControls=   0   'False
      EnableTracker   =   -1  'True
      Filename        =   ""
      InvokeURLs      =   0   'False
      Language        =   -1
      Mute            =   0   'False
      PlayCount       =   1
      PreviewMode     =   0   'False
      Rate            =   1
      SAMILang        =   ""
      SAMIStyle       =   ""
      SAMIFileName    =   ""
      SelectionStart  =   -1
      SelectionEnd    =   -1
      SendOpenStateChangeEvents=   0   'False
      SendWarningEvents=   0   'False
      SendErrorEvents =   0   'False
      SendKeyboardEvents=   0   'False
      SendMouseClickEvents=   0   'False
      SendMouseMoveEvents=   0   'False
      SendPlayStateChangeEvents=   -1  'True
      ShowCaptioning  =   0   'False
      ShowControls    =   -1  'True
      ShowAudioControls=   0   'False
      ShowDisplay     =   0   'False
      ShowGotoBar     =   0   'False
      ShowPositionControls=   0   'False
      ShowStatusBar   =   -1  'True
      ShowTracker     =   -1  'True
      TransparentAtStart=   0   'False
      VideoBorderWidth=   1
      VideoBorderColor=   0
      VideoBorder3D   =   0   'False
      Volume          =   -1500
      WindowlessVideo =   0   'False
   End
End
Attribute VB_Name = "frmAudioPlayer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'##############################################
'#          Coded by Adel D. Mendoza          #
'#        Designed by Adel D. Mendoza         #
'#            AUDIO / MP3 Player              #
'#                                            #
'#        area :  frmAudioPlayer              #
'# description :  Code File Audio/Mp3 Player  #
'#     E-mail  :  adm@rfm.com.ph              #
'#     url     :  http://www.rfm.com.ph       #
'#                                            #
'#         Special Thanks to LaVolpe          #
'#              for the Buttons               #
'##############################################

Dim lblAudio
Dim lblVol
Dim allow_play As String
Dim paused As Boolean
Dim allow_pause As Boolean

Private Sub lvButtons_Close_Click()
   MediaPlayer1.Visible = False
   Unload Me
End Sub

Private Sub lvButtons_Remove_Click()
   Dim sTemp As String
   If List1.ListCount <> 0 Then
      If Text1.Text <> "" Then
         List1.RemoveItem (List1.ListIndex)
         Text1.Text = ""
         If FileExists("c:PlayList.txt") Then
            Kill ("c:/PlayList.txt")
         End If
         Open "c:/PlayList.txt" For Output As #1
         Close #1
         Open "c:/PlayList.txt" For Append As #1
         For I = 0 To List1.ListCount
             sTemp = List1.List(I)
             If sTemp <> "" Then
                Print #1, List1.List(I)
             End If
         Next
         Close #1
      End If
   End If
End Sub

Private Sub lvButtons_Empty_Click()
   List1.Clear
   Text1.Text = ""
   If FileExists("c:/PlayList.txt") Then
      Kill ("c:/PlayList.txt")
   End If
End Sub

Private Sub lvButtons_Pause_Click()
   If allow_pause = True Then
      On Error Resume Next
      If paused = False Then
         MediaPlayer1.Pause
         paused = True
         allow_play = "no"
         Exit Sub
      End If
   End If
End Sub

Private Sub lvButtons_Play_Click()
   lblAudio = 0
   If Text1.Text <> "" Then
      If UCase(Mid(Text1.Text, Len(Text1.Text) - 2, 3)) = "DAT" Or _
         UCase(Mid(Text1.Text, Len(Text1.Text) - 2, 3)) = "AVI" Then
         Me.Width = 11580
         MediaPlayer1.Visible = True
         lblAudio = 1
      End If
   Else
      allow_play = "no"
   End If
   
   If paused = True Then
      MediaPlayer1.Play
      paused = False
      allow_play = "yes"
      Exit Sub
   End If
   If paused = False Then
      MediaPlayer1.Open Text1.Text
      lblSec = "0"
      lblMin = "0"
      lblSec2 = "0"
      Exit Sub
   End If
End Sub

Private Sub lvButtons_Stop_Click()
   MediaPlayer1.Stop
   allow_play = "no"
   lblSec = "0"
   lblMin = "0"
   lblSec2 = "0"
   allow_pause = False
   MediaPlayer1.Visible = False
   Me.Width = 7785
End Sub

Private Sub lvButtons_VolDecrease_Click()
   If lblVol > 0 Then
      lblVol = lblVol - 5
   End If
   Select Case lblVol
   Case "100"
        MediaPlayer1.Volume = 0
   Case "95"
        MediaPlayer1.Volume = -300
   Case "90"
        MediaPlayer1.Volume = -600
   Case "85"
        MediaPlayer1.Volume = -900
   Case "80"
        MediaPlayer1.Volume = -1200
   Case "75"
        MediaPlayer1.Volume = -1500
   Case "70"
        MediaPlayer1.Volume = -1800
   Case "65"
        MediaPlayer1.Volume = -2100
   Case "60"
        MediaPlayer1.Volume = -2400
   Case "55"
        MediaPlayer1.Volume = -2700
   Case "50"
        MediaPlayer1.Volume = -3000
   Case "45"
        MediaPlayer1.Volume = -3300
   Case "40"
        MediaPlayer1.Volume = -3600
   Case "35"
        MediaPlayer1.Volume = -3900
   Case "30"
        MediaPlayer1.Volume = -4200
   Case "25"
        MediaPlayer1.Volume = -4500
   Case "20"
        MediaPlayer1.Volume = -4800
   Case "15"
        MediaPlayer1.Volume = -5100
   Case "10"
        MediaPlayer1.Volume = -5400
   Case "5"
        MediaPlayer1.Volume = -5700
   Case "0"
        MediaPlayer1.Volume = -6000
   End Select
   ProgressBar1.Value = lblVol
   Label2.Caption = Str(lblVol) + "%"
   Exit Sub
End Sub

Private Sub lvButtons_VolIncrease_Click()
   If lblVol < 100 Then
      lblVol = lblVol + 5
   End If
   Select Case lblVol
   Case "100"
        MediaPlayer1.Volume = 0
   Case "95"
        MediaPlayer1.Volume = -300
   Case "90"
        MediaPlayer1.Volume = -600
   Case "85"
        MediaPlayer1.Volume = -900
   Case "80"
        MediaPlayer1.Volume = -1200
   Case "75"
        MediaPlayer1.Volume = -1500
   Case "70"
        MediaPlayer1.Volume = -1800
   Case "65"
        MediaPlayer1.Volume = -2100
   Case "60"
        MediaPlayer1.Volume = -2400
   Case "55"
        MediaPlayer1.Volume = -2700
   Case "50"
        MediaPlayer1.Volume = -3000
   Case "45"
        MediaPlayer1.Volume = -3300
   Case "40"
        MediaPlayer1.Volume = -3600
   Case "35"
        MediaPlayer1.Volume = -3900
   Case "30"
        MediaPlayer1.Volume = -4200
   Case "25"
        MediaPlayer1.Volume = -4500
   Case "20"
        MediaPlayer1.Volume = -4800
   Case "15"
        MediaPlayer1.Volume = -5100
   Case "10"
        MediaPlayer1.Volume = -5400
   Case "5"
        MediaPlayer1.Volume = -5700
   Case "0"
        MediaPlayer1.Volume = -6000
   End Select
   ProgressBar1.Value = lblVol
   Label2.Caption = Str(lblVol) + "%"
   Exit Sub
End Sub

Private Sub lvButtons_Command1_Click()
   If File1.FileName <> "" Then
      MediaPlayer1.Visible = False
      lblAudio = 0
      If UCase(Mid(File1.FileName, Len(File1.FileName) - 2, 3)) = "DAT" Or _
         UCase(Mid(File1.FileName, Len(File1.FileName) - 2, 3)) = "AVI" Then
         Me.Width = 11580
         MediaPlayer1.Visible = True
         lblAudio = 1
      End If
      MediaPlayer1.Open File1.Path & "\" & File1.FileName
   End If
End Sub

Private Sub lvButtons_Command2_Click()
   If File1.FileName <> "" Then
      oldsongs = ""
      List1.AddItem File1.Path & "\" & File1.FileName
      newsong = File1.Path & "\" & File1.FileName
      On Error Resume Next
      Open "c:/PlayList.txt" For Append As #1
      Print #1, "" & newsong & ""
      Close #1
   End If
End Sub

Private Sub Dir1_Change()
   File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
   Dir1.Path = Drive1.Drive
End Sub

Private Sub Form_Load()
   Me.Top = 0
   Me.Left = 0
   Me.Width = 7785
   lblAudio = 0
   
   allow_pause = False
   Dir1.Path = Drive1.Drive
   paused = False
   On Error Resume Next
   If Not FileExists("c:/PlayList.txt") Then
      Open "c:/PlayList.txt" For Output As #1
      Close #1
   End If
   Open "c:/PlayList.txt" For Input As #1
   Do Until EOF(1)
      Input #1, playlistitem
      List1.AddItem playlistitem
   Loop
   Close #1
   allow_play = "no"
   '--------------------------------------
   'set volume setting for the mediaplayer
   '--------------------------------------
   lblVol = 50
   MediaPlayer1.Volume = -3000
   ProgressBar1.Value = lblVol
   Label2.Caption = Str(lblVol) + "%"
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Me.Width = 7785
  MediaPlayer1.Visible = False
  End
End Sub

Private Sub List1_Click()
   Text1.Text = List1.Text
End Sub

Private Sub List1_DblClick()
   lblAudio = 0
   
   Text1.Text = List1.Text
   MediaPlayer1.Stop
   MediaPlayer1.Visible = False
   If Text1.Text <> "" Then
      If UCase(Mid(Text1.Text, Len(Text1.Text) - 2, 3)) = "DAT" Or _
         UCase(Mid(Text1.Text, Len(Text1.Text) - 2, 3)) = "AVI" Then
         Me.Width = 11580
         MediaPlayer1.Visible = True
         lblAudio = 1
      End If
   Else
      allow_play = "no"
   End If

   lblSec = "0"
   lblMin = "0"
   lblSec2 = "0"
   MediaPlayer1.Open Text1.Text
End Sub

Private Sub List1_KeyPress(KeyAscii As Integer)
   Text1.Text = List1.Text
End Sub

Private Sub MediaPlayer1_EndOfStream(ByVal Result As Long)
   allow_play = "no"
   lblSec = "0"
   lblSec2 = "0"
   lblMin = "0"
   allow_pause = False
   If opCont.Value = True Then
      On Error GoTo error1
      allow_pause = True
      List1.ListIndex = List1.ListIndex + 1
      MediaPlayer1.Open Text1.Text
   Exit Sub
   
error1:
  If Text1.Text <> "" Then
     If List1.ListCount <> 0 Then
        List1.ListIndex = 0
     End If
     MediaPlayer1.Open Text1.Text
  Else
     MediaPlayer1.Stop
  End If
End If

If opRepeat.Value = True Then
   If Text1.Text <> "" Then
      allow_pause = True
      MediaPlayer1.Open Text1.Text
   Else
      MediaPlayer1.Stop
   End If
End If
End Sub

Private Sub MediaPlayer1_NewStream()
   lblSec = "0"
   lblSec2 = "0"
   lblMin = "0"
   allow_play = "yes"
   allow_pause = True
End Sub

Private Sub MediaPlayer1_PlayStateChange(ByVal OldState As Long, ByVal NewState As Long)
   If NewState = 0 Then
      Me.Width = 7785
      MediaPlayer1.Visible = False
   End If
End Sub

Private Sub Timer1_Timer()
   If lblAudio = 0 Then
      If allow_play = "yes" Then
         If lblSec = "9" Then
            lblSec = "0"
           lblSec2 = lblSec2 + 1
           If lblSec2 = "6" Then
              lblSec2 = "0"
              lblMin = lblMin + 1
           End If
         Else
           lblSec = lblSec + 1
         End If
      End If
   End If
End Sub

Private Function FileExists(FullFileName As String) As Boolean
   On Error GoTo MakeF
   'If file does Not exist, there will be an Error
   Open FullFileName For Input As #1
   Close #1
   'no error, file exists
   FileExists = True
   Exit Function
   
MakeF:
   'error, file does Not exist
   FileExists = False
   Exit Function
End Function



