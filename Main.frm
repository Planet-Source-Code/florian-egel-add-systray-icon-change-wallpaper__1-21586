VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "FLOMIX Studios - Wallpaper Generator"
   ClientHeight    =   4590
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   6495
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4590
   ScaleWidth      =   6495
   StartUpPosition =   3  'Windows-Standard
   Begin FDesktop.ShellIcon ShellIcon1 
      Left            =   120
      Top             =   1680
      _ExtentX        =   847
      _ExtentY        =   847
      Icon            =   "Main.frx":27A2
      Visible         =   -1  'True
   End
   Begin VB.PictureBox picPreview 
      AutoRedraw      =   -1  'True
      Height          =   1860
      Left            =   3720
      ScaleHeight     =   120
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   160
      TabIndex        =   13
      Top             =   2280
      Width           =   2460
   End
   Begin VB.PictureBox picFlomix 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   960
      Left            =   8280
      Picture         =   "Main.frx":28FC
      ScaleHeight     =   60
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   80
      TabIndex        =   12
      Top             =   2040
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.PictureBox picScreen 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'Kein
      Height          =   615
      Left            =   8400
      ScaleHeight     =   41
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   49
      TabIndex        =   11
      Top             =   3240
      Visible         =   0   'False
      Width           =   735
   End
   Begin FDesktop.StatusBar StatusBar1 
      Align           =   2  'Unten ausrichten
      Height          =   255
      Left            =   0
      TabIndex        =   8
      Top             =   4335
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   450
   End
   Begin VB.Frame Frame1 
      Height          =   1335
      Left            =   120
      TabIndex        =   4
      Top             =   240
      Width           =   1815
      Begin VB.OptionButton optRandom 
         Caption         =   "Random"
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   720
         Width           =   975
      End
      Begin VB.OptionButton optOrder 
         Caption         =   "Order"
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   960
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   840
         TabIndex        =   6
         Text            =   "0"
         Top             =   360
         Width           =   735
      End
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   1200
         Tag             =   "0"
         Top             =   240
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Update"
         Height          =   215
         Left            =   120
         TabIndex        =   5
         Top             =   0
         Width           =   825
      End
      Begin VB.Label Label1 
         Caption         =   "Minutes:"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "Next"
      Height          =   375
      Left            =   2160
      TabIndex        =   3
      Top             =   840
      Width           =   1455
   End
   Begin VB.CommandButton cmdRnd 
      Caption         =   "Random"
      Height          =   375
      Left            =   2160
      TabIndex        =   2
      Top             =   360
      Width           =   1455
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "Update List"
      Height          =   375
      Left            =   2160
      TabIndex        =   1
      Top             =   1320
      Width           =   1455
   End
   Begin VB.ListBox List1 
      Height          =   2010
      Left            =   3720
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   120
      Width           =   2655
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "mnuPopup"
      Visible         =   0   'False
      Begin VB.Menu mnuNext 
         Caption         =   "Next"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim MyPath As String
Dim WallPath As String
Dim Edit As Boolean

Private Sub Check1_Click()
    Timer1.Tag = 0
    Timer1.Enabled = Check1
End Sub

Private Sub cmdNext_Click()
    Timer1.Tag = 0
    If List1.ListCount Then List1.ListIndex = (List1.ListIndex + 1) Mod List1.ListCount
End Sub

Private Sub cmdRnd_Click()
    List1.ListIndex = Int(Rnd * List1.ListCount)
End Sub

Private Sub cmdUpdate_Click()
    Dim OldFile As String
    Dim File As String
    OldFile = List1
    List1.Clear
    File = Dir(WallPath)
    Do While Len(File)
        List1.AddItem File
        File = Dir
    Loop
    List1 = OldFile
End Sub

Private Sub SetPicture(ByVal Filename As String)
    On Error GoTo Fehler
    Dim xFile As String
    StatusBar1.Label1 = "Loading picture..."
    xFile = WinPath & "FlomixWall.bmp"
    If 1 = 2 Then
        picScreen.Cls
        picScreen.PaintPicture LoadPicture(Filename), 0, 0, picScreen.ScaleWidth, picScreen.ScaleHeight
        'FoxAlphaBlend picScreen.HDC, picScreen.ScaleWidth - picFlomix.ScaleWidth, picScreen.ScaleHeight - picFlomix.ScaleHeight - 30, picFlomix.ScaleWidth, picFlomix.ScaleHeight, picFlomix.HDC, 0, 0, 128, 0, 1
    Else
        Set picScreen.Picture = LoadPicture(Filename)
    End If
    'Picture zu Image
    picPreview.PaintPicture picScreen.Picture, 0, 0, picPreview.ScaleWidth, picPreview.ScaleHeight
    picPreview.Refresh
    SavePicture picScreen.Picture, xFile
    StatusBar1.Label1 = "Setting picture..."
    SystemParametersInfo SPI_SETDESKWALLPAPER, 0&, ByVal xFile, SPIF_SENDWININICHANGE Or SPIF_UPDATEINIFILE
    StatusBar1.Label1 = "Ready."
Fehler:
End Sub

Private Sub Form_Load()
    On Error Resume Next
    Edit = True
    MyPath = App.Path
    If Right(MyPath, 1) <> "\" Then MyPath = MyPath & "\"
    WallPath = MyPath & "Wallpapers\"
    If LCase(Command) = "autostart" Then WindowState = 1
    picScreen.Width = Screen.Width
    picScreen.Height = Screen.Height
    Me.StatusBar1.Label1 = "Ready."
    cmdUpdate_Click
    List1 = GetSetting("FLOMIX Studios", "Wallpaper Generator", "Current", "")
    Text1 = GetSetting("FLOMIX Studios", "Wallpaper Generator", "Update Delay", 1)
    Me.Check1 = GetSetting("FLOMIX Studios", "Wallpaper Generator", "Update", 1)
    Edit = False
End Sub

Private Sub Form_Resize()
    If WindowState = 1 Then Hide Else Show
End Sub

Private Sub List1_Click()
    If Not Edit Then
        SetPicture WallPath & List1.List(List1.ListIndex)
        SaveSetting "FLOMIX Studios", "Wallpaper Generator", "Current", List1
    End If
End Sub

Private Sub mnuNext_Click()
    cmdNext_Click
End Sub

Private Sub ShellIcon1_DblClick(Button As Integer)
    If Button = 1 Then WindowState = 0: Show: AppActivate Caption, wait
End Sub

Private Sub ShellIcon1_SingleClick(Button As Integer)
    PopupMenu mnuPopup, 2
End Sub

Private Sub Text1_Validate(Cancel As Boolean)
    Dim Text As String, Pos As Long
    Text = Val(Text1)
    Pos = InStr(Text, ",")
    If Pos Then Text = Left(Text, Pos - 1) & "." & Mid(Text, Pos + 1)
    Text1 = Text
    SaveSetting "FLOMIX Studios", "Wallpaper Generator", "Update Delay", Text
End Sub

Private Sub Timer1_Timer()
    On Error Resume Next
    If Val(Text1) <= 0 Then Exit Sub
    Timer1.Tag = Timer1.Tag + 1
    StatusBar1.StatusText = TimeString(Int(Val(Text1) * 60) - Timer1.Tag)
    StatusBar1.Label2 = List1
    ShellIcon1.ToolTipText = List1 & " (change in " & StatusBar1.StatusText & ".)"
    StatusBar1.Status = Timer1.Tag * 100 / Int((Val(Text1) * 60))
    If Timer1.Tag >= Int(Val(Text1) * 60) Then
        Timer1.Tag = 0
        If optOrder.Value = True Then cmdNext_Click Else cmdRnd_Click
    End If
End Sub
