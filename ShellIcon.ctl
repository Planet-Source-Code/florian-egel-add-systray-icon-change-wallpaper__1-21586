VERSION 5.00
Begin VB.UserControl ShellIcon 
   CanGetFocus     =   0   'False
   ClientHeight    =   1050
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1545
   ClipBehavior    =   0  'Keine
   ClipControls    =   0   'False
   HitBehavior     =   0  'Keine
   InvisibleAtRuntime=   -1  'True
   Picture         =   "ShellIcon.ctx":0000
   ScaleHeight     =   70
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   103
   ToolboxBitmap   =   "ShellIcon.ctx":030A
   Begin VB.Timer tmrMenu 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   720
      Top             =   120
   End
End
Attribute VB_Name = "ShellIcon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'!!!!   You should not change this code; you   !!!!
'!!!!   can customize everything in the IDE.   !!!!
'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!


Private Enum NIM_CONSTANTS
    NIM_ADD = &H0
    NIM_MODIFY = &H1
    NIM_DELETE = &H2
End Enum

Private Enum NIF_CONSTANTS
    NIF_MESSAGE = &H1
    NIF_ICON = &H2
    NIF_TIP = &H4
End Enum

Private Enum WM_CONSTANTS
    WM_MOUSEMOVE = &H200
    WM_LBUTTONDBLCLK = &H203
    WM_LBUTTONDOWN = &H201
    WM_LBUTTONUP = &H202
    WM_RBUTTONDBLCLK = &H206
    WM_RBUTTONDOWN = &H204
    WM_RBUTTONUP = &H205
End Enum

Private Type NOTIFYICONDATA
    cbSize As Long
    hWnd As Long
    uId As Long
    uFlags As NIF_CONSTANTS
    uCallBackMessage As WM_CONSTANTS
    hIcon As Long
    szTip As String * 64
End Type

Private Type POINTAPI
        X As Long
        Y As Long
End Type

Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As NIM_CONSTANTS, pnid As NOTIFYICONDATA) As Boolean
Private Declare Function GetSystemMenu Lib "user32" (ByVal hWnd As Long, ByVal bRevert As Long) As Long
Private Declare Function TrackPopupMenu Lib "user32" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal X As Long, ByVal Y As Long, ByVal nReserved As Long, ByVal hWnd As Long, ByVal lprc As Long) As Long

Dim IconData As NOTIFYICONDATA

Dim m_ToolTipText As String
Dim m_Visible As Boolean
Dim m_Show As Boolean
Dim m_SysMenu As Boolean

Event MouseMove()
Event MouseDown(Button As Integer)
Event MouseUp(Button As Integer)
Event DblClick(Button As Integer)
Event Click(Button As Integer)
Event SingleClick(Button As Integer)

Public Property Get ToolTipText() As String
Attribute ToolTipText.VB_Description = "Gibt den Text zurück, der angezeigt wird, wenn die Maus über dem Steuerelement verweilt, oder legt den Text fest."
    ToolTipText = IconData.szTip
End Property

Public Property Let ToolTipText(ByVal New_ToolTipText As String)
    m_ToolTipText = New_ToolTipText
    IconData.szTip = m_ToolTipText & Chr(0)
    PropertyChanged "ToolTipText"
    If m_Show Then Shell_NotifyIcon NIM_MODIFY, IconData
End Property

Public Property Get Icon() As StdPicture
    Set Icon = Picture
End Property

Public Property Set Icon(ByVal New_Icon As StdPicture)
    Set Picture = New_Icon
    PropertyChanged "Icon"
    IconData.hIcon = New_Icon.Handle
    If m_Show Then Shell_NotifyIcon NIM_MODIFY, IconData
End Property

Public Property Get Visible() As Boolean
    Visible = m_Visible
End Property

Public Property Let Visible(ByVal New_Visible As Boolean)
    m_Visible = New_Visible
    PropertyChanged "Visible"
    Show m_Visible
End Property

Public Property Get SysMenu() As Boolean
    SysMenu = m_SysMenu
End Property

Public Property Let SysMenu(ByVal New_SysMenu As Boolean)
    m_SysMenu = New_SysMenu
    PropertyChanged "SysMenu"
End Property

Private Sub tmrMenu_Timer()
    tmrMenu.Enabled = False
    RaiseEvent SingleClick(1)
End Sub

Private Sub UserControl_Initialize()
    IconData.cbSize = Len(IconData)
    IconData.hWnd = hWnd
    IconData.uId = vbNull
    IconData.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
    IconData.uCallBackMessage = WM_MOUSEMOVE
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Select Case X
        Case WM_MOUSEMOVE: RaiseEvent MouseMove
        Case WM_LBUTTONDBLCLK: RaiseEvent DblClick(1)
        Case WM_LBUTTONDOWN: RaiseEvent MouseDown(1)
        Case WM_LBUTTONUP
            RaiseEvent MouseUp(1)
            RaiseEvent Click(1)
            tmrMenu.Enabled = Not tmrMenu.Enabled
        Case WM_RBUTTONDBLCLK: RaiseEvent DblClick(2)
        Case WM_RBUTTONDOWN: RaiseEvent MouseDown(2)
        Case WM_RBUTTONUP
            If m_SysMenu Then
                ShowSysMenu
            Else
                RaiseEvent MouseUp(2): RaiseEvent Click(2)
            End If
    End Select
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    m_ToolTipText = PropBag.ReadProperty("ToolTipText", "")
    IconData.szTip = m_ToolTipText & Chr(0)
    Set Picture = PropBag.ReadProperty("Icon", Nothing)
    m_Visible = PropBag.ReadProperty("Visible", False)
    m_SysMenu = PropBag.ReadProperty("SysMenu", True)
    IconData.hIcon = Picture.Handle
    If Ambient.UserMode Then Show m_Visible
End Sub

Private Sub UserControl_Resize()
    Width = 480
    Height = 480
End Sub

Private Sub UserControl_Terminate()
    If m_Show Then Show False
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("ToolTipText", m_ToolTipText, "")
    Call PropBag.WriteProperty("Icon", Picture, Nothing)
    Call PropBag.WriteProperty("Visible", m_Visible, False)
    Call PropBag.WriteProperty("SysMenu", m_SysMenu, True)
End Sub

Private Sub Show(Optional ByVal Show As Boolean = True)
    If Show And m_Show = False Then
        If Ambient.UserMode Then
            Shell_NotifyIcon NIM_ADD, IconData
            m_Show = True
        End If
    ElseIf Show = False And m_Show = True Then
        Shell_NotifyIcon NIM_DELETE, IconData
        m_Show = False
    End If
End Sub

Public Sub ShowSysMenu(Optional ByVal hWnd As Long)
    Dim Pos As POINTAPI
    If hWnd = 0 Then hWnd = Parent.hWnd
    GetCursorPos Pos
    TrackPopupMenu GetSystemMenu(hWnd, False), &H200, Pos.X, Pos.Y, hWnd, hWnd, 0
End Sub
