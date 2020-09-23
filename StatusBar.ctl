VERSION 5.00
Begin VB.UserControl StatusBar 
   Alignable       =   -1  'True
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   Begin VB.PictureBox picField 
      BorderStyle     =   0  'Kein
      FillColor       =   &H8000000F&
      ForeColor       =   &H8000000F&
      Height          =   735
      Index           =   2
      Left            =   0
      ScaleHeight     =   49
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   161
      TabIndex        =   3
      Top             =   1800
      Width           =   2415
      Begin VB.Label labField 
         Height          =   375
         Index           =   2
         Left            =   240
         TabIndex        =   4
         Top             =   0
         Width           =   1215
      End
   End
   Begin VB.PictureBox picField 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'Kein
      FillColor       =   &H8000000D&
      ForeColor       =   &H80000008&
      Height          =   735
      Index           =   1
      Left            =   0
      ScaleHeight     =   49
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   98
      TabIndex        =   2
      Top             =   960
      Width           =   1470
   End
   Begin VB.PictureBox picField 
      BorderStyle     =   0  'Kein
      FillColor       =   &H8000000F&
      ForeColor       =   &H8000000F&
      Height          =   735
      Index           =   0
      Left            =   0
      ScaleHeight     =   49
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   97
      TabIndex        =   0
      Top             =   120
      Width           =   1455
      Begin VB.Label labField 
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   1
         Top             =   120
         Width           =   1215
      End
   End
End
Attribute VB_Name = "StatusBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

Private m_Status As Double
Private m_Rect As RECT
Private m_StatusText As String

Private Const GWL_EXSTYLE = (-20)

Private Sub picField_Resize(Index As Integer)
    With picField(Index)
        If Index = 1 Then
            UpdateStatus
        Else
            labField(Index).Move 1, 1, .ScaleWidth - 2, .ScaleHeight - 2
        End If
    End With
End Sub

Private Sub UserControl_Initialize()
    'this is just for the thin inset-effect
    For I = 0 To 2
        SetWindowLong picField(I).hWnd, GWL_EXSTYLE, GetWindowLong(picField(I).hWnd, GWL_EXSTYLE) Or &H8000
    Next I
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    m_Status = PropBag.ReadProperty("Status", 0)
    m_StatusText = PropBag.ReadProperty("StatusText", "")
    labField(0).Caption = PropBag.ReadProperty("Label1", "")
    labField(2).Caption = PropBag.ReadProperty("Label2", "")
    UpdateStatus
End Sub

Private Sub UserControl_Resize()
        picField(2).Move ScaleWidth - picField(2).Width, 0, picField(2).Width, ScaleHeight
        picField(1).Move picField(2).Left - picField(1).Width - 2, 0, picField(1).Width, ScaleHeight
        picField(0).Move 0, 0, picField(1).Left - 2, ScaleHeight
End Sub

Public Property Get Status() As Double
    Status = m_Status
End Property

Public Property Let Status(ByVal NewStatus As Double)
    m_Status = NewStatus
    If m_Status < 0 Then m_Status = 0
    If m_Status > 100 Then m_Status = 100
    PropertyChanged "Status"
    UpdateStatus
End Property

Public Property Get Label1() As String
Attribute Label1.VB_ProcData.VB_Invoke_Property = ";Text"
Attribute Label1.VB_UserMemId = -517
    Label1 = labField(0).Caption
End Property

Public Property Let Label1(ByVal newValue As String)
    labField(0).Caption = newValue
    labField(0).Refresh
    PropertyChanged "Label1"
End Property

Public Property Get StatusText() As String
    StatusText = m_StatusText
End Property

Public Property Let StatusText(ByVal newValue As String)
    m_StatusText = newValue
    UpdateStatus
    PropertyChanged "StatusText"
End Property

Public Property Get Label2() As String
Attribute Label2.VB_ProcData.VB_Invoke_Property = ";Text"
Attribute Label2.VB_UserMemId = -518
    Label2 = labField(2).Caption
End Property

Public Property Let Label2(ByVal newValue As String)
    labField(2).Caption = newValue
    labField(2).Refresh
    PropertyChanged "Label2"
End Property

Private Sub UpdateStatus()
    Dim X As Long, Y As Long, W As Long, H As Long
    m_Rect.Left = 0
    m_Rect.Top = 0
    m_Rect.Right = picField(1).ScaleWidth
    m_Rect.Bottom = picField(1).ScaleHeight
    X = 2: Y = 2
    W = (picField(1).ScaleWidth - 4) * m_Status / 100
    H = picField(1).ScaleHeight - 4
    picField(1).Cls
    If m_Status > 0 Then picField(1).Line (X, Y)-(W + X - 1, H + Y - 1), &H8000000D, BF
    If Len(m_StatusText) > 0 Then
        'this is to invert the part of the text over the dark StatusBar;
        'it inverts the StatusBar, draws the Text on it, and inverts the
        'Bar again. This way the Bar didn't change, whereas the Text is
        'inverted.
        If m_Status > 0 Then BitBlt picField(1).hdc, X, Y, W, H, picField(1).hdc, 0, 0, vbDstInvert
        DrawText picField(1).hdc, m_StatusText, Len(m_StatusText), m_Rect, &H25
        If m_Status > 0 Then BitBlt picField(1).hdc, X, Y, W, H, picField(1).hdc, 0, 0, vbDstInvert
    End If
    picField(1).Refresh
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Status", m_Status, 0)
    Call PropBag.WriteProperty("StatusText", m_StatusText, "")
    Call PropBag.WriteProperty("Label1", labField(0).Caption, "")
    Call PropBag.WriteProperty("Label2", labField(2).Caption, "")
End Sub
