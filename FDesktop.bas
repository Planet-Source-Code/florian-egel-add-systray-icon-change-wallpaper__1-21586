Attribute VB_Name = "basFDesktop"
Public Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As Any, ByVal fuWinIni As Long) As Long
Public Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long

'Public Declare Function FoxAlphaBlend Lib "FoxCBmp3.dll" (ByVal HDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hScrDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal alpha As Byte, Optional ByVal MaskColor As Long, Optional ByVal Flags As Long) As Long

Public Const SPI_SETDESKWALLPAPER = 20
Public Const SPIF_SENDWININICHANGE = &H2
Public Const SPIF_UPDATEINIFILE = &H1

Dim m_WinPath As String

Public Function WinPath() As String
    'This function retrieves the Windows path.
    If m_WinPath = "" Then
        m_WinPath = String(1024, 0)
        GetWindowsDirectory m_WinPath, Len(m_WinPath)
        m_WinPath = Left(m_WinPath, InStr(m_WinPath, Chr(0)) - 1)
        If Right(m_WinPath, 1) <> "\" Then m_WinPath = m_WinPath & "\"
    End If
    WinPath = m_WinPath
End Function

Public Function TimeString(ByVal Seconds As Long) As String
    Dim M As Long, S As Long, H As Long
    S = Seconds
    M = S \ 60: S = S Mod 60
    H = M \ 60: M = M Mod 60
    
    If H Then
        TimeString = H & ":" & String(2 - Len(CStr(M)), "0") & M & " h"
    ElseIf M Then
        TimeString = M & ":" & String(2 - Len(CStr(S)), "0") & S & " min"
    Else
        TimeString = S & " sec"
    End If
    
End Function
