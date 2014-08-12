Attribute VB_Name = "mdlMain"
'对禁用视觉样式的支持
Public Declare Function CreateEllipticRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Long) As Long

'对窗口置顶的支持
Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

'对建立托盘图标的支持
Public Const MAX_TOOLTIP As Integer = 64
Public Const NIF_ICON = &H2
Public Const NIF_MESSAGE = &H1
Public Const NIF_TIP = &H4
Public Const NIM_ADD = &H0
Public Const NIM_DELETE = &H2
Public Const WM_MOUSEMOVE = &H200
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202
Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_RBUTTONDOWN = &H204
Public Const WM_RBUTTONUP = &H205
Public Const WM_RBUTTONDBLCLK = &H206
Public nfIconData As NOTIFYICONDATA
Public Type NOTIFYICONDATA
cbSize As Long
hWnd As Long
uID As Long
uFlags As Long
uCallbackMessage As Long
hIcon As Long
szTip As String * MAX_TOOLTIP
End Type
Public Declare Function Shell_NotifyIcon Lib "shell32.dll" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function SetForegroundWindow Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Const GWL_WNDPROC = -4
Public pWndProc As Long
'---------------------------------------------------------------
'RegisterWindowMessage获取分配给一个字串标识符的消息编号
'返回值Long，&C000 到 &FFFF之间的一个消息编号。零意味着出错
'参数 类型及说明
'lpString String，注册消息的名字
'注解如果没有一个子类处理程序的帮助，这个函数就没有什么用
Public Declare Function RegisterWindowMessage Lib "user32" Alias "RegisterWindowMessageA" _
                        (ByVal lpString As String) As Long
Public MsgTaskbarRestart As Long

'对刷新托盘的支持
Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" _
    (lpVersionInformation As OSVERSIONINFO) As Long
Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128        'Maintenance String For PSS usage
    OsName As String                    '操作系统的名称
End Type
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type
'Private Const WM_MOUSEMOVE = &H200


'对禁用关闭按钮的支持
Private Declare Function GetSystemMenu Lib "user32" (ByVal hWnd As Long, ByVal bRevert As Long) As Long
Private Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Private Const MF_REMOVE = &H1000
Private Const SC_CLOSE = &HF060

Public beProtected As Boolean

'对禁用关闭按钮的支持
Function Disabled(ChWnd As Long)
Dim hMenu, hendMenu As Long
Dim c As Long
hMenu = GetSystemMenu(ChWnd, 0)
RemoveMenu hMenu, SC_CLOSE, MF_REMOVE
End Function

'对重建系统托盘图标的支持
Public Function WndProc(ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    If Msg = WM_TRAYICON Then
        Select Case lParam
            Case WM_LBUTTONDOWN
            FrmAbout.Show 0
            Case WM_RBUTTONDOWN
            SetForegroundWindow hWnd '关键的一步,使菜单重画
            Case WM_RBUTTONUP
            Form1.PopupMenu Form1.mnuIndex  '显示系统菜单
        End Select
    End If
    'Explorer.exe 崩溃之后重建任务栏图标
    If Msg = MsgTaskbarRestart Then
        '原理：Explorer.exe 重新载入后会重建系统任务栏。当系统任务栏建立的时候会向系统内所有
        '注册接收TaskbarCreated 消息的顶级窗口发送一条消息，我们只需要捕捉这个消息，并重建系统托盘的图标即可。
        Shell_NotifyIcon NIM_ADD, nfIconData   '关键的一步,使图标重建
        WndProc = True
    End If
    WndProc = CallWindowProc(pWndProc, hWnd, Msg, wParam, lParam)
End Function

'密码算法
Public Function passWord(word As String) As String
    passWord = MD5("JSXGJ" & word & "!@@#")
    passWord = Right(passWord, 24 - Len(word)) & Right(passWord, Len(word) + 8)
    passWord = Left(passWord, 30)
End Function



Private Function GetSysTrayWnd() As Long
'返回系统托盘的句柄，适合于Windows各版本
Dim Result As Long
Dim Ver As OSVERSIONINFO
    Ver.dwOSVersionInfoSize = 148
    GetVersionEx Ver

 Result = FindWindow("Shell_TrayWnd", vbNullString)
 Result = FindWindowEx(Result, 0, "TrayNotifyWnd", vbNullString)
 'if version is  xp or 2k3 then run the next
 If Ver.dwMajorVersion = 5 And Ver.dwMinorVersion > 0 Then Result = FindWindowEx(Result, 0, "SysPager", vbNullString)
 'if version is xp 2k 2k3 then run the next
 If Ver.dwMajorVersion = 5 Then Result = FindWindowEx(Result, 0, "ToolbarWindow32", vbNullString)
GetSysTrayWnd = Result
End Function

Public Sub RefreshTrayIcon()
'刷新系统托盘图标
Dim hwndTrayToolBar As Long
Dim x, y As Long
Dim rTrayToolBar As RECT
Dim pos As Long
hwndTrayToolBar = GetSysTrayWnd
GetClientRect hwndTrayToolBar, rTrayToolBar
For x = 1 To rTrayToolBar.Right - 1
  For y = 1 To rTrayToolBar.Bottom - 1
   pos = (x And &HFFFF) + (y And &HFFFF) * &H10000
 PostMessage hwndTrayToolBar, WM_MOUSEMOVE, 0, pos '模拟鼠标移上去
 Next
Next
End Sub

