Attribute VB_Name = "mdlFindWindows"
'*******************************************************************************************
'模块名称 EnumWindows
'AUTHOR Morn Woo 20120112
'1. 把下面的模块存储为EnumWindows.bas，并加入到你的工程中
'2 调用说明，直接在程序中使用FindWindows函数来找到窗口句柄；如果需要连续处理多个
'窗口，请自己修改EnumWindowsProc函数，在里面添加处理代码。
'*******************************************************************************************

Private Declare Function RtlAdjustPrivilege& Lib "ntdll" (ByVal Privileges As Long, Optional ByVal NewValue As Long = 1, Optional ByVal Thread As Long, Optional Value As Long)
Public Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const WM_CLOSE = &H10
Private Const WM_QUIT = &H12
Private Const WM_DISTORY = &H2


Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long

Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hWnd As Long) As Long

'EnumWindows枚举所有屏幕上的顶层窗口，并将窗口句柄传送给应用程序定义的回调函数。回调函数返回FALSE将停止枚举，否则EnumWindows函数继续到所有顶层窗口枚举完为止。
Declare Function enumwindows Lib "user32" Alias "EnumWindows" (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Boolean

Private mVarWindowCaptionString As String '需要比较的值
Private mVarFound As Boolean '是否找到指定的窗口
Private mVarFoundWindowHwnd As Long '找到窗口的句柄

'关闭进程所需
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, lpdwProcessId As Long) As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Private Sub KillProcess(ByVal whWnd As Long)
Call RtlAdjustPrivilege(20)

Dim lpdwProcessId As Long
Dim hProcessHandle As Long
GetWindowThreadProcessId whWnd, lpdwProcessId
hProcessHandle = OpenProcess(&H1F0FFF, True, lpdwProcessId)
If hProcessHandle <> 0 Then TerminateProcess hProcessHandle, ByVal 0&
CloseHandle (hProcessHandle)
End Sub


Public Property Let FoundWindowHwnd(ByRef vData As Long)
mVarFoundWindowHwnd = vData
End Property
Public Property Get FoundWindowHwnd() As Long
FoundWindowHwnd = mVarFoundWindowHwnd
End Property


Public Property Let Found(ByRef vData As Boolean)
mVarFound = vData
End Property
Public Property Get Found() As Boolean
Found = mVarFound
End Property

Public Property Let WindowCaptionString(ByRef vData As String)
mVarWindowCaptionString = vData
End Property
Public Property Get WindowCaptionString() As String
WindowCaptionString = mVarWindowCaptionString
End Property

Public Function EnumWindowsProc(ByVal hWnd As Long, ByVal lParam As Long) As Boolean
'Dim sA
Dim s As String
EnumWindowsProc = False
'sA = Trim(WindowCaptionString)
'If Trim(sA) = "" Then Exit Function
s = String(80, Chr(0))
Call GetWindowText(hWnd, s, 80)
s = Left(s, InStr(s, Chr(0)) - 1)
    Dim Wkn As Integer
    Wkn = 0
    While Wkn < frmMain.lstWk.ListCount
        If InStr(1, s, frmMain.lstWk.List(Wkn)) <= 0 Then
            Found = False
        Else
            KillProcess hWnd '结束进程
            SendMessage hWnd, &H11, 0, 0 '发送关机信息
            SendMessage hWnd, WM_QUIT, 0, 0
            SendMessage hWnd, WM_DISTORY, 0, 0
            SendMessage hWnd, WM_CLOSE, 0, 0 '关闭窗口
            ShowWindow hWnd, 0 '隐藏窗口
            frmWarning.Show
            frmWarning.lblWarning = "教室小管家已阻止了被禁止程序的运行！如有问题敬请联系电脑管理员。"
        End If
        Wkn = Wkn + 1
    Wend
    EnumWindowsProc = True '返回真继续枚举。
    'Found = True
    'FoundWindowHwnd = hwnd
End Function

Public Function FindWindows(ByVal sCaption As String) As Long '返回找到的窗口句柄
enumwindows AddressOf EnumWindowsProc, 0&
WindowCaptionString = sCaption
If Found Then
FindWindows = FoundWindowHwnd '这里是找到的窗口句柄
Else
Exit Function
End If

End Function

Public Function FindWindows2(ByVal sCaption As String) As Long '返回找到的窗口句柄
enumwindows AddressOf EnumWindowsProc2, 0&
WindowCaptionString = sCaption
If Found Then
FindWindows2 = FoundWindowHwnd '这里是找到的窗口句柄
Else
Exit Function
End If

End Function

Public Function EnumWindowsProc2(ByVal hWnd As Long, ByVal lParam As Long) As Boolean
'Dim sA
Dim s As String
EnumWindowsProc2 = False
'sA = Trim(WindowCaptionString)
'If Trim(sA) = "" Then Exit Function
s = String(80, Chr(0))
Call GetWindowText(hWnd, s, 80)
s = Left(s, InStr(s, Chr(0)) - 1)
If s <> "" Then
    Dim SWn As Integer
    SWn = 0
    While SWn < frmMain.lstSW.ListCount
        If s = frmMain.lstSW.List(SWn) Then GoTo continue
        SWn = SWn + 1
    Wend
    frmMain.lstSW.AddItem s
End If
continue:
    EnumWindowsProc2 = True '返回真继续枚举。
    'Found = True
    'FoundWindowHwnd = hwnd
End Function
