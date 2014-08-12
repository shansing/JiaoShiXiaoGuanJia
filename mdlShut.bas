Attribute VB_Name = "mdlShut"
Declare Function ExitWindowsEx Lib "user32" (ByVal uFlags As Long, ByVal dwReserved As Long) As Long
'ExitWindowsEx的参数uflags，有四个对应值，分别是：
Public Const EWX_LOGOFF = 0 '退出(注销)
Public Const EWX_SHUTDOWN = 1 '关机
Public Const EWX_REBOOT = 2 '重启动
Public Const EWX_FORCE = 4 '强制关机，即不通知现在活动应用程序让其先自我关闭
Public Const TOKEN_ADJUST_PRIVILEGES = &H20
Public Const TOKEN_QUERY = &H8
Public Const SE_PRIVILEGE_ENABLED = &H2
Public Const ANYSIZE_ARRAY = 1
Type LUID
lowpart As Long
highpart As Long
End Type
Type LUID_AND_ATTRIBUTES
pLuid As LUID
Attributes As Long
End Type
Type TOKEN_PRIVILEGES
PrivilegeCount As Long
Privileges(ANYSIZE_ARRAY) As LUID_AND_ATTRIBUTES
End Type
Declare Function GetCurrentProcess Lib "kernel32" () As Long
Declare Function LookupPrivilegeValue Lib "advapi32.dll" Alias "LookupPrivilegeValueA" (ByVal lpSystemName As String, ByVal lpName As String, lpLuid As LUID) As Long
Declare Function AdjustTokenPrivileges Lib "advapi32.dll" (ByVal TokenHandle As Long, ByVal DisableAllPrivileges As Long, NewState As TOKEN_PRIVILEGES, ByVal BufferLength As Long, PreviousState As TOKEN_PRIVILEGES, ReturnLength As Long) As Long
Declare Function OpenProcessToken Lib "advapi32.dll" (ByVal ProcessHandle As Long, ByVal DesiredAccess As Long, TokenHandle As Long) As Long
'这个函数就是用于NT关机中使用的
Sub AdjustTokenPrivilegesForNT()
Dim hdlProcessHandle As Long
Dim hdlTokenHandle As Long
Dim tmpLuid As LUID
Dim tkp As TOKEN_PRIVILEGES
Dim tkpNewButIgnored As TOKEN_PRIVILEGES
Dim lBufferNeeded As Long
hdlProcessHandle = GetCurrentProcess()
OpenProcessToken hdlProcessHandle, (TOKEN_ADJUST_PRIVILEGES Or _
TOKEN_QUERY), hdlTokenHandle
LookupPrivilegeValue "", "SeShutdownPrivilege", tmpLuid
tkp.PrivilegeCount = 1
tkp.Privileges(0).pLuid = tmpLuid
tkp.Privileges(0).Attributes = SE_PRIVILEGE_ENABLED
AdjustTokenPrivileges hdlTokenHandle, False, tkp, _
Len(tkpNewButIgnored), tkpNewButIgnored, _
lBufferNeeded
End Sub
Public Sub nowShut()
AdjustTokenPrivilegesForNT '在95/98中调用没作用，但为了和NT兼容，写上无
ExitWindowsEx EWX_FORCE Or EWX_SHUTDOWN, 0 '这里将uFlgs换成以上面标记蓝色字中所提到的四个参数之一即可
'下面为举例
         'ExitWindowsEx EWX_FORCE, 0 强迫关机（就是不管有无要保存的东西而强行关闭）
         'ExitWindowsEx EWX_LOGOFF, 0 退出(注销)
End Sub


