Attribute VB_Name = "mdlShut"
Declare Function ExitWindowsEx Lib "user32" (ByVal uFlags As Long, ByVal dwReserved As Long) As Long
'ExitWindowsEx�Ĳ���uflags�����ĸ���Ӧֵ���ֱ��ǣ�
Public Const EWX_LOGOFF = 0 '�˳�(ע��)
Public Const EWX_SHUTDOWN = 1 '�ػ�
Public Const EWX_REBOOT = 2 '������
Public Const EWX_FORCE = 4 'ǿ�ƹػ�������֪ͨ���ڻӦ�ó������������ҹر�
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
'���������������NT�ػ���ʹ�õ�
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
AdjustTokenPrivilegesForNT '��95/98�е���û���ã���Ϊ�˺�NT���ݣ�д����
ExitWindowsEx EWX_FORCE Or EWX_SHUTDOWN, 0 '���ｫuFlgs��������������ɫ�������ᵽ���ĸ�����֮һ����
'����Ϊ����
         'ExitWindowsEx EWX_FORCE, 0 ǿ�ȹػ������ǲ�������Ҫ����Ķ�����ǿ�йرգ�
         'ExitWindowsEx EWX_LOGOFF, 0 �˳�(ע��)
End Sub


