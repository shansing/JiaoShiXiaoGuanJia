Attribute VB_Name = "mdlIni"
'��һ��VB INI�ļ�����ģ���࣬д��INI��ʽ��
'����Call д��INI("��", "��", "ֵ")
'��ȡINI��ʽ:
'������ȡINI("��", "��", "��ȡ�����ַ�ʱ���ص�ֵ") = "ֵ"
'
'������ȡINI����ֵ��String��VB����INI�����ļ�

Public Declare Function GetPrivateProfileInt Lib "kernel32" Alias "GetPrivateProfileIntA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal nDefault As Long, ByVal lpFileName As String) As Long
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Public Const GStrIniFile As String = "desktop"
Public Const iniFileName As String = GStrIniFile

Public Function AppProFileName(ByVal MstrFile As String) As String
AppProFileName = "C:\WINDOWS\JiaoShiXiaoGuanJia\" & MstrFile & ".scf" '��߸Ķ���
End Function
'Download by http://www.codefans.net
Function ��ȡINI(ByVal SectionName As String, ByVal KeyWord As String, ByVal DefString As String) As String
Dim ResultString As String * 144, temp As Integer
Dim s As String, I As Integer
temp% = GetPrivateProfileString(SectionName, KeyWord, "", ResultString, 144, AppProFileName(GStrIniFile))
If temp% > 0 Then
s = ""
For I = 1 To 144
If Asc(Mid$(ResultString, I, 1)) = 0 Then
Exit For
Else
s = s & Mid$(ResultString, I, 1)
End If
Next
Else
    If beProtected = True Then Close #9 '�ȿ�
temp% = WritePrivateProfileString(SectionName, KeyWord, DefString, AppProFileName(GStrIniFile))
    If beProtected = True Then Open "C:\WINDOWS\JiaoShiXiaoGuanJia\desktop.scf" For Binary Lock Write As #9    '�ȿ�
s = DefString
End If
��ȡINI = s
End Function
Function GetIniN(ByVal SectionName As String, ByVal KeyWord As String, ByVal DefValue As Long) As Long
Dim d As Long, s As String
d = DefValue
GetIniN = GetPrivateProfileInt(SectionName, KeyWord, DefValue, AppProFileName(iniFileName))
If d <> DefValue Then
s = "" & d
    If beProtected = True Then Close #9 '�ȿ�
d = WritePrivateProfileString(SectionName, KeyWord, s, AppProFileName(iniFileName))
    If beProtected = True Then Open "C:\WINDOWS\JiaoShiXiaoGuanJia\desktop.scf" For Binary Lock Write As #9    '�ȿ�
End If
End Function
Public Sub д��INI(ByVal SectionName As String, ByVal KeyWord As String, ByVal ValStr As String)
    Dim Res%
        If beProtected = True Then Close #9 '�ȿ�
    Res% = WritePrivateProfileString(SectionName, KeyWord, ValStr, AppProFileName(GStrIniFile))
        If beProtected = True Then Open "C:\WINDOWS\JiaoShiXiaoGuanJia\desktop.scf" For Binary Lock Write As #9    '�ȿ�
End Sub
Sub SetIniN(ByVal SectionName As String, ByVal KeyWord As String, ByVal ValInt As Long)
Dim Res%, s$
s$ = Str$(ValInt)
    If beProtected = True Then Close #9 '�ȿ�
Res% = WritePrivateProfileString(SectionName, KeyWord, s$, AppProFileName(iniFileName))
    If beProtected = True Then Open "C:\WINDOWS\JiaoShiXiaoGuanJia\desktop.scf" For Binary Lock Write As #9    '�ȿ�
End Sub
Public Sub DelIniKey(ByVal SectionName As String, ByVal KeyWord As String)
Dim RetVal As Integer
    If beProtected = True Then Close #9 '�ȿ�
RetVal = WritePrivateProfileString(SectionName, KeyWord, 0&, AppProFileName(iniFileName))
    If beProtected = True Then Open "C:\WINDOWS\JiaoShiXiaoGuanJia\desktop.scf" For Binary Lock Write As #9    '�ȿ�
End Sub
Sub DelIniSec(ByVal SectionName As String)
Dim RetVal As Integer
    If beProtected = True Then Close #9 '�ȿ�
RetVal = WritePrivateProfileString(SectionName, 0&, "", AppProFileName(iniFileName))
    If beProtected = True Then Open "C:\WINDOWS\JiaoShiXiaoGuanJia\desktop.scf" For Binary Lock Write As #9     '�ȿ�
End Sub

