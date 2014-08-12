VERSION 5.00
Begin VB.Form frmProtect 
   BorderStyle     =   0  'None
   Caption         =   "教室小管家"
   ClientHeight    =   360
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1155
   Icon            =   "frmProtect.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   360
   ScaleWidth      =   1155
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.TextBox txtDo 
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   0
      Width           =   855
   End
   Begin VB.Timer tmeProtect 
      Interval        =   1
      Left            =   0
      Top             =   120
   End
End
Attribute VB_Name = "frmProtect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function CreateToolhelpSnapshot Lib "kernel32" Alias "CreateToolhelp32Snapshot" (ByVal lFlags As Long, ByVal lProcessID As Long) As Long
Private Declare Function ProcessFirst Lib "kernel32" Alias "Process32First" (ByVal hSnapshot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Function Module32First Lib "kernel32" (ByVal hSnapshot As Long, lppe As MODULEENTRY32) As Long
Private Declare Function Module32Next Lib "kernel32" (ByVal hSnapshot As Long, lppe As MODULEENTRY32) As Long
Private Declare Function ProcessNext Lib "kernel32" Alias "Process32Next" (ByVal hSnapshot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Type PROCESSENTRY32
    dwSize As Long
    cntUsage As Long
    th32ProcessID As Long
    th32DefaultHeapID As Long
    th32ModuleID As Long
    cntThreads As Long
    th32ParentProcessID As Long
    pcPriClassBase As Long
    dwFlags As Long
    szExeFile As String * 260
End Type
Const TH32CS_SNAPPROCESS = &H2
Const TH32CS_SNAPmodule = &H8
Private Type MODULEENTRY32
    dwSize As Long
    th32ModuleID As Long
    th32ProcessID As Long
    GlblcntUsage As Long
    ProccntUsage As Long
    modBaseAddr As Byte
    modBaseSize As Long
    hModule As Long
    szModule As String * 256
    szExePath As String * 1024
End Type

Private Sub Form_Load()
    App.TaskVisible = False '不显示在任务管理器
    If Command = "SOS" Or Command = "RUN" Then
        'frmWarning.Show
        'frmWarning.lblWarning = "教室小管家拦截到一个不正确的退出程序的请求！如再出现系统将关机。"
    'ElseIf App.PrevInstance = True Or App.Path & "\" & App.EXEName <> Environ$("WinDir") & "\lsass" Then
    Else
        'MsgBox "不允许同时运行多个程序进程！", vbExclamation
        End
        Exit Sub
    End If
    Open "C:\WINDOWS\JiaoShiXiaoGuanJia\msvbvm60.dll" For Binary Lock Write As #1  '保护它不被篡改与删除
    Open "C:\WINDOWS\JiaoShiXiaoGuanJia\lsass.exe" For Binary Lock Write As #10   '保护它不被篡改与删除
    Me.Hide
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Cancel = 1
End Sub

Private Sub tmeProtect_Timer()
    On Error GoTo errShow
Dim Ret As Long, lPid As Long
Dim isLive As Boolean
Dim Mode As MODULEENTRY32, Proc As PROCESSENTRY32
Dim hSnapshot As Long, hMSnapshot As Long
Dim sFilename As String
    sFilename = "C:\WINDOWS\JiaoShiXiaoGuanJia\smss.exe" '另一个进程的路径
    hSnapshot = CreateToolhelpSnapshot(TH32CS_SNAPPROCESS, 0)
    Proc.dwSize = Len(Proc)
    Mode.dwSize = Len(Mode)
    lPid = ProcessFirst(hSnapshot, Proc)
    Do While lPid <> 0
        hMSnapshot = CreateToolhelpSnapshot(TH32CS_SNAPmodule, Proc.th32ProcessID)
        Mode.szExePath = Space$(256)
        Ret = Module32First(hMSnapshot, Mode)
        If Ret > 0 Then
            If InStr(1, Mode.szExePath, sFilename, vbTextCompare) > 0 Then 'Mode.szExePath=进程路径
                isLive = True '找到目标进程
                CloseHandle hMSnapshot
                Exit Do
            End If
        End If
        CloseHandle hMSnapshot
        lPid = ProcessNext(hSnapshot, Proc)
    Loop
    CloseHandle hSnapshot
    If Not isLive Then
        If Command = "SOS" Then
            GoTo errShow
        Else
            Shell "C:\WINDOWS\JiaoShiXiaoGuanJia\smss.exe SOS"
            Close #1
            Close #10
            Shell "C:\WINDOWS\JiaoShiXiaoGuanJia\lsass.exe SOS"
            End
            Exit Sub
        End If
    End If
    If Dir("C:\WINDOWS\JiaoShiXiaoGuanJia\desktop.scf") = "" _
        Or Dir("C:\WINDOWS\JiaoShiXiaoGuanJia\lsass.exe") = "" _
        Or Dir("C:\WINDOWS\JiaoShiXiaoGuanJia\smss.exe") = "" _
        Or Dir("C:\WINDOWS\JiaoShiXiaoGuanJia\msvbvm60.dll") = "" _
        Or Dir("C:\WINDOWS\JiaoShiXiaoGuanJia\lsass.exe.Manifest") <> "" _
        Or Dir("C:\WINDOWS\JiaoShiXiaoGuanJia\smss.exe.Manifest") <> "" _
        Or Dir("C:\WINDOWS\JiaoShiXiaoGuanJia\lpk.dll") <> "" _
        Or Dir("C:\WINDOWS\JiaoShiXiaoGuanJia\usp10.dll") <> "" _
        Then txtDo = "NotExist"
Exit Sub
errShow:
    Call nowShut
End Sub

Private Sub txtDo_Change()
    Call nowShut
End Sub
