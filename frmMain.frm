VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "教室小管家"
   ClientHeight    =   8265
   ClientLeft      =   5565
   ClientTop       =   4920
   ClientWidth     =   11940
   ForeColor       =   &H00C0C0C0&
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8265
   ScaleWidth      =   11940
   StartUpPosition =   2  '屏幕中心
   Begin VB.OptionButton optTab 
      Caption         =   "随机换壁纸"
      Height          =   295
      Index           =   2
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   120
      Width           =   1125
   End
   Begin VB.OptionButton optTab 
      Caption         =   "程序阻止"
      Height          =   295
      Index           =   1
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   120
      Width           =   885
   End
   Begin VB.OptionButton optTab 
      Caption         =   "简单设置"
      Height          =   295
      Index           =   0
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   120
      Value           =   -1  'True
      Width           =   885
   End
   Begin VB.CommandButton cmdTop 
      Caption         =   "请勿使用不正当方法来显示本窗口！或者，你想现在关机？"
      Height          =   180
      Left            =   0
      TabIndex        =   26
      Top             =   0
      Width           =   255
   End
   Begin VB.Frame fraTab 
      BorderStyle     =   0  'None
      Caption         =   "fraTab"
      Height          =   3375
      Index           =   0
      Left            =   120
      TabIndex        =   34
      Top             =   480
      Width           =   5655
      Begin VB.TextBox txtAbout 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   800
         Left            =   1200
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   46
         Text            =   "frmMain.frx":058A
         Top             =   2575
         Width           =   4455
      End
      Begin VB.Frame fraShut 
         Caption         =   "自动关机设置"
         ForeColor       =   &H80000008&
         Height          =   1815
         Left            =   4080
         TabIndex        =   40
         Top             =   0
         Width           =   1575
         Begin VB.TextBox txtNight 
            Height          =   270
            IMEMode         =   3  'DISABLE
            Left            =   120
            MaxLength       =   11
            TabIndex        =   14
            Top             =   960
            Width           =   1095
         End
         Begin VB.Timer tmeShut 
            Enabled         =   0   'False
            Interval        =   10000
            Left            =   1680
            Top             =   0
         End
         Begin VB.TextBox txtNoon 
            Height          =   270
            IMEMode         =   3  'DISABLE
            Left            =   120
            MaxLength       =   11
            TabIndex        =   13
            Top             =   420
            Width           =   1095
         End
         Begin VB.CheckBox chkShut 
            Caption         =   "开启定时关机"
            BeginProperty Font 
               Name            =   "黑体"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   12
            Top             =   0
            Width           =   1380
         End
         Begin VB.Label lblShut 
            BackStyle       =   0  'Transparent
            Caption         =   "到时间后将提示“30秒后关机”"
            ForeColor       =   &H80000011&
            Height          =   375
            Left            =   120
            TabIndex        =   48
            Top             =   1320
            Width           =   1335
         End
         Begin VB.Label lblNoon 
            BackStyle       =   0  'Transparent
            Caption         =   "中午午休："
            ForeColor       =   &H80000007&
            Height          =   255
            Left            =   120
            TabIndex        =   42
            Top             =   240
            Width           =   975
         End
         Begin VB.Label lblNight 
            BackStyle       =   0  'Transparent
            Caption         =   "晚上就寝："
            ForeColor       =   &H80000007&
            Height          =   255
            Left            =   120
            TabIndex        =   41
            Top             =   720
            Width           =   975
         End
      End
      Begin VB.Frame fraBan 
         Caption         =   "系统组件设置"
         ForeColor       =   &H80000008&
         Height          =   1815
         Left            =   2040
         TabIndex        =   39
         Top             =   0
         Width           =   1935
         Begin VB.CheckBox chkTaskmgr 
            Caption         =   "禁用任务管理器"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   7
            Top             =   240
            Width           =   1695
         End
         Begin VB.CheckBox chkShutdown 
            Caption         =   "禁用 Shutdown"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   8
            Top             =   480
            Width           =   1695
         End
         Begin VB.CheckBox chkCMD 
            Caption         =   "禁用命令提示符"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   9
            Top             =   720
            Width           =   1695
         End
         Begin VB.CheckBox chkRegedit 
            Caption         =   "禁用注册表编辑器"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   10
            Top             =   960
            Width           =   1780
         End
         Begin VB.CheckBox chkWScript 
            Caption         =   "禁用 wScript"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   11
            Top             =   1200
            Width           =   1695
         End
         Begin VB.Label lblBan 
            BackStyle       =   0  'Transparent
            Caption         =   "建议配合程序阻止使用"
            ForeColor       =   &H80000011&
            Height          =   375
            Left            =   100
            TabIndex        =   47
            Top             =   1500
            Width           =   1815
         End
      End
      Begin VB.Frame fraBase 
         Caption         =   "基础设置"
         ForeColor       =   &H80000008&
         Height          =   1815
         Left            =   0
         TabIndex        =   35
         Top             =   0
         Width           =   1935
         Begin VB.CheckBox chkProtect 
            Caption         =   "开启自我保护"
            Height          =   180
            Left            =   120
            TabIndex        =   4
            Top             =   480
            Width           =   1575
         End
         Begin VB.CheckBox chkAuto 
            Caption         =   "开机自启动"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   3
            Top             =   240
            Width           =   1695
         End
         Begin VB.TextBox txtNewPass 
            Height          =   270
            IMEMode         =   3  'DISABLE
            Left            =   120
            MaxLength       =   16
            PasswordChar    =   "*"
            TabIndex        =   6
            Top             =   1440
            Width           =   1695
         End
         Begin VB.TextBox txtOldPass 
            Height          =   270
            IMEMode         =   3  'DISABLE
            Left            =   120
            MaxLength       =   16
            PasswordChar    =   "*"
            TabIndex        =   5
            Top             =   960
            Width           =   1695
         End
         Begin VB.Label lblNewPass 
            BackStyle       =   0  'Transparent
            Caption         =   "更改密码："
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   38
            Top             =   750
            Width           =   1095
         End
         Begin VB.Label lblNewPass3 
            BackStyle       =   0  'Transparent
            Caption         =   $"frmMain.frx":059D
            ForeColor       =   &H80000011&
            Height          =   255
            Left            =   120
            TabIndex        =   37
            Top             =   1270
            Width           =   1815
         End
         Begin VB.Label lblNewPass2 
            BackStyle       =   0  'Transparent
            Caption         =   $"frmMain.frx":05B7
            ForeColor       =   &H80000011&
            Height          =   255
            Left            =   1010
            TabIndex        =   36
            Top             =   780
            Width           =   975
         End
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000011&
         X1              =   0
         X2              =   5640
         Y1              =   1920
         Y2              =   1920
      End
      Begin VB.Label lblWriter 
         Caption         =   "欢迎使用"
         ForeColor       =   &H80000011&
         Height          =   225
         Left            =   4920
         TabIndex        =   45
         Top             =   2040
         Width           =   1335
      End
      Begin VB.Label lblVer 
         Caption         =   "2014 Beta"
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   3720
         TabIndex        =   44
         Top             =   2280
         Width           =   1095
      End
      Begin VB.Label lblName 
         BackStyle       =   0  'Transparent
         Caption         =   "教室小管家"
         BeginProperty Font 
            Name            =   "黑体"
            Size            =   24
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   615
         Left            =   1200
         TabIndex        =   43
         Top             =   2040
         Width           =   2655
      End
      Begin VB.Image imgIcon 
         Height          =   960
         Left            =   120
         Top             =   2040
         Width           =   960
      End
   End
   Begin VB.Frame fraTab 
      BorderStyle     =   0  'None
      Caption         =   "随机更换壁纸"
      Height          =   3375
      Index           =   2
      Left            =   5880
      TabIndex        =   33
      Top             =   240
      Visible         =   0   'False
      Width           =   5655
      Begin VB.CommandButton cmdWallRefresh 
         Caption         =   "刷新"
         Height          =   375
         Left            =   2090
         TabIndex        =   23
         Top             =   240
         Width           =   735
      End
      Begin VB.OptionButton optWallC 
         Caption         =   "拉伸"
         Height          =   255
         Left            =   2760
         TabIndex        =   55
         Top             =   2740
         Width           =   735
      End
      Begin VB.OptionButton optWallB 
         Caption         =   "平铺"
         Height          =   255
         Left            =   1920
         TabIndex        =   54
         Top             =   2740
         Width           =   735
      End
      Begin VB.OptionButton optWallA 
         Caption         =   "居中"
         Height          =   255
         Left            =   1080
         TabIndex        =   53
         Top             =   2740
         Width           =   735
      End
      Begin VB.Timer tmeWall 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   2400
         Tag             =   "0"
         Top             =   2640
      End
      Begin VB.TextBox txtWall 
         Height          =   270
         Left            =   3285
         MaxLength       =   6
         TabIndex        =   57
         Text            =   "3600"
         Top             =   3080
         Width           =   615
      End
      Begin VB.CommandButton cmdWall 
         Caption         =   "现在就换个"
         Height          =   375
         Left            =   4320
         TabIndex        =   59
         ToolTipText     =   "现在就随机换个壁纸"
         Top             =   3000
         Width           =   1215
      End
      Begin VB.CheckBox chkWallTime 
         Caption         =   "每隔       秒"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2640
         TabIndex        =   58
         Top             =   3080
         Width           =   1575
      End
      Begin VB.CheckBox chkWallLog 
         Caption         =   "本软件启动时"
         BeginProperty Font 
            Name            =   "黑体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1020
         TabIndex        =   56
         Top             =   3080
         Width           =   1455
      End
      Begin VB.FileListBox filWall 
         Height          =   2430
         Hidden          =   -1  'True
         Left            =   2880
         Pattern         =   "*.bmp;*.dib;*.gif;*.jpg"
         System          =   -1  'True
         TabIndex        =   49
         Top             =   280
         Width           =   2655
      End
      Begin VB.DirListBox dirWall 
         Height          =   1980
         Left            =   120
         TabIndex        =   25
         Top             =   660
         Width           =   2655
      End
      Begin VB.DriveListBox drvWall 
         Height          =   300
         Left            =   120
         TabIndex        =   24
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label lblWall4 
         Caption         =   "显示方式："
         Height          =   255
         Left            =   120
         TabIndex        =   60
         Top             =   2760
         Width           =   975
      End
      Begin VB.Label lblWall3 
         Caption         =   "条件设置："
         Height          =   255
         Left            =   120
         TabIndex        =   52
         Top             =   3120
         Width           =   975
      End
      Begin VB.Label lblWall 
         Caption         =   "图片文件显示："
         Height          =   255
         Left            =   2880
         TabIndex        =   51
         Top             =   120
         Width           =   1335
      End
      Begin VB.Label lblWall1 
         Caption         =   "设定目录："
         Height          =   255
         Left            =   120
         TabIndex        =   50
         Top             =   120
         Width           =   1095
      End
   End
   Begin VB.Timer tmeProtect 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   5400
      Top             =   0
   End
   Begin VB.Frame fraTab 
      BorderStyle     =   0  'None
      Caption         =   "窗口阻止设置"
      ForeColor       =   &H80000008&
      Height          =   3375
      Index           =   1
      Left            =   120
      TabIndex        =   27
      Top             =   4320
      Visible         =   0   'False
      Width           =   5655
      Begin VB.HScrollBar scoWk 
         Height          =   205
         LargeChange     =   5
         Left            =   2520
         Max             =   2000
         Min             =   1
         TabIndex        =   21
         Top             =   3100
         Value           =   1
         Width           =   2415
      End
      Begin VB.CheckBox chkWk 
         Caption         =   "开启程序阻止"
         BeginProperty Font 
            Name            =   "黑体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   3120
         Width           =   1575
      End
      Begin VB.CommandButton cmdSW 
         Caption         =   "刷新"
         Height          =   350
         Left            =   4920
         TabIndex        =   15
         Top             =   30
         Width           =   615
      End
      Begin VB.ListBox lstSW 
         ForeColor       =   &H00000000&
         Height          =   2580
         ItemData        =   "frmMain.frx":05C7
         Left            =   3720
         List            =   "frmMain.frx":05C9
         Sorted          =   -1  'True
         TabIndex        =   16
         Top             =   360
         Width           =   1815
      End
      Begin VB.Timer tmeWk 
         Enabled         =   0   'False
         Interval        =   1
         Left            =   3120
         Top             =   1320
      End
      Begin VB.TextBox txtWkAdd 
         Height          =   270
         Left            =   2040
         TabIndex        =   17
         Top             =   480
         Width           =   1455
      End
      Begin VB.ListBox lstWk 
         ForeColor       =   &H00000000&
         Height          =   2580
         ItemData        =   "frmMain.frx":05CB
         Left            =   120
         List            =   "frmMain.frx":0650
         Sorted          =   -1  'True
         TabIndex        =   19
         Top             =   360
         Width           =   1815
      End
      Begin VB.CommandButton cmdWkRemove 
         BackColor       =   &H00008000&
         Caption         =   "移除"
         Height          =   375
         Left            =   2040
         TabIndex        =   20
         Top             =   1680
         Width           =   975
      End
      Begin VB.CommandButton cmdWkAdd 
         BackColor       =   &H00008000&
         Caption         =   "添加"
         Height          =   375
         Left            =   2040
         TabIndex        =   18
         Top             =   840
         Width           =   975
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000011&
         X1              =   3600
         X2              =   3600
         Y1              =   120
         Y2              =   3000
      End
      Begin VB.Label lblSW 
         Caption         =   "当前活动对象："
         Height          =   255
         Left            =   3720
         TabIndex        =   32
         Top             =   120
         Width           =   1455
      End
      Begin VB.Label lblWk 
         Caption         =   "关键词列表："
         Height          =   255
         Left            =   120
         TabIndex        =   31
         Top             =   120
         Width           =   1455
      End
      Begin VB.Label lblWk1 
         BackStyle       =   0  'Transparent
         Caption         =   "加快响应"
         ForeColor       =   &H80000007&
         Height          =   255
         Left            =   1800
         TabIndex        =   29
         Top             =   3120
         Width           =   735
      End
      Begin VB.Label lblWkInfo 
         BackStyle       =   0  'Transparent
         Caption         =   "总计 N 项"
         ForeColor       =   &H80000007&
         Height          =   615
         Left            =   2040
         TabIndex        =   28
         Top             =   2400
         Width           =   1455
      End
      Begin VB.Label lblWk2 
         BackStyle       =   0  'Transparent
         Caption         =   "减少占用"
         ForeColor       =   &H80000007&
         Height          =   255
         Left            =   4950
         TabIndex        =   30
         Top             =   3120
         Width           =   735
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public cmdShell As String
Private mintCurFrame As Integer ' Current Frame visible

'对更换壁纸的支持
Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByVal lpvParam As Any, ByVal fuWinIni As Long) As Long
Const SPI_SETDESKWALLPAPER = 20
Const SPIF_SENDWININICHANGE = &H2
Const SPIF_UPDATEINIFILE = &H1

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

Private Sub chkAuto_Click()
    Dim Reg As Variant '注册表对象
    Dim KeyVal As String '当前键值
    Dim ChnVal As String '修改后键值
    Dim Key As String '目标键名
    Key = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\RunOnce\JiaoShiXiaoGuanJia" '设置目标键
    ChnVal = Chr(34) & "C:\WINDOWS\JiaoShiXiaoGuanJia\smss.exe" & Chr(34) '设置修改后键值
    Set Reg = CreateObject("Wscript.Shell") '获得注册表对象
    If chkAuto.Value = 1 Then
        On Error GoTo errShow
        Call 写入INI("Base", "Auto", "1")
        If Reg.Regread(Key) <> ChnVal Then GoTo errShow
    Else
        Reg.RegDelete Key
        Call 写入INI("Base", "Auto", "0")
    End If
Exit Sub
errShow:
    Reg.Regwrite Key, ChnVal, "REG_SZ"  '修改键值
End Sub

Private Sub chkCMD_Click()
    If chkCMD.Value = 1 Then
        Open Environ$("WinDir") & "\system32\cmd.exe" For Binary Lock Read As #3
        Call 写入INI("Ban", "CMD", "1")
    Else
        Close #3
        Call 写入INI("Ban", "CMD", "0")
    End If
End Sub

Private Sub chkProtect_Click()
    If chkProtect.Value = 1 Then
        Call 写入INI("Base", "Protect", "1")
        If beProtected = False Or Command = "SOS" Then
            frmWarning.Show
            frmWarning.lblWarning = "您刚刚选择了开启自我保护，这需要重启生效哦。"
        End If
    Else
        Call 写入INI("Base", "Protect", "0")
        If beProtected = True Or Command = "SOS" Then
            frmWarning.Show
            frmWarning.lblWarning = "您刚刚关闭了自我保护，这需要重启生效哦。"
        End If
    End If
End Sub

Private Sub chkShut_Click()
    If chkShut.Value = 1 Then
        txtNoon.Enabled = False
        txtNight.Enabled = False
        tmeShut.Enabled = True
        Call tmeShut_Timer
        Call 写入INI("Shut", "Shut", "1")
    Else
        tmeShut.Enabled = False
        txtNoon.Enabled = True
        txtNight.Enabled = True
        Call 写入INI("Shut", "Shut", "0")
    End If
    lblShut.Caption = "到时间后将提示“30秒后关机”"
End Sub

Private Sub chkWallLog_Click()
    Call 写入INI("Wall", "Login", chkWallLog.Value)
End Sub

Private Sub chkWallTime_Click()
    If chkWallTime.Value = 1 Then
        If txtWall <> "" And txtWall <> "0" And txtWall <> "00" And txtWall <> "000000" And txtWall <> "000" And txtWall <> "0000" And txtWall <> "00000" Then
            txtWall.Text = Replace(Str(Val(txtWall)), " ", "")
            tmeWall.Enabled = True
            txtWall.Enabled = False
        Else
            chkWallTime.Value = 0
            Exit Sub
        End If
    Else
        tmeWall.Enabled = False
        txtWall.Enabled = True
        tmeWall.Tag = 0
    End If
    Call 写入INI("Wall", "Time", chkWallTime.Value)
End Sub

Private Sub chkWScript_Click()
    If chkWScript.Value = 1 Then
        Open Environ$("WinDir") & "\system32\wscript.exe" For Binary Lock Read As #5
        Call 写入INI("Ban", "wScript", "1")
    Else
        Close #5
        Call 写入INI("Ban", "wScript", "0")
    End If
End Sub

Private Sub chkWk_Click()
    If chkWk.Value = 1 Then
        tmeWk.Enabled = True
        Call 写入INI("Wk", "Wk", "1")
    Else
        tmeWk.Enabled = False
        Call 写入INI("Wk", "Wk", "0")
    End If
End Sub

Private Sub chkRegedit_Click()
    If chkRegedit.Value = 1 Then
        Open Environ$("WinDir") & "\regedit.exe" For Binary Lock Read As #4
        Call 写入INI("Ban", "Regedit", "1")
    Else
        Close #4
        Call 写入INI("Ban", "Regedit", "0")
    End If
End Sub

Private Sub chkTaskmgr_Click()
    If chkTaskmgr.Value = 1 Then
        Open Environ$("WinDir") & "\system32\taskmgr.exe" For Binary Lock Read As #1
        Call 写入INI("Ban", "Taskmgr", "1")
    Else
        Close #1
        Call 写入INI("Ban", "Taskmgr", "0")
    End If
End Sub

Private Sub chkShutdown_Click()
    If chkShutdown.Value = 1 Then
        Open Environ$("WinDir") & "\system32\shutdown.exe" For Binary Lock Read As #2
        Call 写入INI("Ban", "Shutdown", "1")
    Else
        Close #2
        Call 写入INI("Ban", "Shutdown", "0")
    End If
End Sub

Private Sub cmdExit_Click()
        Call Shell_NotifyIcon(NIM_DELETE, nfIconData)
        End
End Sub

Private Sub SaveWk()
    Dim Wkn As Integer
    Wkn = 0
    While Wkn < lstWk.ListCount
        Call 写入INI("Wk", "List" & (Wkn + 1), lstWk.List(Wkn))
        Wkn = Wkn + 1
    Wend
    Call 写入INI("Wk", "List" & (Wkn + 1), "教室小管家")
End Sub

Private Sub cmdSW_Click()
    lstSW.Clear
    FindWindows2 ""
End Sub

Private Sub cmdTop_Click()
    nowShut
End Sub

Private Sub cmdWall_Click()
    If dirWall.Path = "C:\WINDOWS\JiaoShiXiaoGuanJia" Or dirWall.Path = "c:\WINDOWS\JiaoShiXiaoGuanJia" Then Exit Sub
    If filWall.ListCount > 0 Then
        Dim retryNo As Integer
        retryNo = 0
start:
        On Error GoTo errShow
        Dim wallID As Integer
        Randomize
        wallID = Int(filWall.ListCount * Rnd + 1)
        SavePicture LoadPicture(dirWall.Path & "\" & filWall.List(wallID)), "C:\WINDOWS\JiaoShiXiaoGuanJia\wallpaper.bmp"
        Dim Reg As Variant '注册表对象
        Set Reg = CreateObject("Wscript.Shell") '获得注册表对象
        If optWallA.Value = True Then
            Reg.Regwrite "HKEY_CURRENT_USER\Control Panel\desktop\TileWallpaper", "0", "REG_SZ"
            Reg.Regwrite "HKEY_CURRENT_USER\Control Panel\desktop\WallpaperStyle", "0", "REG_SZ"
        ElseIf optWallB.Value = True Then
            Reg.Regwrite "HKEY_CURRENT_USER\Control Panel\desktop\TileWallpaper", "1", "REG_SZ"
            Reg.Regwrite "HKEY_CURRENT_USER\Control Panel\desktop\WallpaperStyle", "0", "REG_SZ"
        Else
            Reg.Regwrite "HKEY_CURRENT_USER\Control Panel\desktop\TileWallpaper", "0", "REG_SZ"
            Reg.Regwrite "HKEY_CURRENT_USER\Control Panel\desktop\WallpaperStyle", "2", "REG_SZ"
        End If
        SystemParametersInfo SPI_SETDESKWALLPAPER, 0, "C:\WINDOWS\JiaoShiXiaoGuanJia\wallpaper.bmp", 0
        Reg.Regwrite "HKEY_CURRENT_USER\Control Panel\desktop\Wallpaper", "C:\WINDOWS\JiaoShiXiaoGuanJia\wallpaper.bmp", "REG_SZ"
    End If
Exit Sub
errShow:
    retryNo = retryNo + 1
    If retryNo < 3 Then GoTo start
End Sub

Private Sub cmdWallRefresh_Click()
    drvWall.Refresh
    dirWall.Refresh
    filWall.Refresh
End Sub

Private Sub cmdWkAdd_Click()
    If txtWkAdd = "" Then
        lblWkInfo = "请输入关键词！"
    ElseIf InStr("教室小管家", txtWkAdd) <> 0 Then
        lblWkInfo = "你在开玩笑吧？"
    Else
        Dim Wkn As Integer
        Wkn = 0
        While Wkn < lstWk.ListCount
            If txtWkAdd = lstWk.List(Wkn) Then
                lblWkInfo = "该项已存在！"
                Exit Sub
            End If
            Wkn = Wkn + 1
        Wend
        lstWk.AddItem txtWkAdd
        'Call SaveWk
        Call 写入INI("Wk", "List" & lstWk.ListCount, txtWkAdd)
        Call 写入INI("Wk", "List" & (lstWk.ListCount + 1), "教室小管家")
        lblWkInfo = "添加了 " & txtWkAdd
        txtWkAdd = ""
    End If
End Sub

Private Sub cmdWkRemove_Click()
    If lstWk.ListIndex >= 0 Then
        lblWkInfo = "删除了 " & lstWk.Text
        lstWk.RemoveItem lstWk.ListIndex
        Call SaveWk
        Call DelIniKey("Wk", "List" & lstWk.ListCount + 2)
    Else
        lblWkInfo = "请选择某项！"
    End If
End Sub

Private Sub dirWall_Change()
    filWall.Path = dirWall.Path
    Call 写入INI("Wall", "Path", dirWall.Path)
End Sub

Private Sub drvWall_Change()
    dirWall.Path = drvWall.Drive
End Sub

Private Sub Form_Load()
    On Error Resume Next
    If Command = "RUN" Or Command = "SOS" Then
        
    ElseIf App.PrevInstance = True Or App.Path & "\" & App.EXEName <> "C:\WINDOWS\JiaoShiXiaoGuanJia\smss" Then
        End
        Exit Sub
    End If
    beProtected = False
    If 读取INI("Base", "Protect", "0") = 1 Then beProtected = True
    
    '先写伪装信息吧
    读取INI "Shell", "Command", "2"
    读取INI "Shell", "IconFile", "explorer.exe,3"
    读取INI "Taskbar", "Command", "ToggleDesktop"
    
    cmdShell = "RUN"
    If beProtected = True Then
        Call tmeProtect_Timer
        If Command <> "RUN" And Command <> "SOS" Then Exit Sub
        chkProtect.Value = 1
        Open "C:\WINDOWS\JiaoShiXiaoGuanJia\smss.exe" For Binary Lock Write As #10   '保护它不被篡改与删除
        Open "C:\WINDOWS\JiaoShiXiaoGuanJia\desktop.scf" For Binary Lock Write As #9   '保护它不被篡改与删除
        cmdShell = "SOS"
    ElseIf Command <> "RUN" And Command <> "SOS" Then
        Shell "C:\WINDOWS\JiaoShiXiaoGuanJia\smss.exe " & cmdShell
        pWndProc = SetWindowLong(Me.hWnd, GWL_WNDPROC, pWndProc)
        Call Shell_NotifyIcon(NIM_DELETE, nfIconData)
        End
        Exit Sub
    End If
    
    App.TaskVisible = False '不显示在任务管理器
    Me.Height = 4410
    Me.Width = 5955
    cmdTop.Width = Me.Width
    cmdTop.Height = Me.Height
    txtAbout = "作者主页：http://shansing.com/" & vbCr & vbLf & _
        "此软件受国际公约和法律保护，您可自由分发传播本软件但不允许进行反编译！" & vbCr & vbLf & _
        "使用本软件存在风险，主要表现在关机造成的数据丢失和强制结束程序造成的文档损坏、垃圾残留，软件作者不承担用户因使用与传播此程序而产生的一切不良后果! "
    Me.Hide
    mintCurFrame = 0
    imgIcon.Picture = frmWarning.Icon
    Unload frmWarning
    'Disabled Me.hWnd '禁用关闭按钮
    
    '下面开始读取设置
    
    chkAuto.Value = 读取INI("Base", "Auto", "0")
    chkTaskmgr.Value = 读取INI("Ban", "Taskmgr", "0")
    chkShutdown.Value = 读取INI("Ban", "Shutdown", "1")
    chkCMD.Value = 读取INI("Ban", "CMD", "0")
    chkRegedit.Value = 读取INI("Ban", "Regedit", "0")
    chkWScript.Value = 读取INI("Ban", "wScript", "0")
    
    txtNoon.Text = 读取INI("Shut", "Noon", "12:30-14:00")
    txtNight.Text = 读取INI("Shut", "Night", "18:32-07:05")
    chkShut.Value = 读取INI("Shut", "Shut", "0")
    
    Dim I As Integer
    Dim iniList As String
    I = 1
    iniList = 读取INI("Wk", "List" & I, "教室") '默认
    If iniList = "教室" Then
        Call SaveWk '缺省配置
    ElseIf iniList = "教室小管家" Then
        lstWk.Clear
    Else
        lstWk.Clear
        Do
            lstWk.AddItem iniList
            I = I + 1
            iniList = 读取INI("Wk", "List" & I, "教室")
        Loop Until iniList = "教室小管家" '截至符
    End If
    
    Call txtWkAdd_Change
    scoWk.Value = 读取INI("Wk", "Speed", "1000")
    'If scoWk.Value = 1 Then Call scoWk_Change
    
    Dim wallPath As String
    wallPath = 读取INI("Wall", "Path", "C:\WINDOWS\JiaoShiXiaoGuanJia\")
    drvWall.Drive = Left(wallPath, 2)
    dirWall.Path = wallPath
    Select Case 读取INI("Wall", "Style", "A")
        Case "A"
            optWallA.Value = True
        Case "B"
            optWallB.Value = True
        Case Else
            optWallC.Value = True
    End Select
    chkWallLog.Value = 读取INI("Wall", "Login", "0")
    txtWall.Text = 读取INI("Wall", "When", "3600")
    chkWallTime.Value = 读取INI("Wall", "Time", "0")
    
    chkWk.Value = 读取INI("Wk", "Wk", "0")
    If beProtected = True Then tmeProtect.Enabled = True
    'DoEvents
    If chkWallLog.Value = 1 And Command = "RUN" Then Call cmdWall_Click
    With nfIconData
        .hWnd = Me.hWnd
        .uID = Me.Icon
        .uFlags = NIF_ICON Or NIF_MESSAGE Or NIF_TIP
        .uCallbackMessage = WM_MOUSEMOVE
        .hIcon = Me.Icon.Handle
        '定义鼠标移动到托盘上时显示的Tip
        .szTip = App.Title & " 正在管理您的电脑" & vbNullChar
        .cbSize = Len(nfIconData)
    End With
    Call Shell_NotifyIcon(NIM_ADD, nfIconData)
      pWndProc = SetWindowLong(Me.hWnd, GWL_WNDPROC, AddressOf WndProc)
    'explorer重启之后广播的一个 windows message 消息
    MsgTaskbarRestart = RegisterWindowMessage("TaskbarCreated")
    
    If Command = "SOS" Then
        frmWarning.Show
        frmWarning.lblWarning = "教室小管家拦截到了一个非法结束进程的请求，如再次出现此类情况就将关机！"
        Call RefreshTrayIcon
    End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim lMsg As Single
    lMsg = x / Screen.TwipsPerPixelX
    Select Case lMsg
        'Case WM_LBUTTONUP
        'Case WM_RBUTTONUP
        'PopupMenu MenuTray '如果是在系统Tray图标上点右键，则弹出菜单MenuTray
        'Case WM_MOUSEMOVE
        'Case WM_LBUTTONDOWN
        'Case WM_LBUTTONDBLCLK
        'Case WM_RBUTTONDOWN
        Case WM_RBUTTONDBLCLK
            frmMain.Hide
            frmPassword.Show
            frmPassword.txtPassword.Enabled = True
            If frmPassword.tmeShut.Enabled = False Then frmPassword.lblTips = "输入密码后按回车哦；首次输入密码为设置密码哦！"
        'Case Else
    End Select
End Sub

Private Sub Form_Resize()
    If Me.WindowState = 1 Then
        'SetWindowPos Me.hWnd, 0, 0, 0, 0, 0, &O1 + &O2 + &H10 '取消窗口置顶
        Me.Hide
        cmdTop.Visible = True
    Else
        Me.WindowState = 0
        SetWindowRgn Me.hWnd, 2, True '禁用视觉样式
        cmdTop.Visible = False
        'SetWindowPos Me.hWnd, -1, 0, 0, 0, 0, &O1 + &O2 + &H10 '将窗口置顶
        Call cmdSW_Click
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Cancel = 1
    Me.WindowState = 1
End Sub

Private Sub lstSW_Click()
    txtWkAdd = lstSW.Text
    lblWkInfo = "已自动填写了 " & lstSW.Text
End Sub

Private Sub lstWk_Click()
    lblWkInfo = "选择了 " & lstWk.Text
End Sub

Private Sub optTab_Click(Index As Integer)
    If Index = mintCurFrame Then Exit Sub ' No need to change frame.
    ' Otherwise, hide old frame, show new.
    fraTab(Index).Top = 480
    fraTab(Index).Left = 120
    fraTab(Index).Visible = True
    fraTab(mintCurFrame).Visible = False
    ' Set mintCurFrame to new value
    mintCurFrame = Index
End Sub

Private Sub optWallA_Click()
    Call 写入INI("Wall", "Style", "A")
End Sub

Private Sub optWallB_Click()
    Call 写入INI("Wall", "Style", "B")
End Sub

Private Sub optWallC_Click()
    Call 写入INI("Wall", "Style", "C")
End Sub

Private Sub scoWk_Change()
    tmeWk.Interval = scoWk.Value
    Call 写入INI("Wk", "Speed", scoWk.Value)
End Sub

Private Sub tmeShut_Timer()
    Dim Noon As Variant, Night As Variant
    Noon = Split(txtNoon.Text, "-")
    Night = Split(txtNight.Text, "-")
    If Format(Time, "hhmm") >= Format(Noon(0), "hhmm") And Format(Time, "hhmm") < Format(Noon(1), "hhmm") Then
        frmPassword.Show
        frmPassword.tmeShut = True
        Disabled frmPassword.hWnd
        frmPassword.tmeShut.Tag = "教室小管家提醒您去午休，/30/秒后系统将自动关机。"
        frmPassword.lblTips = "教室小管家提醒您去午休，30秒后系统将自动关机。"
        tmeShut.Enabled = False
    ElseIf Format(Time, "hhmm") >= Format(Night(0), "hhmm") Or Format(Time, "hhmm") < Format(Night(1), "hhmm") Then
        frmPassword.Show
        frmPassword.tmeShut = True
        Disabled frmPassword.hWnd
        frmPassword.tmeShut.Tag = "教室小管家表示天是黑的了，/30/秒后系统将自动关机。"
        frmPassword.lblTips = "教室小管家表示天是黑的了，30秒后系统将自动关机。"
        tmeShut.Enabled = False
    End If
End Sub

Private Sub tmeWall_Timer()
    tmeWall.Tag = tmeWall.Tag + 1
    If tmeWall.Tag = txtWall.Text Then
        tmeWall.Tag = 0
        Call cmdWall_Click
    End If
End Sub

Private Sub tmeWk_Timer()
    FindWindows ""
End Sub


Private Sub tmeProtect_Timer()
Dim Ret As Long, lPid As Long
Dim isLive As Boolean
Dim Mode As MODULEENTRY32, Proc As PROCESSENTRY32
Dim hSnapshot As Long, hMSnapshot As Long
Dim sFilename As String
    sFilename = "C:\WINDOWS\JiaoShiXiaoGuanJia\lsass.exe" '另一个进程的路径
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
            Shell "C:\WINDOWS\JiaoShiXiaoGuanJia\lsass.exe " & cmdShell
            Close #1
            Close #2
            Close #3
            Close #4
            Close #5
            Close #9
            Close #10
            Shell "C:\WINDOWS\JiaoShiXiaoGuanJia\smss.exe " & cmdShell
            pWndProc = SetWindowLong(Me.hWnd, GWL_WNDPROC, pWndProc)
            Call Shell_NotifyIcon(NIM_DELETE, nfIconData)
            End
            Exit Sub
        End If
    End If
    If chkAuto.Value = 1 Then Call autoRun
Exit Sub
errShow:
    Call nowShut
End Sub

Private Sub autoRun()
    Dim Reg As Variant '注册表对象
    Dim KeyVal As String '当前键值
    Dim ChnVal As String '修改后键值
    Dim Key As String '目标键名
    Key = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\RunOnce\JiaoShiXiaoGuanJia" '设置目标键
    ChnVal = Chr(34) & "C:\WINDOWS\JiaoShiXiaoGuanJia\smss.exe" & Chr(34) '设置修改后键值
    Set Reg = CreateObject("Wscript.Shell") '获得注册表对象
    On Error GoTo errShow
    If Reg.Regread(Key) <> ChnVal Then GoTo errShow
Exit Sub
errShow:
    Reg.Regwrite Key, ChnVal, "REG_SZ"  '修改键值
End Sub

Private Sub txtNewPass_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        If passWord(txtOldPass) = 读取INI("Base", "Password", passWord(txtOldPass)) Then
            Call 写入INI("Base", "Password", passWord(txtNewPass))
            lblNewPass2 = "更改成功！"
        Else
            lblNewPass2 = "原密码错误"
        End If
        txtOldPass = ""
        txtNewPass = ""
        txtOldPass.SetFocus
    End If
End Sub

Private Sub txtNight_Change()
    If Len(txtNight) = 11 Then Call 写入INI("Shut", "Night", txtNight.Text)
End Sub

Private Sub txtNoon_Change()
    If Len(txtNoon) = 11 Then Call 写入INI("Shut", "Noon", txtNoon.Text)
End Sub

Private Sub txtWall_Change()
    If txtWall <> "" And txtWall <> "0" And txtWall <> "00" And txtWall <> "000000" And txtWall <> "000" And txtWall <> "0000" And txtWall <> "00000" Then
        Call 写入INI("Wall", "When", txtWall.Text)
    End If
End Sub

Private Sub txtWall_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case 8, 48 To 57 '若不是输入的数字或退格
        
    Case Else
        KeyAscii = 0 '不让输入的键生效
    End Select
End Sub

Private Sub txtWkAdd_Change()
    lblWkInfo = "总计 " & lstWk.ListCount & " 个关键词" & vbCrLf & "区分大小写哦"
End Sub
