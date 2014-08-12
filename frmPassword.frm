VERSION 5.00
Begin VB.Form frmPassword 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "教室小管家"
   ClientHeight    =   1485
   ClientLeft      =   11145
   ClientTop       =   9495
   ClientWidth     =   4110
   ForeColor       =   &H00000000&
   Icon            =   "frmPassword.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1485
   ScaleWidth      =   4110
   ShowInTaskbar   =   0   'False
   Begin VB.Timer tmeTop 
      Interval        =   100
      Left            =   0
      Top             =   0
   End
   Begin VB.Timer tmeShut 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   0
      Tag             =   "教室小管家未能拦截某程序，/31/秒后系统将自动关机。"
      Top             =   840
   End
   Begin VB.TextBox txtPassword 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00FF0000&
      Height          =   270
      IMEMode         =   3  'DISABLE
      Left            =   1800
      MaxLength       =   16
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   1200
      Width           =   1575
   End
   Begin VB.Image imgIcon 
      Height          =   960
      Left            =   120
      Top             =   120
      Width           =   960
   End
   Begin VB.Label lblTips 
      BackStyle       =   0  'Transparent
      Caption         =   "教室小管家未能拦截某程序，30秒后系统将自动关机。"
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1095
      Left            =   1160
      TabIndex        =   2
      Top             =   75
      Width           =   2895
   End
   Begin VB.Label lblPassword 
      BackStyle       =   0  'Transparent
      Caption         =   "是管理员么？ 输错三次了喂！！"
      ForeColor       =   &H00808000&
      Height          =   255
      Left            =   720
      TabIndex        =   1
      Top             =   1260
      Width           =   2655
   End
End
Attribute VB_Name = "frmPassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public wrongNo As Integer

'右下角
Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As Any, ByVal fuWinIni As Long) As Long
Private Const SPI_GETWORKAREA = 48
Private Type RECT
Left As Long
Top As Long
Right As Long
Bottom As Long
End Type



Private Sub Form_Load()
    imgIcon.Picture = frmMain.imgIcon
    wrongNo = 0
    SetWindowRgn Me.hWnd, 2, True '禁用视觉样式
    
    '移到右下角
    Dim lRes As Long
    Dim rectVal As RECT
    Dim TaskbarHeight As Integer
    lRes = SystemParametersInfo(SPI_GETWORKAREA, 0, rectVal, 0)
    TaskbarHeight = Screen.Height - rectVal.Bottom * Screen.TwipsPerPixelY
    If TaskbarHeight = 0 Then TaskbarHeight = 450
    Me.Move Screen.Width - Me.Width, Screen.Height - Me.Height - TaskbarHeight, Me.Width, Me.Height
End Sub

Private Sub Form_Unload(Cancel As Integer)
    tmeTop.Enabled = False
    SetWindowPos Me.hWnd, 0, 0, 0, 0, 0, &O1 + &O2 + &H10 '取消窗口置顶
    If tmeShut.Enabled = True Then
        Cancel = 1
        Call tmeShut_Timer
    End If
End Sub

Private Sub tmeShut_Timer()
    Dim warnTxt As Variant
    warnTxt = Split(tmeShut.Tag, "/")
    If warnTxt(1) = 1 Then
        Call nowShut
        frmMain.tmeShut.Enabled = True
        tmeShut.Enabled = False
    Else
        lblTips.Caption = warnTxt(0) & (warnTxt(1) - 1) & warnTxt(2)
        tmeShut.Tag = warnTxt(0) & "/" & (warnTxt(1) - 1) & "/" & warnTxt(2)
    End If
End Sub

Private Sub tmeTop_Timer()
    SetWindowPos Me.hWnd, -1, 0, 0, 0, 0, &O1 + &O2 + &H10 '将窗口置顶
End Sub

Private Sub txtPassword_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        If passWord(txtPassword) = 读取INI("Base", "Password", passWord(txtPassword)) Then
            If tmeShut.Enabled = True Then frmMain.lblShut.Caption = "本次关机已被管理员取消。"
            tmeShut.Enabled = False
            Unload Me
            frmMain.WindowState = 0
            frmMain.Show
        Else
            wrongNo = wrongNo + 1
            If wrongNo = 3 Then txtPassword.Visible = False
        End If
        txtPassword = ""
    End If
End Sub

