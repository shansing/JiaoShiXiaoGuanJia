VERSION 5.00
Begin VB.Form frmWarning 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "����С�ܼ�"
   ClientHeight    =   1740
   ClientLeft      =   10680
   ClientTop       =   9495
   ClientWidth     =   4830
   Icon            =   "frmWarning.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmWarning.frx":3332
   ScaleHeight     =   1740
   ScaleWidth      =   4830
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton cmdOK 
      Caption         =   "��������"
      Default         =   -1  'True
      Height          =   375
      Left            =   3120
      TabIndex        =   1
      Top             =   1320
      Width           =   1455
   End
   Begin VB.Label lblTips 
      BackStyle       =   0  'Transparent
      Caption         =   "��ʾ"
      BeginProperty Font 
         Name            =   "����"
         Size            =   21.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000011&
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   1250
      Width           =   1815
   End
   Begin VB.Image imgIcon 
      Height          =   960
      Left            =   120
      Top             =   120
      Width           =   960
   End
   Begin VB.Label lblWarning 
      BackStyle       =   0  'Transparent
      Caption         =   "����С�ܼ�����ֹ�˱���ֹ��������У��������⾴����ϵ���Թ���Ա��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1455
      Left            =   1160
      TabIndex        =   0
      Top             =   120
      Width           =   3615
   End
End
Attribute VB_Name = "frmWarning"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOK_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    imgIcon.Picture = frmMain.imgIcon
    SetWindowPos Me.hWnd, -1, 0, 0, 0, 0, &O1 + &O2 + &H10 '�������ö�
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SetWindowPos Me.hWnd, 0, 0, 0, 0, 0, &O1 + &O2 + &H10 'ȡ�������ö�
End Sub
