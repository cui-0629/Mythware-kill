VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   0  'None
   ClientHeight    =   5670
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9615
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5670
   ScaleWidth      =   9615
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Enabled         =   0   'False
      Height          =   495
      Left            =   6480
      TabIndex        =   2
      Top             =   4680
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   600
      Top             =   1200
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "s"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   20.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4320
      TabIndex        =   5
      Top             =   4680
      Width           =   735
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "10"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   20.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3600
      TabIndex        =   4
      Top             =   4680
      Width           =   735
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Label3"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   20.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1560
      TabIndex        =   3
      Top             =   4680
      Width           =   1935
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   20.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   480
      TabIndex        =   1
      Top             =   2400
      Width           =   8535
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   39.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   480
      TabIndex        =   0
      Top             =   600
      Width           =   8415
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim lang As String, LocaleID, LangUage%, JS%
Private Declare Function GetSystemDefaultLCID Lib "kernel32" () As Long

Private Sub Command1_Click()
    Form1.Show
    Unload Form2
End Sub

Private Sub Form_Load()
    Rem 把窗体居中
    Dim X1#, Y1#
    Left = (Screen.Width - Me.Width) / 2
    Top = (Screen.Height - Me.Height) / 2
    Dim LocaleID As Long
    LocaleID = GetSystemDefaultLCID
    Select Case LocaleID
    Case &H404
    LangUage = 1 '中文繁体
    Case &H804
    LangUage = 2 '中文简体
    lang = "1"
    Case &H409
    LangUage = 3 '英文
    lang = "2"
    End Select
    If LangUage = 1 Then
        Me.Caption = "欢迎"
        Label1.Caption = "欢迎"
        Label2.Caption = "本软件只用于研究和学习使用若用于不正当用途本软件创造者概不负责"
        Label3.Caption = "剩余"
        Command1.Caption = "跳过"
    End If
    If LangUage = 2 Then
        Me.Caption = "欢迎"
        Label1.Caption = "欢迎"
        Label2.Caption = "本软件只用于研究和学习使用若用于不正当用途本软件创造者概不负责"
        Label3.Caption = "剩余"
        Command1.Caption = "跳过"
    End If
    If LangUage = 3 Then
        Me.Caption = "Welcome"
        Label1.Caption = "Welcome"
        Label2.Caption = "This software is only used for research and study purposes"
        Label3.Caption = "rest"
        Command1.Caption = "skip"
    End If
    JS = 10
    Me.Height = 5670
    Me.Width = 9615
End Sub

Private Sub Timer1_Timer()
    JS = JS - 1
    If JS = 8 Then Command1.Enabled = True
    Label4.Caption = JS
    If JS = 0 Then
        Form1.Show
        Unload Form2
    End If
End Sub
