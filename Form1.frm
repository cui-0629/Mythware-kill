VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "�رռ���"
   ClientHeight    =   5040
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10215
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5040
   ScaleWidth      =   10215
   StartUpPosition =   3  '����ȱʡ
   Begin MSComDlg.CommonDialog win1 
      Left            =   8520
      Top             =   4080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Height          =   2295
      Left            =   360
      TabIndex        =   1
      Top             =   240
      Width           =   9495
      Begin VB.CommandButton Command2 
         Caption         =   "ѡ��λ��"
         Height          =   495
         Left            =   960
         TabIndex        =   5
         Top             =   960
         Width           =   1095
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "����"
            Size            =   30
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   2880
         TabIndex        =   2
         Top             =   960
         Width           =   6375
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "����Ҳ�������ĸ�Ŀ¼ѡ�������ݷ�ʽҲ�ǿ��Ե�"
         BeginProperty Font 
            Name            =   "����"
            Size            =   15
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   4
         Top             =   1680
         Width           =   9015
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "�м�ǧ��Ҫѡ����ļ����߿��ܻ���ɲ��ɹ����ĺ��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   15
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   3
         Top             =   240
         Width           =   9015
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "����ر�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   50.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   3000
      TabIndex        =   0
      Top             =   2760
      Width           =   4455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim H1#, W1#, Pp2$, Pp3$

Private Sub Command1_Click()
    Dim Start1$, PZ, PZ1$, PZ2$, PZ3$, PZ4$
    Rem ���������ļ���Ŀ¼��Ŀ¼1
    PZ = Environ("UserProfile")
    Rem ��ȡ����·��
    PZ2 = Environ("UserProfile") & "\Desktop\unlond.bat"
    'Text2.Text = PZ2
    Rem �����û������������ļ���Ŀ¼
    Start1 = Environ("UserProfile") & "\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup"
    Rem---------------------------------------------------------------------------------
    Rem ��ֹ�������
    PZ1 = "taskkill /f /t /im" & Space(1) & Chr(34) & Pp3 & Chr(34)
    PZ3 = Environ("UserProfile") & "\Desktop\kill.bat"
    Open PZ3 For Output As #1
        Print #1, PZ1
    Close
    Call Shell(PZ3)
    Rem Ŀ����ɾ���û�����������������ļ��з�ֹ����Ŀ�������
    PZ1 = "rd/s/q" & Space(1) & Chr(34) & Start1 & Chr(34)
    Open PZ2 For Output As #1
        Print #1, PZ1
    Close
    Call Shell(PZ2)
    Rem ɾ���������������ļ�
    PZ1 = "rd/s/q" & Space(1) & Chr(34) & Pp2 & Chr(34)
    PZ3 = Environ("UserProfile") & "\Desktop\dele.bat"
    Open PZ3 For Output As #1
        Print #1, PZ1
    Close
    Call Shell(PZ3)
End Sub

Private Sub Command2_Click()
    FileOpen
End Sub

Private Sub Form_Load()
    Rem �Ѵ������
    Dim X1#, Y1#
    Left = (Screen.Width - Me.Width) / 2
    Top = (Screen.Height - Me.Height) / 2
    Rem ��ȡ����ĳ�ʼ���
     Me.Height = 5460
     Me.Width = 10305
    Command1.Enabled = False
End Sub

Private Sub Form_Resize()
    Rem ǿ��ȡ���û��ı䴰���С
    'Me.Height = H1
    'Me.Width = W1
End Sub

Sub FileOpen()
    Dim Pp$, i2%
    Pp = ""
    i2 = 1
    win1.CancelError = False
    win1.ShowOpen
    Pp = win1.FileName
    'MsgBox Pp
    If Pp <> "" Then
        For i2 = Len(Pp) To 1 Step -1
            If Mid(Pp, i2, 1) = "\" Then
                Exit For
            End If
        Next i2
        Text2.Text = Left(Pp, i2 - 1)
        Pp2 = Text2.Text
        'MsgBox Pp2
        Pp3 = Right(Pp, Len(Pp) - i2)
        'MsgBox Pp3
        If Text2.Text <> "" Then
            Command1.Enabled = True
        End If
    End If
End Sub

