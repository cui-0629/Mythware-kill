Attribute VB_Name = "Module1"
Option Explicit
Dim lang As String, LocaleID, LangUage%
Private Declare Function GetSystemDefaultLCID Lib "kernel32" () As Long

Sub main()
Dim LocaleID As Long
    LocaleID = GetSystemDefaultLCID
    Select Case LocaleID
    Case &H404
    LangUage = 1 '���ķ���
    Case &H804
    LangUage = 2 '���ļ���
    lang = "1"
    Case &H409
    LangUage = 3 'Ӣ��
    lang = "2"
    End Select
    If LangUage = 1 Then
        MsgBox "��ܛ���ǾWվ��xvannxvan.cn�������_�lδ�I���κ��˵���ӳ���" & Chr(10) & "ϣ������Ҳ�e�����_Դ�Ĵ��a�M��ӯ����������", vbDefaultButton1, "��ʾ"
    End If
    If LangUage = 2 Then
        MsgBox "���������վ��xvannxvan.cn�����ҿ���δ�����κ��˵���ӳ���" & Chr(10) & "ϣ������Ҳ�����ҿ�Դ�Ĵ������ӯ����������", vbDefaultButton1, "��ʾ"
    End If
    If LangUage = 3 Then
        MsgBox "The Software is derived from the ��xvanxvan.cn�� website" & Chr(10) & "Can't make money with my open source code, respect each other thank you", vbDefaultButton1, "prompt"
    End If
    Form2.Show
End Sub
