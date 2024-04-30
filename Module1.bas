Attribute VB_Name = "Module1"
Option Explicit
Dim lang As String, LocaleID, LangUage%
Private Declare Function GetSystemDefaultLCID Lib "kernel32" () As Long

Sub main()
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
        MsgBox "本軟件是網站「xvannxvan.cn」獨家開發未盜用任何人的外接程序" & Chr(10) & "希望他人也別用我開源的代碼進行盈利互相尊重", vbDefaultButton1, "提示"
    End If
    If LangUage = 2 Then
        MsgBox "本软件是网站“xvannxvan.cn”独家开发未盗用任何人的外接程序" & Chr(10) & "希望他人也别用我开源的代码进行盈利互相尊重", vbDefaultButton1, "提示"
    End If
    If LangUage = 3 Then
        MsgBox "The Software is derived from the “xvanxvan.cn” website" & Chr(10) & "Can't make money with my open source code, respect each other thank you", vbDefaultButton1, "prompt"
    End If
    Form2.Show
End Sub
