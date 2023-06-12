VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsFetchURL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'HankLin edited @ 20210823
'使用雲端授權結合Google app script當作後臺控制
'應用xlpadlock編譯xls變成exe檔案即可達到小白無法侵入的等級


Dim XMLHTTP As Object
Dim DOM As Object
'Dim clientObj As New clsUserInformation

Const URL As String = "https://script.google.com/macros/s/AKfycbw2QjF7LNZXPW9nLglW2yGpPK5Hme4NSK8Q8whivNw9VVV5kPcOFvYzu5DEtKN-zVlWSg/exec"

Sub Class_Initialize()

Set XMLHTTP = CreateObject("Microsoft.XMLHTTP")
Set DOM = CreateObject("Htmlfile")

End Sub

Function CreateURL(ByVal mode As String, Optional ByVal user_name As String, Optional ByVal user_company As String, Optional ByVal user_mail As String, Optional ByVal msg As String) ', ByVal myName As String)

'MAC_ADDRESS
'MY_NAME
'MY_COMPANY
'MY_MAIL

user_mac = getMacAddress ' getUnicode(clientObj.getMacAddress)
user_name = getUnicode(user_name)
user_company = getUnicode(user_company)
'user_mail = user_mail
msg = getUnicode(msg)

  'let mode=e.parameter.mode; //方法選擇
  'let mac=deUnicode(e.parameter.mac); //帳號
  'let myname=deUnicode(e.parameter.myname); //密碼
  'let workgroup=deUnicode(e.parameter.workgroup); //機關名
  'let mymail=deUnicode(e.parameter.mymail);//電子郵件

CreateURL = URL & "?func=" & mode & _
                    "&mac=" & user_mac & _
                    "&myname=" & user_name & _
                    "&workgroup=" & user_company & _
                    "&mymail=" & user_mail & _
                    "&msg=" & msg
                    
End Function

Function ExecHTTP(ByVal myURL As String)

On Error GoTo ERRORHANDLE

With XMLHTTP

    .Open "GET", myURL, False
    .send
    
    'Debug.Print myURL
    
    If .Status = 200 Then
    
        tmp = Split(.responsetext, "***")
        ExecHTTP = tmp(1)
        
    Else
ERRORHANDLE:
        MsgBox "你沒有連線!", vbCritical
        
        clientObj.hideCmd
        
        'ThisWorkbook.Close Savechanges:=False

    End If
    
End With

End Function

Sub signup(ByVal id As String, ByVal password As String, ByVal wg As String, ByVal myName As String)

'myURL = CreateURL("1", id, password, wg, "123", myName)


If ExecHTTP(myURL) = "signed" Then
    MsgBox "帳號或者該電腦已經被註冊過了!", vbCritical
    
Else
    MsgBox "註冊成功，請務必牢記帳號!!", vbInformation
End If

End Sub

Function checkAccessByIP()

myURL = CreateURL("2", "test", "123", "wg", "123", "test")

tmp = Split(ExecHTTP(myURL), ":")

If tmp(0) <> "Pass" Then
    MsgBox "IP驗證失敗!!請登入帳號密碼", vbCritical
    checkAccessByIP = False
    Exit Function
Else
    'Application.StatusBar = "設計者：林宗漢(Hank Lin) @2021," & tmp(3) & "使用者：" & tmp(2)
    checkAccessByIP = True
    frm_login.Label6 = "Pass"
End If

If tmp(1) <> "" Then
    MsgBox "系統公告>>" & vbNewLine & tmp(1), vbInformation
    checkAccessByIP = True
End If

End Function

Sub checkAccesByID(ByVal id As String, ByVal password As String)

myURL = CreateURL("3", id, password, "wg", "123", "123")

tmp = Split(ExecHTTP(myURL), ":")

If tmp(0) <> "Pass" Then
    MsgBox "驗證失敗!!", vbCritical: ThisWorkbook.Close SaveChanges:=False
Else
    'Application.StatusBar = "設計者：林宗漢(Hank Lin) @2021," & tmp(3) & "使用者：" & tmp(2)
    frm_login.Label6 = "Pass"
End If

If tmp(1) <> "" Then
    MsgBox "系統公告>>" & vbNewLine & tmp(1), vbInformation
End If

End Sub

Sub getReport(ByVal prompt)

myURL = CreateURL("4", "test", "1234", "wg", prompt, "123")

If ExecHTTP(myURL) = "ok" Then
    MsgBox "回饋意見已發送給作者!", vbInformation
Else
    MsgBox "此電腦未進行註冊!!請透過Gmail發送意見", vbCritical
End If

End Sub

Sub getPassword(ByVal id As String)

myURL = CreateURL("5", id, "1234", "wg", "123", "123")

s = ExecHTTP(myURL)

If s <> "undefined" Then
    MsgBox "您的密碼為" & vbNewLine & s, vbInformation
Else
    MsgBox "找不到密碼，請聯絡作者!", vbCritical
End If

End Sub

'============method===============

Private Function GetIPAddress()
        Const strComputer As String = "."   ' Computer name. Dot means local computer
        Dim objWMIService, IPConfigSet, IPConfig, IPAddress, i
        Dim strIPAddress As String

        ' Connect to the WMI service
        Set objWMIService = GetObject("winmgmts:" _
            & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

        ' Get all TCP/IP-enabled network adapters
        Set IPConfigSet = objWMIService.ExecQuery _
            ("Select * from Win32_NetworkAdapterConfiguration Where IPEnabled=TRUE")

        ' Get all IP addresses associated with these adapters
        For Each IPConfig In IPConfigSet
            IPAddress = IPConfig.IPAddress
            If Not IsNull(IPAddress) Then
                If InStr(1, IPConfig.Description, "WAN (", vbTextCompare) Then
                   MsgBox "網頁 IP = " + IPAddress(0)
                End If
                strIPAddress = strIPAddress & Join(IPAddress, "/") + vbCrLf
            End If
        Next

        GetIPAddress = strIPAddress

        'MsgBox strIPAddress
    End Function
    
Private Function getUnicode(s)

For i = 1 To Len(s)

    myUnicode = myUnicode & "," & CharToUnicode(mid(s, i, 1))

Next

getUnicode = mid(myUnicode, 2)

End Function

    
Private Function CharToUnicode(strChar As String)

    'https://stackoverflow.com/questions/4006246/getting-the-unicode-value-of-a-char-in-vb/34780532

    Dim lngUnicode As Long

    lngUnicode = AscW(strChar)

    If lngUnicode < 0 Then
        lngUnicode = 65536 + lngUnicode
    End If

    CharToUnicode = lngUnicode
    
End Function

Function getMacAddress()

Dim objVMI As Object
Dim vAdptr As Variant
Dim objAdptr As Object
'Dim adptrCnt As Long


Set objVMI = GetObject("winmgmts:\\" & "." & "\root\cimv2")
Set vAdptr = objVMI.ExecQuery("SELECT * FROM Win32_NetworkAdapterConfiguration WHERE IPEnabled = True")

For Each objAdptr In vAdptr
    If Not IsNull(objAdptr.MACAddress) And IsArray(objAdptr.IPAddress) Then
        For adptrCnt = 0 To UBound(objAdptr.IPAddress)
        If Not objAdptr.IPAddress(adptrCnt) = "0.0.0.0" Then
            GetNetworkConnectionMACAddress = objAdptr.MACAddress
            Exit For
        End If
        Next
    End If
Next

getMacAddress = GetNetworkConnectionMACAddress

End Function

Sub Login(ByVal mac_add As String)

If checkIsExist(mac_add) = False Then
    
    Names("UserName").Value = mac_add
    
    MsgBox "初始化成功，歡迎使用!", vbInformation
    
    ans = ExecHTTP(CreateURL("Sign")) ', InputBox("輸入姓名"), InputBox("輸入機關名稱"), InputBox("輸入電子郵件")))

End If

msg = ExecHTTP(CreateURL("Access"))

If msg <> "PASS" Then

    ERRORForm.Show

    'ThisWorkbook.Close False

End If

End Sub

Function checkIsExist(ByVal mac_add As String)

recMac = Names("UserName").Value
newMac = mid(recMac, 3, Len(recMac) - 3)

checkIsExist = False

If newMac = mac_add Then checkIsExist = True

End Function

