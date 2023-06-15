VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_login 
   Caption         =   "Login"
   ClientHeight    =   4425
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   4875
   OleObjectBlob   =   "frm_login.frx":0000
   StartUpPosition =   1  '所屬視窗中央
End
Attribute VB_Name = "frm_login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub CommandButton1_Click()

Dim id As String
Dim password As String

id = Me.tboID
password = Me.tboPASSWORD

Call login_XMLHTTP(id, password)

Unload Me

End Sub

Private Sub CommandButton2_Click()
ActiveWorkbook.FollowHyperlink (Me.Label3.Caption)
End Sub

Sub login_XMLHTTP(ByVal id As String, ByVal password As String)

Dim URL As String
Dim XMLHTTP As Object

Set XMLHTTP = CreateObject("Microsoft.XMLHTTP")
Set DOM = CreateObject("Htmlfile")

URL_tmp = "https://script.google.com/macros/s/AKfycbwuFlRd8e6Ndah-EgG01kpF07dgj6W52-2SZpXHLwfseVAxELibU74FYQ/exec"
Debug.Print "id=" & id
Debug.Print "password=" & password

URL = URL_tmp & "?id=" & id & "&password=" & password & "&ip=" & GetIPAddress & "&wkname=" & ThisWorkbook.Name

With XMLHTTP

    .Open "GET", URL, False
    .send
    
    If .Status = 200 Then
    
        If .responsetext Like "*Login Fail*" Then
            MsgBox "LoginFail"
        Else
            MsgBox "LoginSucces"
        End If
        
    End If
    
End With

End Sub

Function GetIPAddress()
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


