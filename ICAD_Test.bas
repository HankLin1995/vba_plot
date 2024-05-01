Attribute VB_Name = "ICAD_Test"
'Author:HankLin
'Date:2021/5/19
'Workspace:YLIA

Sub test_openlayer()

Dim o As New clsHsec_CAD

Call o.openLayers("") '平面圖-中心線,平面圖-橫斷樁L")

End Sub

Sub test_layer2() '只能傳遞英文字母

Dim objLayer As IntelliCAD.Layer
Dim strLayerName As String

strLayerName = "DIM"

'Set objLayer = test_layer("平面圖-橫斷樁")

Set objLayer = IntelliCAD.ActiveDocument.Layers(strLayerName)
If objLayer Is Nothing Then
Else
IntelliCAD.ActiveDocument.ActiveLayer = objLayer
End If

End Sub


Function test_layer(ByVal layername As String)

Dim myLayer As Layer
Dim mylayers As Layers
Set mylayers = IntelliCAD.ActiveDocument.Layers

For Each ly In mylayers

    Set myLayer = ly
    
    '不能使用myLayers.item("平面圖-橫斷樁")做回傳

    Set IntelliCAD.ActiveDocument.ActiveLayer = myLayer

Next

'Set myLayer = mylayers.Item("平面圖-橫斷樁")

End Function

'*****BUG*****

Sub test_hatch()

Set o = mo.AddHatch(0, PatternName, True)

End Sub

Sub test_ssetFilter()

Set o = CAD.CreateSSET(, "8", "橫斷面樁")

For Each it In o

    Debug.Print it.Layer

Next

End Sub

'-----PASS-----

Sub test_addtextbox()

Dim CAD As New clsACAD
Dim txtpt(2) As Double

txtpt(0) = 100
txtpt(1) = 100
txtpt(2) = 0

Set txtobj = CAD.AddText("test", txtpt, 100, 2)
Call CAD.AddTextBox(txtobj, 2)

End Sub

Sub test_addctext()

Dim CAD As New clsACAD

Dim txtpt(2) As Double

Call CAD.AddCText("t", txtpt, 5)

End Sub

Sub test_addmixtext()

Dim txtpt(2) As Double
txtpt(0) = 0
txtpt(1) = 0
txtpt(2) = 0

Call CAD.AddMixText("test", txtpt, 5, 1, 1)

End Sub
Sub test_getstring()

Dim CAD As New clsACAD

Debug.Print CAD.GetString("test")

End Sub

Sub test_polyline()

Dim CAD As New clsACAD

Dim Ldpt(2) As Double
Dim rupt(2) As Double

rupt(0) = 100
rupt(1) = 100

'Set plobj = CAD.PlotRec(Ldpt, rupt)
Set PLobj = CAD.PlotRecFillet(Ldpt, rupt, 10)

End Sub

Sub test_getPoint()

Dim CAD As New clsACAD

o = CAD.GetPoint("test")

End Sub

Sub test_addtext()

Dim CAD As New clsACAD
Dim txtpt(2) As Double
Dim txtobj As Object

txtpt(0) = 100
txtpt(1) = 100

Set txtobj = CAD.AddText("test", txtpt, 100, 2)

txtobj.rotate CAD.tranPoint(txtpt), 3.14 / 2 '如果是新的pt(object)便不行

Set txtobj = CAD.AddCText("testc", txtpt, 100)

End Sub

Sub test_addline()

Dim CAD As New clsACAD
Dim spt(2) As Double
Dim ept(2) As Double

ept(0) = 100
ept(1) = 200

Set lineObj = CAD.AddLine(spt, ept)

End Sub


Sub test_lwpline()

Dim CAD As New clsACAD

Dim vertices(2 * 2 - 1) As Double

vertices(2) = 100
vertices(3) = 100

CAD.AddLWPolyLine (vertices)

End Sub

Sub test_lineCO()

Dim CAD As New clsACAD
Dim X As Double
Dim Y As Double
Dim r As Double

X = 100
Y = 100
r = 20

Call CAD.AddCircleCO(X, Y, r)
Call CAD.AddPointCO(X, Y)

End Sub

Sub test_addpt()

Dim CAD As New clsACAD

Dim pt(2) As Double

pt(0) = 100
pt(1) = 200

Call CAD.AddPoint(pt)

End Sub

Sub test_layers()

Dim CAD As New clsACAD

Set o = CAD.acadDoc.Layers

For Each L In o

    Debug.Print L.Name

Next

End Sub
