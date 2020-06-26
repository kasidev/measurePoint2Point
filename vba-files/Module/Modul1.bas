Attribute VB_Name = "Modul1"
Option Explicit

Sub Makro1()
Attribute Makro1.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Makro1 Makro
'

'
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 370.5, 135, 374.25, _
        225.75).Select
        Selection.ShapeRange.Line.BeginArrowheadStyle = msoArrowheadTriangle
        Selection.ShapeRange.Line.EndArrowheadStyle = msoArrowheadTriangle
End Sub
Sub Makro2()
Attribute Makro2.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Makro2 Makro
'

'
Dim a As Variant
a = Tabelle1.ChartObjects(1).Chart. _
 SeriesCollection(2).Points(2).Left

Debug.Print a

End Sub
