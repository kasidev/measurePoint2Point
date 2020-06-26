Attribute VB_Name = "measurePointToPoint"

Public Sub Testing()

'
' Makro2 Makro
'

'
Dim a As Variant
a = Tabelle1.ChartObjects(1).Chart. _
 SeriesCollection(2).Points(2).Left

Debug.Print a

'test comment 2


End Sub
