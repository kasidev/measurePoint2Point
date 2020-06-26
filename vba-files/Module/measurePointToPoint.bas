Attribute VB_Name = "measurePointToPoint"

Public Sub Testing()

'
' Makro2 Makro
'
Tabelle1.ChartObjects(1).Chart. _
    SeriesCollection(2).Points(2).MarkerSize = 50
'
Dim dblLeftPoint As Double
dblLeftPoint = Tabelle1.ChartObjects(1).Chart. _
    SeriesCollection(2).Points(2).Left

dblLeftPoint= dblLeftPoint + Tabelle1.ChartObjects(1).Chart. _
    ChartArea.Left

Debug.Print dblLeftPoint

Dim dblTopPoint As Double

dblTopPoint = Tabelle1.ChartObjects(1).Chart. _
    SeriesCollection(2).Points(2).Top

dblTopPoint= dblTopPoint + Tabelle1.ChartObjects(1).Chart. _
    ChartArea.Top

Debug.Print dblTopPoint

Tabelle1.Shapes.AddConnector _
        Type:=msoConnectorStraight, _
        BeginX:=0, BeginY:=0, _
        EndX:=dblLeftPoint, EndY:=dblTopPoint



End Sub
