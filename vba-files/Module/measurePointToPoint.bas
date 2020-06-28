Attribute VB_Name = "measurePointToPoint"

Public Sub Testing()

'
' Makro2 Makro

'Declare variables
Dim dblLeftPoint As Double
Dim dblTopPoint As Double
Dim dblArrowShift As Double
Dim dblTotalSeries As Double
Dim dblTotalPoints As Double
Dim i_1 As Double
Dim i_2 As Double
Dim dblXCorStart As Double
Dim dblYCorStart As Double
Dim dblXCorEnd As Double
Dim dblYCorEnd As Double
Dim dblXCorChart As Double
Dim dblYCorChart As Double
Dim dblXshift As Double
Dim dblYshift As Double
'
    Tabelle1.ChartObjects(1).Chart. _
        SeriesCollection(2).Points(2).MarkerSize = 50
    '

    dblLeftPoint = Tabelle1.ChartObjects(1).Chart. _
        SeriesCollection(2).Points(2).Left

    dblLeftPoint= dblLeftPoint + Tabelle1.ChartObjects(1).Chart. _
        ChartArea.Left

    Debug.Print dblLeftPoint


    dblTopPoint = Tabelle1.ChartObjects(1).Chart. _
        SeriesCollection(2).Points(2).Top

    dblTopPoint= dblTopPoint + Tabelle1.ChartObjects(1).Chart. _
        ChartArea.Top

    Debug.Print dblTopPoint

    Tabelle1.Shapes.AddConnector _
            Type:=msoConnectorStraight, _
            BeginX:=0, BeginY:=0, _
            EndX:=dblLeftPoint, EndY:=dblTopPoint

    Debug.Print "point count" ; Tabelle1.ChartObjects(1).Chart. _
        SeriesCollection(2).Points.Count

    Debug.Print "collection" ; Tabelle1.ChartObjects(1).Chart. _
        SeriesCollection.Count

'Connector drawing loop

'define arrow shift
dblArrowShift=10

'get the x and y co-ordinates of the upperleft corner of the chart
dblXCorChart =Tabelle1.ChartObjects(1).Chart. _
        ChartArea.Left
dblYCorChart = Tabelle1.ChartObjects(1).Chart. _
        ChartArea.Top

'count the amount of data series
dblTotalSeries=Tabelle1.ChartObjects(1).Chart. _
    SeriesCollection.Count

'iterate through all data series
For i_1= 2 to dblTotalSeries
    'get the number of points in each data series
    dblTotalPoints = Tabelle1.ChartObjects(1).Chart. _
        SeriesCollection(i_1).Points.Count
    
    'iterate through all the point pairs
    For i_2= 1 to dblTotalPoints-1
    'Get the x and y co-ordinate in points of the start point
    dblXCorStart = dblXCorChart +  Tabelle1.ChartObjects(1).Chart. _
        SeriesCollection(i_1).Points(i_2).Left

    dblYCorStart = dblYCorChart + Tabelle1.ChartObjects(1).Chart. _
        SeriesCollection(i_1).Points(i_2).Top

    'Get the x and y co-ordinate in points of the end point
    dblXCorEnd = dblXCorChart +  Tabelle1.ChartObjects(1).Chart. _
        SeriesCollection(i_1).Points(i_2+1).Left

    dblYCorEnd = dblYCorChart + Tabelle1.ChartObjects(1).Chart. _
        SeriesCollection(i_1).Points(i_2+1).Top

    'decide on to which sides the arrow shall be shifted
    Select Case i_1

        'shift up
        Case 2
        dblXshift= 0
        dblYshift= -1*dblArrowShift

        'shift down
        Case 3
        dblXshift= 0
        dblYshift= 1*dblArrowShift

        'shift right
        Case 4
        dblXshift= -1* dblArrowShift
        dblYshift= 0

        'shift left
        Case 5
        dblXshift= 1* dblArrowShift
        dblYshift= 0

        Case Else
        dblXshift= dblArrowShift
        dblYshift= dblArrowShift

    End Select
     Debug.Print Tabelle1.Shapes.Count
    Debug.Print Tabelle1.ChartObjects(1).Chart. _
        SeriesCollection(i_1).Points(i_2).MarkerForegroundColor
    Tabelle1.Shapes.AddConnector _
            Type:=msoConnectorStraight, _
            BeginX:=dblXCorStart+dblXshift, BeginY:=dblYCorStart+dblYshift, _
            EndX:=dblXCorEnd+dblXshift, EndY:=dblYCorEnd+dblYshift

    Debug.Print Tabelle1.Shapes.Count



    Next i_2 
Next i_1

End Sub

