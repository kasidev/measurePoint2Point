Public Sub drawArrows()
'Declare all variables
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
Dim intLineColorR As Integer
Dim intLineColorG As Integer
Dim intLineColorB As Integer

'Connector drawing loop

'define arrow shift
dblArrowShift=20

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
        dblXshift= 0+3
        dblYshift= -1*dblArrowShift

        'shift down
        Case 3
        dblXshift= 0+3
        dblYshift= 1*dblArrowShift

        'shift right
        Case 4
        dblXshift= -1* dblArrowShift
        dblYshift= 0+3

        'shift left
        Case 5
        dblXshift= 1* dblArrowShift
        dblYshift= 0+3

        Case Else
        dblXshift= dblArrowShift
        dblYshift= dblArrowShift

    End Select

    Tabelle1.Shapes.AddConnector _
            Type:=msoConnectorStraight, _
            BeginX:=dblXCorStart+dblXshift, BeginY:=dblYCorStart+dblYshift, _
            EndX:=dblXCorEnd+dblXshift, EndY:=dblYCorEnd+dblYshift

    Tabelle1.Shapes(Tabelle1.Shapes.Count).Select
    With Selection.ShapeRange.Line
        .ForeColor.RGB = RGB(0, 0, 0)
        .EndArrowheadStyle = msoArrowheadOpen
        .BeginArrowheadStyle = msoArrowheadOpen
    End With


    Next i_2 
Next i_1

Call addTextBox()
End Sub