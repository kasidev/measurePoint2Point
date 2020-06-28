Public Sub addTextBox()
'Declare all variables
Dim dblLeftPoint As Double
Dim dblTopPoint As Double
Dim dblTextShiftHorizontal As Double
Dim dblTextShiftVertical As Double
Dim dblTotalSeries As Double
Dim dblTotalPoints As Double
Dim i_1 As Double
Dim i_2 As Double
Dim i_3 As Double
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
Dim dblXCorText As Double
Dim dblYCorText As Double
Dim dblTextBoxwidth As Double
Dim strTextBoxValue as String
Dim dblTotalRows As Double
Dim dblTextIndex As Double
Dim textOrientation As Variant
Dim dblTextBoxHeigt As Double

'Texbox creation loop

'define shift horizontal
dblTextShiftHorizontal=35

'define shift horizontal
dblTextShiftVertical=35


'count total rows
dblTotalRows=Tabelle1.UsedRange.Rows.Count

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

        'Punkte Oben
        Case 2
        dblTextIndex=0
        i_3=1
        textOrientation=msoTextOrientationHorizontal
        dblTextBoxwidth = 40
        dblTextBoxHeigt=20
        dblXCorText= dblXCorStart+ ((dblXCorEnd-dblXCorStart)/2)-(dblTextBoxwidth/2)+10
        dblYCorText=dblYCorStart-dblTextShiftVertical

        'Value for text box
        Do While dblTextIndex=0
      
            If Tabelle1.Range("C" & i_3).Value = "Punkte Oben" Then
                dblTextIndex=i_3+i_2
            Else
                i_3=i_3+1
            End If

        Loop

        'Punkte unten
        Case 3
        dblTextIndex=0
        i_3=1
        textOrientation=msoTextOrientationHorizontal
        dblTextBoxwidth = 40
        dblTextBoxHeigt=20
        dblXCorText= dblXCorStart+ ((dblXCorEnd-dblXCorStart)/2)-(dblTextBoxwidth/2)+10
        dblYCorText=dblYCorStart+dblTextBoxHeigt

        'Value for text box
        Do While dblTextIndex=0
      
            If Tabelle1.Range("C" & i_3).Value = "Punkte Unten" Then
                dblTextIndex=i_3+i_2
            Else
                i_3=i_3+1
            End If

        Loop


        'Punkte Links
        Case 4
        dblTextIndex=0
        i_3=1
        textOrientation=msoTextOrientationUpward
        dblXCorText= dblXCorStart-dblTextShiftHorizontal-3
        dblTextBoxwidth = 20
        dblTextBoxHeigt=40
        dblYCorText=dblYCorStart-((dblYCorStart-dblYCorEnd)/2)-(dblTextBoxHeigt/2)

        'Value for text box
        Do While dblTextIndex=0
      
            If Tabelle1.Range("C" & i_3).Value = "Punkte Links" Then
                dblTextIndex=i_3+i_2
            Else
                i_3=i_3+1
            End If

        Loop

        'shift left
        Case 5
        dblTextIndex=0
        dblTextBoxwidth = 20
        dblTextBoxHeigt=40
        i_3=1
        textOrientation=msoTextOrientationUpward
        dblXCorText= dblXCorStart+dblTextShiftHorizontal-dblTextBoxwidth
        dblYCorText=dblYCorStart-((dblYCorStart-dblYCorEnd)/2)-(dblTextBoxHeigt/2)

        'Value for text box
        Do While dblTextIndex=0
      
            If Tabelle1.Range("C" & i_3).Value = "Punkte Rechts" Then
                dblTextIndex=i_3+i_2
            Else
                i_3=i_3+1
            End If

        Loop

        Case Else
        dblXshift= dblArrowShift
        dblYshift= dblArrowShift

    End Select

    Tabelle1.Shapes.AddTextbox _
    Orientation:=textOrientation , _
    Left:= dblXCorText, _
    Top:= dblYCorText, _
    Width:= dblTextBoxwidth, _
    Height:= dblTextBoxHeigt


    Tabelle1.Shapes(Tabelle1.Shapes.Count).Select
   'Dim boltest As Variant
    'boltest = Selection
    With Selection.ShapeRange(1)
      .TextFrame2.TextRange.Characters.Text = Tabelle1.Range("F" & dblTextIndex)
      .Line.Visible = msoFalse
      .Fill.Visible = msoFalse
      .TextFrame2.TextRange.Font.Name = "Courier New"
      .TextFrame2.TextRange.Font.Size = 10


    End With

    

    Next i_2 
Next i_1
End Sub