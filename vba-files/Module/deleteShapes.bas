Public Sub deleteAllShapes()
Dim i_shape As Integer
Dim dblTotalShapes as Double

dblTotalShapes=Tabelle1.Shapes.Count

If dblTotalShapes>3 Then

    For i_shape = dblTotalShapes To 4 Step -1
            Tabelle1.Shapes(i_shape).Delete
    next i_shape
        
End If


End Sub