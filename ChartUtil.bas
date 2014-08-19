Private Sub SetXAxisLog()
    Dim x() As Double
    i = 0
    For Each curve In ActiveChart.SeriesCollection
        
        For Each element In curve.XValues
            ReDim Preserve x(i)
            'Debug.Print i
            x(i) = element
            i = i + 1
        Next
    Next
    xMin = Application.Min(x)
    xMax = Application.Max(x)
    
    
    
    If xMax / xMin > 10 Then
        ActiveChart.Axes(xlCategory).ScaleType = xlLogarithmic
        tmp1 = Log(xMin) / Log(10#)
        tmp2 = Int(tmp1)
        tmp3 = 10 ^ tmp2
        ActiveChart.Axes(xlCategory).CrossesAt = tmp3
    End If
    
End Sub
