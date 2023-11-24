Attribute VB_Name = "Module2"
Function LinearInterpolationError(xValue As Double, xRange As Range, yRange As Range, errorX As Double) As Double
    ' Find the error of a linear interpolation given a known set of x and y values
    
    ' Check if the xValue is within the range of x values
    If xValue < WorksheetFunction.Min(xRange) Or xValue > WorksheetFunction.Max(xRange) Then
            LinearInterpolationError = CVErr(xlErrValue)  ' Return an error if xValue is outside the range
            Exit Function
    End If
    
    ' Find the index of the closest x value in the range
    Dim i As Long
    For i = 1 To xRange.Rows.Count
        If xRange.Cells(i, 1).Value >= xValue Then
            Exit For
        End If
    Next i
    
    ' Perform linear interpolation
    Dim x0 As Double, x1 As Double, y0 As Double, y1 As Double
    x0 = xRange.Cells(i - 1, 1).Value
    x1 = xRange.Cells(i, 1).Value
    y0 = yRange.Cells(i - 1, 1).Value
    y1 = yRange.Cells(i, 1).Value
    
    LinearInterpolationError = (y1 - y0) * errorX / (x1 - x0)
    
End Function
