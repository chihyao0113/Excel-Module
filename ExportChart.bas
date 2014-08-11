Attribute VB_Name = "ExportChart"
Sub ExportActiveChart()
'Debug.Print ActiveWorkbook.path
'Debug.Print ActiveChart.ChartTitle.Text

ActiveChart.Export ActiveWorkbook.path & "\" & ActiveChart.ChartTitle.Text & ".png"
End Sub

Sub ExportSelectedCharts()
    Dim i As Integer
    For i = 1 To Selection.Count
        With Selection(i) 'assumed to be a chart object; could verify with TypeName
            With .Chart
                .Export ActiveWorkbook.path & "\" & .ChartTitle.Text & ".png"
            End With
             'real code here
        End With
    Next i
End Su
