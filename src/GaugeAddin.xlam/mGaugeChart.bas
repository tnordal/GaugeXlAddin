Attribute VB_Name = "mGaugeChart"
' ------------------------------------------------------
' Name: mGaugeChart
' Kind: Module
' Purpose: Procedures for setting up Gauge Chart
' Author: Tom Nordal
' Date: 2018-08-18
' Version: 0.5
' ------------------------------------------------------
Option Explicit
Option Private Module

Public Sub BasicGaugeStep(Optional ColorRange1 As Long = 255, Optional ColorRange2 As Long = 65535, Optional ColorRange3 As Long = 5287936)
' ----------------------------------------------------------------
' Procedure Name: BasicGaugeStep
' Purpose: Main procedure for making Gauge Chart
' Procedure Kind: Sub
' Procedure Access: Public
' Author: Tom Nordal
' Date: 2018-12-27
' Version: 0.6
' Changed: Add some more styling (shading donut)
' ----------------------------------------------------------------


On Error GoTo BasicGaugeStep_Error

    ActiveSheet.Shapes.AddChart2(201, xlColumnClustered).Select
    ActiveChart.SetSourceData Source:=gChartRangeDonut
    Set gCht = ActiveChart
    
    gCht.ChartArea.Top = gActiveCell.Top
    gCht.ChartArea.Left = gActiveCell.Left
    
    With gCht
        .SetElement (msoElementChartTitleNone)
        .SetElement (msoElementLegendNone)
        
        .SeriesCollection(1).ChartType = xlDoughnut
        .ChartGroups(1).FirstSliceAngle = 270
        .SeriesCollection.Add gChartRangePie            'Added 2018-12-26
        .SeriesCollection(2).ChartType = xlPie
        .SeriesCollection(2).Name = gChartRangePieName
        .SeriesCollection(2).AxisGroup = 2
        .ChartGroups(2).FirstSliceAngle = 270
        
    End With
'Exit Sub
    
    With gCht
        'Hide Doughnut point 4
        .SeriesCollection(1).Points(4).Format.Fill.visible = msoFalse
        .SeriesCollection(1).Points(4).Format.Line.visible = msoFalse
    End With

    With gCht
        'Hide Pie points 1 and 3
        .SeriesCollection(2).Points(1).Format.Fill.visible = msoFalse
        .SeriesCollection(2).Points(1).Format.Line.visible = msoFalse
        .SeriesCollection(2).Points(3).Format.Fill.visible = msoFalse
        .SeriesCollection(2).Points(3).Format.Line.visible = msoFalse
        .ChartGroups(1).FirstSliceAngle = 270
        .ChartGroups(2).FirstSliceAngle = 270
        .SeriesCollection(2).AxisGroup = 2
    End With
    
        
    'Color Needle
    With gCht.SeriesCollection(2).Points(2).Format.Fill
        .visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorText1
        .ForeColor.TintAndShade = 0
        .ForeColor.Brightness = 0
        .Transparency = 0
        .Solid
    End With
    
    gCht.SeriesCollection(2).Points(2).Format.Line.visible = msoFalse
    
    'Color Doughnut first sector
    With gCht.SeriesCollection(1).Points(1).Format.Fill
        .visible = msoTrue
        .ForeColor.RGB = ColorRange1
        .Transparency = 0
        .Solid
    End With

     'Color Doughnut first sector
    With gCht.SeriesCollection(1).Points(1).Format.Glow
        .Color = ColorRange1 'msoThemeColorAccent4 .ObjectThemeColor
        .Color.TintAndShade = 1
        .Color.Brightness = 1
        .Transparency = 0.6
        .Radius = 8
    End With
    
    'Color Doughnut second sector
    With gCht.SeriesCollection(1).Points(2).Format.Fill
        .visible = msoTrue
        .ForeColor.RGB = ColorRange2
        .Transparency = 0
        .Solid
    End With
     
     'Color Doughnut second sector
    With gCht.SeriesCollection(1).Points(2).Format.Glow
        .Color = ColorRange2 'msoThemeColorAccent4 .ObjectThemeColor
        .Color.TintAndShade = 1
        .Color.Brightness = 1
        .Transparency = 0.6
        .Radius = 8
    End With

    'Color Doughnut third sector
    With gCht.SeriesCollection(1).Points(3).Format.Fill
        .visible = msoTrue
        .ForeColor.RGB = ColorRange3
        .Transparency = 0
        .Solid
    End With
    
      'Color Doughnut third sector
    With gCht.SeriesCollection(1).Points(3).Format.Glow
        .Color = ColorRange2 'msoThemeColorAccent4 .ObjectThemeColor
        .Color.TintAndShade = 1
        .Color.Brightness = 1
        .Transparency = 0.6
        .Radius = 8
    End With
    
    gCht.ChartArea.Format.Fill.visible = msoFalse
    gCht.ChartArea.Format.Line.visible = msoFalse
  
    
BasicGaugeStep_No_Error:
    On Error GoTo 0
    Exit Sub

BasicGaugeStep_Error:

    Call Error_Handle(BasicGaugeStep, Err.Number, Err.description, Erl)
    GoTo BasicGaugeStep_No_Error
End Sub

