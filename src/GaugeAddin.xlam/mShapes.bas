Attribute VB_Name = "mShapes"
' ------------------------------------------------------
' Name: mShapes
' Kind: Module
' Purpose: Procedures for adding shapes to Gauge Chart
' Author: Tom Nordal
' Date: 2018-08-18
' Version: 0.5
' ------------------------------------------------------
Option Explicit

Sub AddShapeCenter(cht As Chart, Optional FontSize As Long = 8, Optional FontColor As Long = 0)
' ----------------------------------------------------------------
' Procedure Name: AddShapeCenter
' Purpose: Add shape to the center of the chart, for current value
' Procedure Kind: Sub
' Procedure Access: Public
' Parameter cht (Chart): Chart to place the shape, gives shapes position
' Author: Tom Nordal
' Date: 2018-08-18
' Version: 1.0
' ----------------------------------------------------------------
    
    Dim chtTop As Long
    Dim chtLeft As Long
    Dim chtWidth As Long
    Dim chtHight As Long
    
    Dim gShpCenterTop As Long
    Dim gShpCenterLeft As Long
    Dim gShpCenterWidth As Long
    Dim gShpCenterHight As Long
        
    chtTop = cht.ChartArea.Top
    chtLeft = cht.ChartArea.Left
    chtHight = cht.ChartArea.Height
    chtWidth = cht.ChartArea.Width
    

    gShpCenterWidth = chtWidth * 0.15
    gShpCenterHight = chtHight * 0.1
    
    gShpCenterTop = chtTop + (chtHight / 2) - (gShpCenterHight / 2)
    
    gShpCenterLeft = chtLeft + (chtWidth / 2) - (gShpCenterWidth / 2)
    
    
    Set gShpCenter = ActiveSheet.Shapes.AddTextbox(msoTextOrientationHorizontal, gShpCenterLeft, gShpCenterTop, gShpCenterWidth, gShpCenterHight)

    gShpCenter.TextFrame2.VerticalAnchor = msoAnchorMiddle
    gShpCenter.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
    gShpCenter.Select
    Selection.Formula = gCenterValueRange.Address

    gShpCenter.TextFrame2.TextRange.Font.Name = "+mn-lt"
    gShpCenter.TextFrame2.TextRange.Font.Bold = msoTrue
    gShpCenter.TextFrame2.TextRange.Font.Size = FontSize
    
    gShpCenter.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = FontColor
    
    gShpCenter.Fill.visible = msoFalse
    gShpCenter.Line.visible = msoFalse
    
    


End Sub

Public Sub MoveShape(ShapeName As String, ws As Worksheet, r As Range)

    
    ActiveSheet.Shapes.Range(Array(ShapeName)).Select
    Selection.Cut
    
    ws.Select
    r.Select
    ws.Paste
    r.Activate
    
    
End Sub

Sub AddShapeHeading(cht As Chart, Optional FontSize As Long = 10, Optional FontColor As Long = 0)
' ----------------------------------------------------------------
' Procedure Name: AddShapeHeading
' Purpose: Adds heading to chart
' Procedure Kind: Sub
' Procedure Access: Public
' Parameter cht (Chart): Chart for the heading, gives shape position
' Author: Tom Nordal
' Date: 2018-08-18
' Version: 1.0
' ----------------------------------------------------------------
    
    Dim chtTop As Long
    Dim chtLeft As Long
    Dim chtWidth As Long
    Dim chtHight As Long
    
    Dim gShpHeadingTop As Long
    Dim gShpHeadingLeft As Long
    Dim gShpHeadingWidth As Long
    Dim gShpHeadingHight As Long
        
    chtTop = cht.ChartArea.Top
    chtLeft = cht.ChartArea.Left
    chtHight = cht.ChartArea.Height
    chtWidth = cht.ChartArea.Width
    

    gShpHeadingWidth = chtWidth * 0.3
    gShpHeadingHight = chtHight * 0.1
    
    gShpHeadingTop = chtTop + 1 '+ (chtHight / 2) - (gShpHeadingHight / 2)
    
    gShpHeadingLeft = chtLeft + (chtWidth / 2) - (gShpHeadingWidth / 2)
    
    
    Set gShpHeading = ActiveSheet.Shapes.AddTextbox(msoTextOrientationHorizontal, gShpHeadingLeft, gShpHeadingTop, gShpHeadingWidth, gShpHeadingHight)

    gShpHeading.TextFrame2.VerticalAnchor = msoAnchorMiddle
    gShpHeading.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
    gShpHeading.Select
    Selection.Formula = gHeadingRange.Address

    gShpHeading.TextFrame2.TextRange.Font.Name = "+mn-lt"
    gShpHeading.TextFrame2.TextRange.Font.Bold = msoTrue
    gShpHeading.TextFrame2.TextRange.Font.Size = FontSize
    
    gShpHeading.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = FontColor
    
    gShpHeading.Fill.visible = msoFalse
    gShpHeading.Line.visible = msoFalse
    
    


End Sub

Sub AddShapeSubHeading(cht As Chart, Optional FontSize As Long = 9, Optional FontColor As Long = 0)
' ----------------------------------------------------------------
' Procedure Name: AddShapeSubHeading
' Purpose: Add shap for Sub Heading
' Procedure Kind: Sub
' Procedure Access: Public
' Parameter cht (Chart): Chart to add shap, gives position
' Author: Tom Nordal
' Date: 2018-08-18
' Version: 1.0
' ----------------------------------------------------------------
    
    Dim chtTop As Long
    Dim chtLeft As Long
    Dim chtWidth As Long
    Dim chtHight As Long
    
    Dim gShpHeadingTop As Long
    Dim gShpHeadingLeft As Long
    Dim gShpHeadingWidth As Long
    Dim gShpHeadingHight As Long
        
    chtTop = cht.ChartArea.Top
    chtLeft = cht.ChartArea.Left
    chtHight = cht.ChartArea.Height
    chtWidth = cht.ChartArea.Width
    

    gShpHeadingWidth = chtWidth * 0.4
    gShpHeadingHight = chtHight * 0.1
    
    gShpHeadingTop = chtTop + (chtHight / 2) + (chtHight * 0.05)
    
    gShpHeadingLeft = chtLeft + (chtWidth / 2) - (gShpHeadingWidth / 2)
    
    
    Set gShpSubHeading = ActiveSheet.Shapes.AddTextbox(msoTextOrientationHorizontal, gShpHeadingLeft, gShpHeadingTop, gShpHeadingWidth, gShpHeadingHight)

    gShpSubHeading.TextFrame2.VerticalAnchor = msoAnchorMiddle
    gShpSubHeading.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
    gShpSubHeading.Select
    Selection.Formula = gSubHeadingRange.Address

    gShpSubHeading.TextFrame2.TextRange.Font.Name = "+mn-lt"
    gShpSubHeading.TextFrame2.TextRange.Font.Bold = msoTrue
    gShpSubHeading.TextFrame2.TextRange.Font.Size = FontSize
    
    gShpSubHeading.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = FontColor
    
    gShpSubHeading.Fill.visible = msoFalse
    gShpSubHeading.Line.visible = msoFalse
    
    


End Sub


Sub AddShapeRight(cht As Chart, Optional FontSize As Long = 8, Optional FontColor As Long = 0)
' ----------------------------------------------------------------
' Procedure Name: AddShapeRight
' Purpose: Add shap at max of the chart
' Procedure Kind: Sub
' Procedure Access: Public
' Parameter cht (Chart): Chart to add the shape, gives position
' Author: Tom Nordal
' Date: 2018-08-18
' Version: 1.0
' ----------------------------------------------------------------
    
    Dim chtTop As Long
    Dim chtLeft As Long
    Dim chtWidth As Long
    Dim chtHight As Long
    
    Dim gShpHeadingTop As Long
    Dim gShpHeadingLeft As Long
    Dim gShpHeadingWidth As Long
    Dim gShpHeadingHight As Long
        
    chtTop = cht.ChartArea.Top
    chtLeft = cht.ChartArea.Left
    chtHight = cht.ChartArea.Height
    chtWidth = cht.ChartArea.Width
    

    gShpHeadingWidth = chtWidth * 0.15
    gShpHeadingHight = chtHight * 0.1
    
    gShpHeadingTop = chtTop + (chtHight / 2) - (gShpHeadingHight / 2)
    
    gShpHeadingLeft = chtLeft + (chtWidth / 2) + chtWidth * 0.15
    
    
    Set gShpRight = ActiveSheet.Shapes.AddTextbox(msoTextOrientationHorizontal, gShpHeadingLeft, gShpHeadingTop, gShpHeadingWidth, gShpHeadingHight)

    gShpRight.TextFrame2.VerticalAnchor = msoAnchorMiddle
    gShpRight.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignLeft
    gShpRight.Select
    Selection.Formula = gRightValueRange.Address

    gShpRight.TextFrame2.TextRange.Font.Name = "+mn-lt"
    gShpRight.TextFrame2.TextRange.Font.Bold = msoTrue
    gShpRight.TextFrame2.TextRange.Font.Size = FontSize
    
    gShpRight.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = FontColor
    
    gShpRight.Fill.visible = msoFalse
    gShpRight.Line.visible = msoFalse
    
    


End Sub
Sub AddShapRoundedRectangle(cht As Chart)
' ----------------------------------------------------------------
' Procedure Name: AddShapRoundedRectangle
' Purpose: Add background for the chart
' Procedure Kind: Sub
' Procedure Access: Public
' Parameter cht (Chart): Chart to add background, gives position
' Author: Tom Nordal
' Date: 2018-08-18
' Version: 1.0
' ----------------------------------------------------------------
    
    Dim chtTop As Long
    Dim chtLeft As Long
    Dim chtWidth As Long
    Dim chtHight As Long

    Dim shpTop As Long
    Dim shpLeft As Long
    Dim shpWidth As Long
    Dim shpHight As Long

    chtTop = cht.ChartArea.Top
    chtLeft = cht.ChartArea.Left
    chtHight = cht.ChartArea.Height
    chtWidth = cht.ChartArea.Width
    
    shpTop = chtTop - (chtHight * 0.1)
    shpWidth = chtWidth * 0.7
    shpLeft = chtLeft + (chtWidth / 2) - (shpWidth / 2)
    shpHight = (chtHight * 0.7) + (chtHight * 0.15)
    
    
    Set gShpBackground = ActiveSheet.Shapes.AddShape(msoShapeRoundedRectangle, shpLeft, shpTop, shpWidth, shpHight)
    
    gShpBackground.ZOrder msoSendToBack
    gShpBackground.SoftEdge.Radius = 15
    gShpBackground.Fill.ForeColor.RGB = RGB(192, 192, 192)

End Sub




