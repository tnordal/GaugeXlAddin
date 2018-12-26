Attribute VB_Name = "mMain"
' ------------------------------------------------------
' Name: mMain
' Kind: Module
' Purpose: Procedures for making Gauge Chart
' Author: Tom Nordal
' Date: 2018-08-18
' Version: 0.5
' ------------------------------------------------------
Option Explicit

Public gCht As Chart
Public gShpCenter As Shape
Public gShpHeading As Shape
Public gShpSubHeading As Shape
Public gShpRight As Shape
Public gShpBackground As Shape

Public gChartRangeID As Range
Public gChartRangeName As Range
Public gChartRange As Range
Public gChartRangeDonut As Range
Public gChartRangePie As Range
Public gChartRangePieName As Range
Public gChartRangeMaxValue As Range
Public gChartRangeActualValue As Range
Public gChartEndRange1 As Range
Public gChartEndRange2 As Range
Public gHeadingRange As Range
Public gSubHeadingRange As Range
Public gCenterValueRange As Range
Public gRightValueRange As Range

Public gActiveCell As Range


Public Sub Testing()
    ColorDialog
End Sub


Sub BuildGaugeChart()
'    On Error GoTo BuildGaugeChart_Error

' ----------------------------------------------------------------
' Procedure Name: BuildGaugeChart
' Purpose: Main procedure for making Gauge Chart
' Procedure Kind: Sub
' Procedure Access: Public
' Author: Tom Nordal
' Date: 2018-08-18
' Version: 0.6
' ----------------------------------------------------------------
    Dim GaugeGroup As ShapeRange
    Dim frm As New frmGaugeChart
    
    Dim wsActive As Worksheet
    Dim rActiveCell As Range
    
    Set wsActive = ActiveSheet
    Set rActiveCell = ActiveCell

    frm.Show

    If Not frm.ReturnValue Then
        Unload frm
        Exit Sub
    End If


    CopyGaugeSetup

    gChartRangeName.Formula = frm.txtChartName.Text
    gHeadingRange.Formula = frm.txtHeading.Text
    gSubHeadingRange.Formula = frm.txtSubHeading.Text
    
    gChartRangeActualValue = frm.txtActualValue.Text
    gChartRangeMaxValue = frm.txtMaxValue.Text
    
    gChartEndRange1 = frm.txtRange1Max.Text
    gChartEndRange2 = frm.txtRange2Max.Text

               
        
    BasicGaugeStep frm.lblRange1Color.BackColor, frm.lblRange2Color.BackColor, frm.lblRange3Color.BackColor
    
    AddShapeHeading gCht, frm.cmbFonSizeHeading.Value, frm.lblFontColorHeading.BackColor

    AddShapeSubHeading gCht, frm.cmbFontSizeSubHeading.Value, frm.lblFontColorSubHeding.BackColor

    AddShapeCenter gCht, frm.cmbFontSizeActualValue.Value, frm.lblFontColorActualValue.BackColor

    AddShapeRight gCht, frm.cmbFontSizeMaxValue.Value, frm.lblFonColorMaxValue.BackColor

    AddShapRoundedRectangle gCht


    Set GaugeGroup = ActiveSheet.Shapes.Range(Array( _
        gCht.Parent.Name, _
        gShpCenter.Name, _
        gShpHeading.Name, _
        gShpSubHeading.Name, _
        gShpRight.Name, gShpBackground.Name))

    GaugeGroup.Group

    GaugeGroup.Name = gChartRangeName & gChartRangeID
    
    MoveShape gChartRangeName & gChartRangeID, wsActive, rActiveCell
        
    Unload frm
    


BuildGaugeChart_No_Error:
    On Error GoTo 0
    Exit Sub

BuildGaugeChart_Error:

    MsgBox "Error " & Err.Number & " (" & Err.description & ") in procedure BuildGaugeChart, line " & Erl & "."
    GoTo BuildGaugeChart_No_Error

End Sub

Public Sub CopyGaugeSetup()
' ----------------------------------------------------------------
' Procedure Name: CopyGaugeSetup
' Purpose: Copy values for setting up Gauge chart
' Procedure Kind: Sub
' Procedure Access: Public
' Author: Tom Nordal
' Date: 2018-12-26
' Version: 1.1
' ----------------------------------------------------------------
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim i As Integer
    
    Set wb = ThisWorkbook

    
    If WorksheetExist("ChartSetup") Then
        Set ws = Worksheets("ChartSetup")
        ws.Activate
    Else
        Set ws = ActiveWorkbook.Worksheets.Add
        ws.Name = "ChartSetup"
        Set gActiveCell = Range("A1")
        wsGaugeSetting.Range("rGaugeSetupHeadingRange").Copy
        gActiveCell.PasteSpecial xlPasteAll
    End If
        
    wsGaugeSetting.Range("rGaugeSetupRange").Copy
    
    i = Application.WorksheetFunction.CountA(ws.Range("1:1"))
    Set gActiveCell = ws.Cells(1, i + 1)
    
    gActiveCell.PasteSpecial xlPasteAll
    
    Application.CutCopyMode = False
    
    Set gChartRangeID = gActiveCell
    Set gChartRangeName = gActiveCell.Offset(1)
    Set gChartRangeDonut = gActiveCell.Offset(11).Resize(5, 1)
    Set gChartRangePie = gActiveCell.Offset(17).Resize(3, 1)
    Set gChartRangePieName = gActiveCell.Offset(16)
    Set gChartRangeMaxValue = gActiveCell.Offset(7)
    Set gChartRangeActualValue = gActiveCell.Offset(8)
    Set gChartEndRange1 = gActiveCell.Offset(9)
    Set gChartEndRange2 = gActiveCell.Offset(10)
    Set gHeadingRange = gActiveCell.Offset(3)
    Set gSubHeadingRange = gActiveCell.Offset(4)
    Set gRightValueRange = gActiveCell.Offset(5)
    Set gCenterValueRange = gActiveCell.Offset(6)

    gChartRangeID = i
    
End Sub

Public Sub About()
Dim AppVersion As String

AppVersion = wsConfig.Range("rAppVersion")

    MsgBox "Application: DashAddin" & vbCrLf _
        & "Version      : " & AppVersion & vbCrLf _
        & "Date           : 2018" & vbCrLf _
        & "Copyright Tom Nordal (c) 2018", vbOKOnly Or vbInformation, "DashAddin"

End Sub

