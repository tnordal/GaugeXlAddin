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

Public gwsChartSettings As Worksheet


Public Sub Testing()
    CleanChartSettingsheet
End Sub

Sub CleanChartSettingsheet()
Dim shp As Shape

Set gwsChartSettings = Worksheets("ChartSetup")

gwsChartSettings.Activate

For Each shp In gwsChartSettings.Shapes
    shp.Delete
Next
End Sub

Sub BuildGaugeChart()
10        On Error GoTo BuildGaugeChart_Error

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
          
          Dim gwsActive As Worksheet
          Dim grActiveCell As Range
          
20        Set gwsActive = ActiveSheet
30        Set grActiveCell = ActiveCell

40        frm.Show

50        If Not frm.ReturnValue Then
60            Unload frm
70            Exit Sub
80        End If


90        CopyGaugeSetup

100       gChartRangeName.Formula = frm.txtChartName.Text
110       gHeadingRange.Formula = ReturnValueFromForm(frm.refHeading.Value)
120       gSubHeadingRange.Formula = ReturnValueFromForm(frm.refSubHeading.Value)
          
130       gChartRangeActualValue.Formula = ReturnValueFromForm(frm.refActualValue.Value)
140       gChartRangeMaxValue.Formula = ReturnValueFromForm(frm.refMaxValue.Value)
          
150       gChartEndRange1.Formula = ReturnValueFromForm(frm.refRange1Max.Value)
160       gChartEndRange2.Formula = ReturnValueFromForm(frm.refRange2Max.Value)

170       gwsActive.Activate
180       grActiveCell.Activate
              
190       BasicGaugeStep frm.lblRange1Color.BackColor, frm.lblRange2Color.BackColor, frm.lblRange3Color.BackColor
          
200       AddShapeHeading gCht, frm.cmbFonSizeHeading.Value, frm.lblFontColorHeading.BackColor

210       AddShapeSubHeading gCht, frm.cmbFontSizeSubHeading.Value, frm.lblFontColorSubHeding.BackColor

220       AddShapeCenter gCht, frm.cmbFontSizeActualValue.Value, frm.lblFontColorActualValue.BackColor

230       AddShapeRight gCht, frm.cmbFontSizeMaxValue.Value, frm.lblFonColorMaxValue.BackColor

240       AddShapRoundedRectangle gCht, frm.lblBackgroundColor.BackColor

          
250       Set GaugeGroup = ActiveSheet.Shapes.Range(Array( _
              gCht.Parent.Name, _
              gShpCenter.Name, _
              gShpHeading.Name, _
              gShpSubHeading.Name, _
              gShpRight.Name, gShpBackground.Name))

260       GaugeGroup.Group

270       GaugeGroup.Name = gChartRangeName & gChartRangeID
          
280       GaugeGroup.Top = grActiveCell.Top
290       GaugeGroup.Left = grActiveCell.Left
              
300       Unload frm
          
          

BuildGaugeChart_No_Error:
310       On Error GoTo 0
320       Exit Sub

BuildGaugeChart_Error:

'330       MsgBox "Error " & Err.Number & " (" & Err.description & ") in procedure BuildGaugeChart, line " & Erl & "."
340       Call Error_Handle("BuildGaugeChart", Err.Number, Err.description, Erl)
350       GoTo BuildGaugeChart_No_Error

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
    
    Dim i As Integer
    
    Set wb = ThisWorkbook

    
    If WorksheetExist("ChartSetup") Then
        Set gwsChartSettings = Worksheets("ChartSetup")
        gwsChartSettings.Activate
    Else
        Set gwsChartSettings = ActiveWorkbook.Worksheets.Add
        gwsChartSettings.Name = "ChartSetup"
        Set gActiveCell = Range("A1")
        wsGaugeSetting.Range("rGaugeSetupHeadingRange").Copy
        gActiveCell.PasteSpecial xlPasteAll
    End If
        
    wsGaugeSetting.Range("rGaugeSetupRange").Copy
    
    i = Application.WorksheetFunction.CountA(gwsChartSettings.Range("1:1"))
    Set gActiveCell = gwsChartSettings.Cells(1, i + 1)
    
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

