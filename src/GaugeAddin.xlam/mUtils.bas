Attribute VB_Name = "mUtils"
Option Explicit
Option Private Module


Private Sub AA()

Application.Dialogs(xlDialogEditColor).Show 2, 255, 0, 0


End Sub
Public Function GetColorCode(StartColorCode As Long) As Long
Dim iRGBRed As Integer
Dim iRGBGreen As Integer
Dim iRGBBlue As Integer

Dim lResultCode As Long

iRGBRed = StartColorCode Mod 256
iRGBGreen = (StartColorCode / 256) Mod 256
iRGBBlue = StartColorCode / 65536

If Application.Dialogs(xlDialogEditColor).Show(1, iRGBRed, iRGBGreen, iRGBBlue) = True Then
    lResultCode = ActiveWorkbook.Colors(1)
Else
    lResultCode = StartColorCode
End If

GetColorCode = lResultCode

End Function
Public Sub ColorDialog01()

'Create variables for the color codes
Dim FullColorCode As Long
Dim RGBRed As Integer
Dim RGBGreen As Integer
Dim RGBBlue As Integer

'Get the color code from the cell named "RGBColor"
FullColorCode = Range("rRGBColorTest").Interior.Color

'Get the RGB value for each color (possible values 0 - 255)
RGBRed = FullColorCode Mod 256
RGBGreen = (FullColorCode \ 256) Mod 256
RGBBlue = FullColorCode \ 65536

'Open the ColorPicker dialog box, applying the RGB color as the default
If Application.Dialogs(xlDialogEditColor).Show _
    (1, RGBRed, RGBGreen, RGBBlue) = True Then

    'Set the variable RGBColorCode equal to the value
    'selected the DialogBox
    FullColorCode = ActiveWorkbook.Colors(1)
    
    'Set the color of the cell named "RGBColor"
    Range("rRGBColorTest").Interior.Color = FullColorCode
    Range("rRGBColorTest").Value = FullColorCode

Else
   
    'Do nothing if the user selected cancel

End If

End Sub
Public Function WorksheetExist(wsName As String) As Boolean
    Dim ws As Worksheet
    Dim bResult As Boolean

    bResult = False
    For Each ws In ActiveWorkbook.Worksheets
        If ws.Name Like wsName Then
            bResult = True
            Exit For
        End If
    Next

    WorksheetExist = bResult
 
End Function


