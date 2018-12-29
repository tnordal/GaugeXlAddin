Attribute VB_Name = "mUtils"
Option Explicit
Option Private Module

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
Public Sub ColorDialog()

'Create variables for the color codes
Dim FullColorCode As Long
Dim RGBRed As Integer
Dim RGBGreen As Integer
Dim RGBBlue As Integer

'Get the color code from the cell named "RGBColor"
FullColorCode = ActiveCell.Interior.Color

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
    ActiveCell.Interior.Color = FullColorCode
    ActiveCell.Value = FullColorCode

Else
   
    'Do nothing if the user selected cancel

End If

End Sub
Public Function ReturnValueFromForm(formValue As Variant)
Dim vReturnValue As Variant
    
    If IsNumeric(formValue) Then
        vReturnValue = CDbl(formValue)
    ElseIf InStr(formValue, "!$") Then
        vReturnValue = "=" & formValue
    Else
        vReturnValue = formValue
    End If
        
    ReturnValueFromForm = vReturnValue
    
End Function

Private Sub testIsFormual()
Dim s As Variant

On Error GoTo testIsFormual_Error
s = 40
s = "Sheet2!$B$3"


Debug.Print ReturnValueFromForm(s)



    
testIsFormual_No_Error:
    On Error GoTo 0
    Exit Sub

testIsFormual_Error:

    Call Error_Handle(testIsFormual, Err.Number, Err.description, Erl)
    GoTo testIsFormual_No_Error
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


