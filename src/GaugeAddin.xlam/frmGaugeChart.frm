VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmGaugeChart 
   Caption         =   "Gauge Chart Setup"
   ClientHeight    =   6900
   ClientLeft      =   96
   ClientTop       =   408
   ClientWidth     =   7704
   OleObjectBlob   =   "frmGaugeChart.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmGaugeChart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private m_Cancel As Boolean

Private m_bReturnValue As Boolean
Public Property Get ReturnValue() As Boolean

    ReturnValue = m_bReturnValue

End Property
Public Property Let ReturnValue(ByVal bNewValue As Boolean)

    m_bReturnValue = bNewValue

End Property

Private Sub btnCancel_Click()
    m_bReturnValue = False
    Me.Hide
End Sub

Private Sub btnCreateChart_Click()
    m_bReturnValue = True
    Me.Hide
End Sub

Private Sub lblBackgroundColor_Click()
    lblBackgroundColor.BackColor = GetColorCode(lblBackgroundColor.BackColor)
End Sub

Private Sub lblFonColorMaxValue_Click()
    lblFonColorMaxValue.BackColor = GetColorCode(lblFonColorMaxValue.BackColor)
End Sub

Private Sub lblFontColorActualValue_Click()
    lblFontColorActualValue.BackColor = GetColorCode(lblFontColorActualValue.BackColor)
End Sub

Private Sub lblFontColorHeading_Click()
    lblFontColorHeading.BackColor = GetColorCode(lblFontColorHeading.BackColor)
End Sub

Private Sub lblFontColorSubHeding_Click()
    lblFontColorSubHeding.BackColor = GetColorCode(lblFontColorSubHeding.BackColor)
End Sub

Private Sub lblRange1Color_Click()
    lblRange1Color.BackColor = GetColorCode(lblRange1Color.BackColor)
End Sub

Private Sub lblRange2Color_Click()
    lblRange2Color.BackColor = GetColorCode(lblRange2Color.BackColor)
End Sub

Private Sub lblRange3Color_Click()
    lblRange3Color.BackColor = GetColorCode(lblRange3Color.BackColor)
End Sub


Private Sub UserForm_Initialize()
    Dim i As Integer, f As Integer

    For i = 6 To 16
        cmbFonSizeHeading.AddItem i
        cmbFontSizeSubHeading.AddItem i
        cmbFontSizeMaxValue.AddItem i
        cmbFontSizeActualValue.AddItem i
    Next

        cmbFonSizeHeading.ListIndex = 4
        cmbFontSizeSubHeading.ListIndex = 3
        cmbFontSizeMaxValue.ListIndex = 2
        cmbFontSizeActualValue.ListIndex = 2
    
    m_Cancel = True
    
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    Cancel = m_Cancel
End Sub

