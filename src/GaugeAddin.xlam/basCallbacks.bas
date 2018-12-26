Attribute VB_Name = "basCallbacks"

'################################################################
'#                                                              #
'#      Created with / Erstellt mit:                            #
'#      IDBE Ribbon Creator 2013                                #
'#      Version 1.1003                                          #
'#                                                              #
'#      (c) 2007-2013 IDBE Avenius                              #
'#                                                              #
'#      http://www.ribboncreator2013.com                        #
'#      http://www.ribboncreator2010.com                        #
'#      http://www.ribboncreator.com                            #
'#      http://www.accessribon.com                              #
'#      http://www.avenius.com                                  #
'#                                                              #
'#      You may send change requests or report errors to:       #
'#      Aenderungswuensche oder Fehler bitte an:                #
'#                                                              #
'#      mailto://info@ribboncreator2013.com                     #
'#                                                              #
'################################################################


' Globals

Public gobjRibbon As IRibbonUI

Public bolEnabled As Boolean    ' Used in Callback "getEnabled"
                                ' Further informations in Callback "getEnabled"
                                ' Für Callback "getEnabled"
                                ' Genauere Informationen in Callback "getEnabled".
                               
Public bolVisible As Boolean    ' Used in Callback "getVisible"
                                ' More information in Callback "getVisible
                                ' Für Callback "getVisible"
                                ' Further informations in Callback "getVisible

' For Sample Callback "GetContent"
' Fuer Beispiel Callback "GetContent"
Public Type ItemsVal
    id As String
    label As String
    imageMso As String
End Type


' Callbacks

Public Sub OnRibbonLoad(ribbon As IRibbonUI)
'Callbackname in XML File "onLoad"

    Set gobjRibbon = ribbon
End Sub

Public Sub OnActionButton(control As IRibbonControl)
'Callback in XML File "onAction"

    ' Callback for event button click
    ' Callback für Button Click
    
    Select Case control.id
        'Case "btnInfo"
        '    DoCmd.OpenForm "frmMyForm"
        Case "btnSimpleGauge"
            BuildGaugeChart
        Case "btnTesting"
            Testing
        Case "btnAbout"
            About
        Case Else
            MsgBox "Button """ & control.id & """ clicked" & vbCrLf & _
                           "Es wurde auf Button """ & control.id & """ in Ribbon geklickt", _
                           vbInformation
    End Select
End Sub

'Command Button

Sub OnActionButtonHelp(control As IRibbonControl, ByRef CancelDefault)
    ' Callbackname in XML File Command "onAction"

    ' Callback for command event button click
    ' Callback fuer Command Button Click

    MsgBox "Button ""Help"" clicked" & vbCrLf & _
                           "Es wurde auf Button ""Hilfe"" geklickt", _
                           vbInformation
    CancelDefault = True

End Sub

Sub OnActionCheckBox(control As IRibbonControl, _
                               pressed As Boolean)
    ' Callbackname in XML File "OnActionCheckBox"
    
    ' Callback for event checkbox click
    ' Callback für Checkbox Click

    Select Case control.id
        'Case "chkMyCheckbox"
        '    If pressed = True Then
        '
        '    Else
        '
        '    End If
        '
        Case Else
            MsgBox "The Value of the Checkbox """ & control.id & """ is: " & pressed & vbCrLf & _
                   "Der Wert der Checkbox """ & control.id & """ ist: " & pressed, _
                   vbInformation
    End Select

End Sub

Sub GetPressedCheckBox(control As IRibbonControl, _
                       ByRef bolReturn)
    
    ' Callbackname in XML File "GetPressedCheckBox"
    
    ' Callback for checkbox
    ' indicates how the control is displayed
    ' Callback für Checkbox wie das Control
    ' angezeigt werden soll

    Select Case control.id
        Case Else
            If getTheValue(control.Tag, "DefaultValue") = "1" Then
                bolReturn = True
            Else
                bolReturn = False
            End If
    End Select

End Sub


Sub OnActionTglButton(control As IRibbonControl, _
                       pressed As Boolean)
                              
    ' Callbackname in XML File "onAction"
    
    ' Callback für einen Toggle Button Klick
    ' Callback for a Toggle Buttons click event

    Select Case control.id
        '    If pressed = True Then
        '
        '    Else
        '
        '    End If
        Case Else
            MsgBox "The Value of the Toggle Button """ & control.id & """ is: " & pressed & vbCrLf & _
                   "Der Wert der Toggle Button """ & control.id & """ ist: " & pressed, _
                   vbInformation
    End Select

End Sub

Sub GetPressedTglButton(control As IRibbonControl, _
                       ByRef pressed)
' Callbackname in XML File "getPressed"

' Callback für ein Access ToogleButton Control wie dieser Angezeigt werden soll
' Callback for an Access ToogleButton Control. Indicates how the control is displayed

    Select Case control.id
        Case Else
            If getTheValue(control.Tag, "DefaultValue") = "1" Then
                pressed = True
            Else
                pressed = False
            End If
    End Select
End Sub

Public Sub GetEnabled(control As IRibbonControl, ByRef enabled)
    ' Callbackname in XML File "getEnabled"
    
    ' To set the property "enabled" to a Ribbon Control
    ' For further information see: http://www.accessribbon.de/en/index.php?Downloads:12
    ' Setzen der Enabled Eigenschaft eines Ribbon Controls
    ' Weitere Informationen: http://www.accessribbon.de/index.php?Downloads:12

    Select Case control.id
        'Case "ID_XMLRibbControl"
        '    enabled = bolEnabled
        Case Else
            enabled = True
    End Select
End Sub

Public Sub GetVisible(control As IRibbonControl, ByRef visible)
    ' Callbackname in XML File "getVisible"
    
    ' To set the property "visible" to a Ribbon Control
    ' For further information see: http://www.accessribbon.de/en/index.php?Downloads:12
    ' Setzen der Visible Eigenschaft eines Ribbon Controls
    ' Weitere Informationen: http://www.accessribbon.de/index.php?Downloads:12

    Select Case control.id
        'Case "ID_XMLRibbControl"
        '    visible = bolVisible
        Case Else
            visible = True
    End Select
End Sub

Sub GetLabel(control As IRibbonControl, ByRef label)
    ' Callbackname in XML File "getLabel"
    ' To set the property "label" to a Ribbon Control

    Select Case control.id
        ''GetLabel''
        Case Else
            label = "*getLabel*"

    End Select

End Sub

Sub GetScreentip(control As IRibbonControl, ByRef screentip)
    ' Callbackname in XML File "getScreentip"
    ' To set the property "screentip" to a Ribbon Control

    Select Case control.id
        ''GetScreentip''
        Case Else
            screentip = "*getScreentip*"

    End Select

End Sub

Sub GetSupertip(control As IRibbonControl, ByRef supertip)
    ' Callbackname in XML File "getSupertip"
    ' To set the property "supertip" to a Ribbon Control

    Select Case control.id
        ''GetSupertip''
        Case Else
            supertip = "*getSupertip*"

    End Select

End Sub

Sub GetDescription(control As IRibbonControl, ByRef description)
    ' Callbackname in XML File "getDescription"
    ' To set the property "description" to a Ribbon Control

    Select Case control.id
        ''GetDescription''
        Case Else
            description = "*getDescription*"

    End Select

End Sub

Sub GetTitle(control As IRibbonControl, ByRef title)
    ' Callbackname in XML File "getTitle"
    ' To set the property "title" to a Ribbon MenuSeparator Control

    Select Case control.id
        ''GetTitle''
        Case Else
            title = "*getTitle*"

    End Select

End Sub

'EditBox

Sub GetTextEditBox(control As IRibbonControl, _
                             ByRef strText)
    ' Callbackname in XML File "GetTextEditBox"
    
    ' Callback für EditBox welcher Wert in der
    ' EditBox eingetragen werden soll.
    ' Callback for an EditBox Control
    ' Indicates which value is to set to the control

    Select Case control.id
        Case Else
            strText = getTheValue(control.Tag, "DefaultValue")
    End Select
    
End Sub

Sub OnChangeEditBox(control As IRibbonControl, _
                              strText As String)
    ' Callbackname in XML File "OnChangeEditBox"
    
    ' Callback Editbox: Rückgabewert der Editbox
    ' Callback Editbox: Return value of the Editbox

    Select Case control.id
        'Case "MyEbx"
            'If strText = "Password" Then
            '
            'End If
        Case Else
            MsgBox "The Value of the EditBox """ & control.id & """ is: " & strText & vbCrLf & _
                   "Der Wert der EditBox """ & control.id & """ ist: " & strText, _
                   vbInformation
    End Select

End Sub

'DropDown

Sub OnActionDropDown(control As IRibbonControl, _
                             selectedId As String, _
                             selectedIndex As Integer)
    ' Callbackname in XML File "OnActionDropDown"
    
    ' Callback onAction (DropDown)
    
    Select Case control.id
        'Case "MyItemID"
        '   Select Case selectedId
        '...
        '   End Select
        Case Else
            Select Case selectedId
                Case Else
                    MsgBox "The selected ItemID of DropDown-Control """ & control.id & """ is : """ & selectedId & """" & vbCrLf & _
                           "Die selektierte ItemID des DropDown-Control """ & control.id & """ ist : """ & selectedId & """", _
                           vbInformation
            End Select
    End Select

End Sub

Sub GetSelectedItemIndexDropDown(control As IRibbonControl, _
                                 ByRef index)
    ' Callbackname in XML File "GetSelectedItemIndexDropDown"
    
    ' Callback getSelectedItemIndex (DropDown)
    
    Dim varIndex As Variant
    varIndex = getTheValue(control.Tag, "DefaultValue")
    
    If IsNumeric(varIndex) Then
        Select Case control.id
            Case Else
                index = varIndex
        End Select
    End If

End Sub

'Gallery

Sub OnActionGallery(control As IRibbonControl, _
                     selectedId As String, _
                     selectedIndex As Integer)
    ' Callbackname in XML File "OnActionGallery"
    
    ' Callback onAction (Gallery)
    
    Select Case control.id
        'Case "MyGalleryID"
        '   Select Case selectedId
        '      Case "MyGalleryItemID"
        '
        Case Else
            Select Case selectedId
                Case Else
                    MsgBox "The selected ItemID of Gallery-Control """ & control.id & """ is : """ & selectedId & """" & vbCrLf & _
                           "Die selektierte ItemID des Gallery-Control """ & control.id & """ ist : """ & selectedId & """", _
                           vbInformation
            End Select
    End Select

End Sub

Sub GetSelectedItemIndexGallery(control As IRibbonControl, _
                                   ByRef index)
    ' Callbackname in XML File "GetSelectedItemIndexGallery"
    
    ' Callback getSelectedItemIndex (Gallery)
    
    Dim varIndex As Variant
    varIndex = getTheValue(control.Tag, "DefaultValue")
    
    If IsNumeric(varIndex) Then
        Select Case control.id

            Case Else
                index = varIndex

        End Select

    End If

End Sub

'Combobox

Sub GetTextComboBox(control As IRibbonControl, _
                      ByRef strText)

    ' Callbackname im XML File "GetTextComboBox"
    
    ' Callback getText (Combobox)
                           
    Select Case control.id
        
        Case Else
            strText = getTheValue(control.Tag, "DefaultValue")
    End Select

End Sub


Sub OnChangeComboBox(control As IRibbonControl, _
                               strText As String)
                           
    ' Callbackname im XML File "OnChangeCombobox"
    
    ' Callback onChange (Combobox)
   
    Select Case control.id
        
        Case Else
            MsgBox "The selected Item of Combobox-Control """ & control.id & """ is : """ & strText & """" & vbCrLf & _
                   "Das selektierte Item des Combobox-Control """ & control.id & """ ist : """ & strText & """", _
                   vbInformation
    End Select

End Sub


' DynamicMenu

Sub GetContent(control As IRibbonControl, _
               ByRef XMLString)

    ' Sample for a Ribbon XML "getContent" Callback
    ' See also http://www.accessribbon.de/en/index.php?Access_-_Ribbons:Callbacks:dynamicMenu_-_getContent
    '     and: http://www.accessribbon.de/en/index.php?Access_-_Ribbons:Ribbon_XML___Controls:Dynamic_Menu

    ' Beispiel fuer einen Ribbon XML - "getContent" Callback
    ' Siehe auch: http://www.accessribbon.de/index.php?Access_-_Ribbons:Callbacks:dynamicMenu_-_getContent
    '       und : http://www.accessribbon.de/?Access_-_Ribbons:Ribbon_XML___Controls:Dynamic_Menu

    Select Case control.id

        Case Else
            XMLString = getXMLForDynamicMenu()
    End Select
 
End Sub


' Helper Function
' Hilfsfunktionen

Public Function getXMLForDynamicMenu() As String
    
    ' Creates a XML String for DynamicMenu CallBack - getContent
    
    ' Erstellt den Inhalt fuer das DynamicMenu im Callback getContent
    
    Dim lngDummy    As Long
    Dim strDummy    As String
    Dim strContent  As String
    
    Dim Items(4) As ItemsVal
    Items(0).id = "btnDy1"
    Items(0).label = "Item 1"
    Items(0).imageMso = "_1"
    Items(1).id = "btnDy2"
    Items(1).label = "Item 2"
    Items(1).imageMso = "_2"
    Items(2).id = "btnDy3"
    Items(2).label = "Item 3"
    Items(2).imageMso = "_3"
    Items(3).id = "btnDy4"
    Items(3).label = "Item 4"
    Items(3).imageMso = "_4"
    Items(4).id = "btnDy5"
    Items(4).label = "Item 5"
    Items(4).imageMso = "_5"
    
    strDummy = "<menu xmlns=""http://schemas.microsoft.com/office/2009/07/customui"">" & vbCrLf
    
        For lngDummy = LBound(Items) To UBound(Items)
            strContent = strContent & _
                "<button id=""" & Items(lngDummy).id & """" & _
                " label=""" & Items(lngDummy).label & """" & _
                " imageMso=""" & Items(lngDummy).imageMso & """" & _
                " onAction=""OnActionButton""/>" & vbCrLf
        Next
 

    strDummy = strDummy & strContent & "</menu>"
    getXMLForDynamicMenu = strDummy

End Function

Public Function getTheValue(strTag As String, strValue As String) As String
   ' *************************************************************
   ' Erstellt von     : Avenius
   ' Parameter        : Input String, SuchValue String
   ' Erstellungsdatum : 05.01.2008
   ' Bemerkungen      :
   ' Änderungen       :
   '
   ' Beispiel
   ' getTheValue("DefaultValue:=Test;Enabled:=0;Visible:=1", "DefaultValue")
   ' Return           : "Test"
   ' *************************************************************
      
   On Error Resume Next
      
   Dim workTb()     As String
   Dim Ele()        As String
   Dim myVariabs()  As String
   Dim i            As Integer

      workTb = Split(strTag, ";")
      
      ReDim myVariabs(LBound(workTb) To UBound(workTb), 0 To 1)
      For i = LBound(workTb) To UBound(workTb)
         Ele = Split(workTb(i), ":=")
         myVariabs(i, 0) = Ele(0)
         If UBound(Ele) = 1 Then
            myVariabs(i, 1) = Ele(1)
         End If
      Next
      
      For i = LBound(myVariabs) To UBound(myVariabs)
         If strValue = myVariabs(i, 0) Then
            getTheValue = myVariabs(i, 1)
         End If
      Next
      
End Function





