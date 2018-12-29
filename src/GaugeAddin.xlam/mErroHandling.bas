Attribute VB_Name = "mErroHandling"
' ------------------------------------------------------
' Name: mErroHandling
' Kind: Module
' Purpose: Global Error Handler, Loging to log-file
' Author: Copy from: https://bettersolutions.com/vba/error-handling/log-file.htm
' Date: 2018-12-29
' ------------------------------------------------------

Public g_objFSO As Scripting.FileSystemObject
Public g_scrText As Scripting.TextStream


Public Sub Error_Handle(ByVal sRoutineName As String, ByVal sErrorNo As String, ByVal sErrorDescription As String, ByVal sLineNumber As String)
Dim sMessage As String
Dim sLogFile As String

    sLogFile = ThisWorkbook.Path & "\ErrorLog_GaugeAddin.txt"

   sMessage = sErrorNo & " - " & sErrorDescription & "- Line: " & sLineNumber
   
   Call MsgBox(sMessage, vbCritical, sRoutineName & " - Error")
   Call LogFile_WriteError(sLogFile, sRoutineName, sErrorNo, sErrorDescription, sLineNumber)
   
End Sub

Public Function LogFile_WriteError(ByVal sLogFile As String, ByVal sRoutineName As String, ByVal sErrorNumber, ByVal sMessage As String, sLineNumber As String)
Dim sText As String
   On Error GoTo ErrorHandler
   
   If (g_objFSO Is Nothing) Then
      Set g_objFSO = New FileSystemObject
   End If
   
   If (g_scrText Is Nothing) Then
      If (g_objFSO.FileExists(sLogFile) = False) Then
         Set g_scrText = g_objFSO.OpenTextFile(sLogFile, IOMode.ForWriting, True)
         g_scrText.WriteLine "Date,RoutineName,ErrorNumber,ErrorMessage,LineNumber"
      Else
         Set g_scrText = g_objFSO.OpenTextFile(sLogFile, IOMode.ForAppending)
      End If
   End If
   
   sText = sText & Format(Date, "yyyy-MM-dd") & "-" & Format(Time(), "HH:mm")
   sText = sText & "," & sRoutineName
   sText = sText & "," & sErrorNumber
   sText = sText & "," & sMessage
   sText = sText & "," & sLineNumber
   g_scrText.WriteLine sText
   g_scrText.Close
   
   Set g_scrText = Nothing
   Exit Function
   
ErrorHandler:
   Set g_scrText = Nothing
   Call MsgBox("Unable to write to log file", vbCritical, "LogFile_WriteError")
End Function

