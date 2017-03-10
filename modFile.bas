Attribute VB_Name = "modFile"
Option Explicit

Public Sub WriteFile(ByRef sFileName As String, ByRef sContent As String)
    Open sFileName For Output As #1
        Print #1, sContent
    Close
End Sub



