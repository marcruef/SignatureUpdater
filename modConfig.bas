Attribute VB_Name = "modConfig"
Option Explicit

Public Const APP_NAME As String = "Signature Updater 1.0"

Public Const DEFAULT_SIGFILE As String = "signature.txt"
Public Const DEFAULT_FEEDURL As String = "http://www.scip.ch/?rss.vuldb"
Public Const DEFAULT_INTRO As String = "Letzte News: "

Public app_configuration As String

Public config_feedurl As String
Public config_sigfile As String
Public config_intro As String
Public config_autoclose As Integer
Public config_structure As String

Public Sub LoadConfigFromFile(Optional ByRef sConfigurationFileName As String)
    Dim iFreeFile As Integer
    Dim sTempString As String
    
    If (LenB(sConfigurationFileName)) Then
        app_configuration = App.Path & "\" & sConfigurationFileName
    Else
        app_configuration = App.Path & "\settings.ini"
    End If

    If (Dir$(app_configuration, 16) <> "") Then
        iFreeFile = FreeFile
        Open app_configuration For Input As #iFreeFile
            Do While Not EOF(iFreeFile)
                Line Input #iFreeFile, sTempString
                
                If (Left$(sTempString, 1) <> "#") Then
                    sTempString = Replace(sTempString, vbCrLf, vbNullString, , 1, vbBinaryCompare)
                    
                    If (InStrB(1, sTempString, "=", vbBinaryCompare)) Then
                        If (Left$(sTempString, 8) = "feedurl=") Then
                            config_feedurl = Mid$(sTempString, 9, Len(sTempString))
                        ElseIf (Left$(sTempString, 6) = "intro=") Then
                            config_intro = Mid$(sTempString, 7, Len(sTempString))
                        ElseIf (Left$(sTempString, 8) = "sigfile=") Then
                            config_sigfile = Mid$(sTempString, 9, Len(sTempString))
                        ElseIf (Left$(sTempString, 10) = "autoclose=") Then
                            config_autoclose = Val(Mid$(sTempString, 11, Len(sTempString)))
                        ElseIf (Left$(sTempString, 10) = "structure=") Then
                            config_structure = Mid$(sTempString, 11, Len(sTempString))
                        End If
                    End If
                End If
            Loop
        Close
    End If

    If (LenB(config_feedurl) = 0) Then
        config_feedurl = DEFAULT_FEEDURL
    End If

    If (LenB(config_intro) = 0) Then
        config_intro = DEFAULT_INTRO
    End If

    If (LenB(config_sigfile) = 0) Then
        config_sigfile = App.Path & "\" & DEFAULT_SIGFILE
    End If
    
    If (LenB(config_autoclose) = 0) Then
        config_autoclose = 0
    End If

    If (LenB(config_structure) = 0) Then
        config_structure = "$intro$\n$title - $link"
    End If
End Sub

