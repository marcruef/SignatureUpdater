Attribute VB_Name = "modHttp"
Option Explicit

Private Declare Function InternetOpen Lib "wininet" Alias "InternetOpenA" (ByVal sAgent As String, ByVal lAccessType As Long, ByVal sProxyName As String, ByVal sProxyBypass As String, ByVal lFlags As Long) As Long
Private Declare Function InternetOpenUrl Lib "wininet" Alias "InternetOpenUrlA" (ByVal hInternetSession As Long, ByVal sUrl As String, ByVal sHeaders As String, ByVal lHeadersLength As Long, ByVal lFlags As Long, ByVal lContext As Long) As Long
Private Declare Function InternetReadFile Lib "wininet" (ByVal hFile As Long, ByVal sBuffer As String, ByVal lNumBytesToRead As Long, lNumberOfBytesRead As Long) As Integer
Private Declare Function InternetCloseHandle Lib "wininet" (ByVal hInet As Long) As Integer

Private Const INTERNET_OPEN_TYPE_PRECONFIG As Integer = 0
Private Const INTERNET_FLAG_RELOAD = &H80000000

Public Function HttpGetRequest(ByRef sUrl As String) As String
    Dim hInternetSession As Long
    Dim hURLFile As Long
    Dim sReadBuffer            As String * 4096
    Dim sBuffer                As String
    Dim lNumberOfBytesRead     As Long
    Dim bDoLoop As Boolean
    Dim lTotalBytes As Long
    
    hInternetSession = InternetOpen(APP_NAME, INTERNET_OPEN_TYPE_PRECONFIG, vbNullString, vbNullString, 0)
    
    If CBool(hInternetSession) Then
        hURLFile = InternetOpenUrl(hInternetSession, sUrl, vbNullString, 0, INTERNET_FLAG_RELOAD, 0)
        If CBool(hURLFile) Then
            bDoLoop = True
                While bDoLoop
                    sReadBuffer = ""
                    bDoLoop = InternetReadFile(hURLFile, sReadBuffer, Len(sReadBuffer), lNumberOfBytesRead)
                    sBuffer = sBuffer & Left$(sReadBuffer, lNumberOfBytesRead)
                    If Not CBool(lNumberOfBytesRead) Then bDoLoop = False
                    DoEvents
                    lTotalBytes = lTotalBytes + lNumberOfBytesRead
                Wend
                HttpGetRequest = sBuffer
        End If
    End If
    
    Call InternetCloseHandle(hURLFile)
    Call InternetCloseHandle(hInternetSession)
End Function

