Attribute VB_Name = "modParse"
Option Explicit

Public Function RssLastEntry(ByRef sRssFeed As String) As String
    Dim sContent As String
    
    sContent = CreateStructure( _
        config_structure, _
        config_intro, _
        RssExtractItem(RssExtractItem(sRssFeed, "item"), "title"), _
        RssExtractItem(RssExtractItem(sRssFeed, "item"), "link"))
    
    Call WriteFile(config_sigfile, sContent)
    
    RssLastEntry = sContent
End Function

Public Function RssExtractItem(ByRef sRssData As String, ByRef sTagName As String) As String
    Dim iItemStart As Integer
    Dim iItemEnd As Integer
    Dim sTagStart As String
    Dim sTagEnd As String
    
    sTagStart = "<" & sTagName & ">"
    sTagEnd = "</" & sTagName & ">"
    
    iItemStart = InStr(1, sRssData, sTagStart, vbBinaryCompare)
    
    If (iItemStart) Then
        iItemEnd = InStr(iItemStart + Len(sTagStart), sRssData, sTagEnd, vbBinaryCompare)
    
        If (iItemEnd > iItemStart) Then
            RssExtractItem = Trim$(Mid$(sRssData, iItemStart + Len(sTagStart), iItemEnd - iItemStart - Len(sTagEnd) + 1))
        End If
    End If
End Function

Public Function CreateStructure(ByVal sStructure As String, ByRef sIntro As String, ByRef sTitle As String, ByRef sLink As String) As String
    sStructure = Replace(sStructure, "$intro", sIntro, 1, , vbBinaryCompare)
    sStructure = Replace(sStructure, "$title", sTitle, 1, , vbBinaryCompare)
    sStructure = Replace(sStructure, "$link", sLink, 1, , vbBinaryCompare)
    
    sStructure = Replace(sStructure, "$time", Time, 1, , vbBinaryCompare)
    sStructure = Replace(sStructure, "$date", Date, 1, , vbBinaryCompare)
    
    sStructure = Replace(sStructure, "\n", vbNewLine, 1, , vbBinaryCompare)
    sStructure = Replace(sStructure, "\t", vbTab, 1, , vbBinaryCompare)

    CreateStructure = sStructure
End Function
