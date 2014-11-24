<%
Function RemoveHTML(strHTML)
Dim objRegExp, Match, Matches 
Set objRegExp = New Regexp
	strHTML=Replace(strHTML,vbCrLf,"")
	strHTML=Replace(strHTML,"&nbsp;"," ")
	objRegExp.IgnoreCase = True
	objRegExp.Global = True
	objRegExp.Pattern = "<.+?>"
	Set Matches = objRegExp.Execute(strHTML)
	For Each Match in Matches 
	strHtml=Replace(strHTML,Match.Value,"")
	Next
	RemoveHTML=strHTML
Set objRegExp = Nothing
End Function
Function U2UTF8(Byval a_iNum)
    Dim sResult,sUTF8
    Dim iTemp,iHexNum,i

    iHexNum = Trim(a_iNum)

    If iHexNum = "" Then
        Exit Function
    End If

    sResult = ""

    If (iHexNum < 128) Then
        sResult = sResult & iHexNum
    ElseIf (iHexNum < 2048) Then
        sResult = ChrB(&H80 + (iHexNum And &H3F))
        iHexNum = iHexNum \ &H40
        sResult = ChrB(&HC0 + (iHexNum And &H1F)) & sResult
    ElseIf (iHexNum < 65536) Then
        sResult = ChrB(&H80 + (iHexNum And &H3F))
        iHexNum = iHexNum \ &H40
        sResult = ChrB(&H80 + (iHexNum And &H3F)) & sResult
        iHexNum = iHexNum \ &H40
        sResult = ChrB(&HE0 + (iHexNum And &HF)) & sResult
    End If

    U2UTF8 = sResult
End Function
Dim http



Function GB2UTF(Byval a_sStr)
    Dim sGB,sResult,sTemp
    Dim iLen,iUnicode,iTemp,i

    sGB = Trim(a_sStr)
    iLen = Len(sGB)
    For i = 1 To iLen
         sTemp = Mid(sGB,i,1)
         iTemp = Asc(sTemp)

         If (iTemp>127 OR iTemp<0) Then
             iUnicode = AscW(sTemp)
             If iUnicode<0 Then
                 iUnicode = iUnicode + 65536
             End If
        Else
            iUnicode = iTemp
        End If

        sResult = sResult & U2UTF8(iUnicode)
    Next

    GB2UTF = sResult
End Function

%>
