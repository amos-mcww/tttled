<!--#include file="inc/plus.inc.asp"-->
<%
'**************************************************************
' Software name: SmartStat
' Web: http://www.cactussoft.cn
' Copyright (C) 2008��2010 ���������� ��Ȩ����
'**************************************************************
'�жϺ�̨��¼�Ƿ����
if GetIsAdmin()=False then
	Response.Write "<script language='javascript'>"
	Response.Write "parent.parent.parent.window.navigate('"&theURL&"login.asp');"
	Response.Write "</script>"
	Response.End
end if
StartNow=Trim(Request("StartNow"))
EndNow=Trim(Request("EndNow"))
strTypeTo=Trim(Request("TypeTo"))

If StartNow<>"" and EndNow<>"" then
	strReportTitle="��" & StartNow & "��" & EndNow
Else
	strReportTitle="ȫ��"
End if

if strTypeTo=0 then
	strReportName="��·"
elseif strTypeTo=1 then
	strReportName="�����ַ"
elseif strTypeTo=2 then
	strReportName="�ؼ���"
elseif strTypeTo=4 then
	strReportName="ҳ�����"
end if

'��ȡ����
strWhere=" where "
if StartNow<>"" then strWhere=strWhere & "and (LogPageLastTime+"&Site_MasterTimeZone&"/24 >= #" & StartNow & "#) "
if EndNow<>"" then strWhere=strWhere & "and (LogPageLastTime+"&Site_MasterTimeZone&"/24 < #" & EndNow & "#) "
strWhere=strWhere & "and site_id=" & SiteID & " "
strWhere=replace(strWhere,"where and","where")
if trim(strWhere)="where" then strWhere=""

totalClient=conn.execute("select sum(LogPageTotal) from tblOriginPage " & strWhere & " and LogPageType="& strTypeTo & "")(0)
'response.write totalClient
'response.End
' �õ�����ߵĵ�ǰʱ��
uNow = dateadd("h",Site_MasterTimeZone-smart_ZoneServer,now())
thisUrl="?TypeTo="&strTypeTo&"&StartNow="&StartNow&"&EndNow="&EndNow
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<link href="style/style.css" rel="stylesheet" type="text/css">
<title><%=strReportName%></title>
</head>
<body>
<table width="99%" border="0" align="center" cellpadding="3" cellspacing="0">
  <tr>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td class="SmartTitle"><%=strReportName%></td>
  </tr>
  <tr>
    <td><table border="0" align="right" cellpadding="5" cellspacing="0">
        <tr>
          <td><a href="?TypeTo=<%=strTypeTo%>&StartNow=<%=GetNow(0)%>&EndNow=<%=GetNow(2)%>">����</a></td>
          <td><a href="?TypeTo=<%=strTypeTo%>&StartNow=<%=GetNow(1)%>&EndNow=<%=GetNow(0)%>">����</a></td>
          <td><a href="?TypeTo=<%=strTypeTo%>&StartNow=<%=GetNow(7)%>&EndNow=<%=GetNow(0)%>">���7��</a></td>
          <td><a href="?TypeTo=<%=strTypeTo%>&StartNow=<%=GetNow(30)%>&EndNow=<%=GetNow(0)%>">���30��</a></td>
          <td><a href="?TypeTo=<%=strTypeTo%>">ȫ��</a></td>
        </tr>
    </table></td>
  </tr>
  <tr>
    <td><strong><%=strReportName%></strong>��<%=strReportTitle%>��</td>
  </tr>
  <tr>
    <td><table width="100%" border="0" cellpadding="2" cellspacing="1" bgcolor="cedced">
      <tr>
        <td bgcolor="e9f1fa">&nbsp;</td>
        <td width="15%" bgcolor="e9f1fa"><div align="center"><strong>������</strong></div></td>
        <td width="15%" bgcolor="e9f1fa"><div align="center"><strong>����</strong></div></td>
        <td width="15%" bgcolor="e9f1fa"><div align="center"><strong>����������</strong></div></td>
      </tr>
<%
set rs=server.createobject("adodb.recordset")
SQL = "select LogPageContent,LogPageLastURL,LogPageTotal,LogPageLastTime from tblOriginPage "&strWhere&" and LogPageType="&strTypeTo&" order by LogPageTotal desc"
'response.write SQL
'response.End
rs.open sql,conn,1,1
If rs.bof or rs.eof Then 
%>
      <tr>
        <td colspan="4" bgcolor="#FFFFFF">��������</td>
        </tr>
<%
else
	dim a
	Set a=New PageList
	listLimit=20
	'ģ���ҳһҳ��ʾ��ҳ��
	pageLists=5
	'���÷�ҳ����
	
	'��ҳѭ����ʼ
	totalList=rs.recordcount
	a.InitPara=array(totalList,listLimit,pageLists)
	'a.PageList()
	a.PageListUrl(thisUrl)
	rs.pagesize=listLimit
	page = request("page")
	If page="" Then
		page=1
	End If
	rs.absolutepage = page
	for i=1 to rs.pagesize
%>
<tr>
<td bgcolor="#FFFFFF">
<%
select case strTypeTo
	case 0
		if rs(0)="" then response.write "ֱ�ӷ���" else response.write rs(0) & " [<a href="&rs(1)&" target=_blank>����</a>]" end if
	case 1
		response.write rs(0) & " [<a href="&rs(0)&" target=_blank>����</a>]"
	case 2
		response.write rs(0)
	case 4
		response.write rs(0) & " [<a href="&rs(0)&" target=_blank>����</a>]"
end select
%></td>
        <td bgcolor="#FFFFFF"><div align="center"><%=rs(2)%></div></td>
        <td bgcolor="#FFFFFF"><div align="center"><%=GetNumberRound(rs(2),totalClient)%></div></td>
        <td bgcolor="#FFFFFF"><div align="center"><%=rs(3)%></div></td>
      </tr>
<%
rs.movenext
if rs.eof then exit for
next
%>
    </table></td>
  </tr>
  <tr>
    <td><%=a.PageInfo%></td>
  </tr>
<%
rs.close
end if
%>
</table>
<center>
<span class=smallTxt><%=smart_SystemCopyright%></span><br>
</center>
</body>
</html>