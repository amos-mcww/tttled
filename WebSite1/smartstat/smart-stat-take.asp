<!--#include file="inc/plus.inc.asp"-->
<%
'**************************************************************
' Software name: SmartStat
' Web: http://www.cactussoft.cn
' Copyright (C) 2008��2010 ��������� ��Ȩ����
'**************************************************************
'�жϺ�̨��¼�Ƿ����
if GetIsAdmin()=False then
	Response.Write "<script language='javascript'>"
	Response.Write "parent.parent.parent.window.navigate('"&theURL&"login.asp');"
	Response.Write "</script>"
	Response.End
end if
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<link href="style/style.css" rel="stylesheet" type="text/css">
<title>��ȡ����</title>
</head>
<body>
<script type="text/javascript"> 
	function copyText(){ 
		var e=document.getElementById("take");//������content 
		e.select(); //ѡ����� 
		document.execCommand("copy"); //ִ��������������� 
    } 
</script>
<table width="99%" border="0" align="center" cellpadding="3" cellspacing="0">
  <tr>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td class="SmartTitle">��ȡ����</td>
  </tr>
  <tr>
    <td><textarea name="take" rows="10" id="take" style="wIDth:100%">
<%
Response.Write "<!-- ͳ�ƴ��뿪ʼ -->" & VbCrLf
Response.Write "<script language=""JavaScript"" src="""&theurl&"smartstat.asp?siteID=1"" type=""text/JavaScript""></script>"& VbCrLf
Response.Write "<!-- ͳ�ƴ������ -->" & VbCrLf
%></textarea></td>
  </tr>
  <tr>
    <td><input type=button value="���ƴ���(C)" onClick="copyText()" class="SmartButton"></td>
  </tr>
</table>
<center>
<span class=smallTxt><%=smart_SystemCopyright%></span><br>
</center>
</body>
</html>
