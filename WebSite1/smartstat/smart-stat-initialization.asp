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

act=Trim(Request("act"))

if act="del" then
	call del()
else
	call smarthome()
end if

if FoundErr=True then
	call ErrMsg(Message)
end if

sub smarthome()
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<link href="style/style.css" rel="stylesheet" type="text/css">
<title>վ���ʼ��</title>
</head>
<body>
<table width="99%" border="0" align="center" cellpadding="3" cellspacing="0">
  <tr>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td class="SmartTitle">վ���ʼ��</td>
  </tr>
  <tr>
    <td><strong class="SmartRed">����</strong></td>
  </tr>
  <tr>
    <td>վ�����ݳ�ʼ�����ܽ������վ���ȫ��ͳ�����ݣ�����վ�ָ�����δ��ʼͳ��ʱ��״̬��( �����ʼ�������г��ִ������볬������Ա��ϵֹͣϵͳ��Ȼ���ٽ��г�ʼ����)</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td>[<a href="?siteid=1&amp;act=del">�������ȷ�ϳ�ʼ��,ע���վ���������ݽ����</a>]</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
  </tr>
</table>
<%
end sub

sub del
	conn.execute ("update tblSite set UserStartTime=#1900-1-1#,UserTodayDate=#1900-1-1# where site_id=" & siteid)
	conn.execute ("delete * from tblClient where site_id=" & siteid)
	conn.execute ("delete * from tblOriginPage where site_id=" & siteid)
	conn.execute ("delete * from tblOnline where site_id=" & siteid)
	conn.execute ("delete * from tblView where site_id=" & siteid)
	
	'�ɹ���ʾ	   
	Call performMsg("�����ɹ�","")
End sub
%>
<center>
<span class=smallTxt><%=smart_SystemCopyright%></span><br>
</center>
</body>
</html>
