<!--#include file="../Clkj_Inc/inc.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>模板管理</title>
<link href="backdoor.css" type="text/css" rel="stylesheet">
</head>
<body>
<%
IF Request.QueryString("del")<>"" Then
dim FolderUrl
FolderUrl="../Clkj_Template/"&Request.QueryString("del")
Set fso = CreateObject("Scripting.FileSystemObject")
Set fso2 = fso.GetFolder(Server.MapPath(FolderUrl)) 
fso2.delete
response.Redirect "nimda_Templates.asp?lei=ok"
End IF

'取出模板标记
		set rs=server.createobject("adodb.recordset")
		exec="select * from clkj_config where clkj_config_id=1"
		rs.open exec,conn,1,1
		mbbj=rs("mb")
%>
<table width="100%" cellpadding="0" cellspacing="0">
  <tr>
    <td height="39" class="tr_bg"><strong>&nbsp;模板管理</strong> <font color="#FF0000"><%=request.querystring("Lei")%> &nbsp;</font></td>
  </tr>
</table>


<table width="100%" cellpadding="0" cellspacing="0">
<tr>
<td height="30" align="center" valign="middle" bgcolor="#F7F7F7"><strong>序号</strong></td>
<td align="left" valign="middle" bgcolor="#F7F7F7"><strong>模板文件名称</strong></td>
<td align="left" valign="middle" bgcolor="#F7F7F7"><strong>创建时间</strong></td>
<td align="left" valign="middle" bgcolor="#F7F7F7"><strong>操作</strong></td>
</tr>
<%
Dim fso, f,f1,s,sf,i
i=1
  Set fso = CreateObject( "Scripting.FileSystemObject") 
  Set f = fso.GetFolder(Server.MapPath("../Clkj_Template")) 
  Set sf = f.SubFolders
  For Each f1 in sf 
    s=f1.name 
	d=f1.DateCreated
%> 
<td height="30" align="center" valign="middle"><%=i%></td>
<td align="left" valign="middle"><%=s%></td>
<td align="left" valign="middle"><%=d%></td>
<td align="left" valign="middle"><a href="Nimda_Template.asp?MoBanurl=<%=s%>" onClick="return confirm('是否应用模板?');">应用模板</a> | <a href="?deddl=<%=s%>" onClick="return confirm('是否删除模板?');">删除模板</a> | 
<%
if mbbj=s then
response.write "<font color=red>当前模板</font> | <a href='../' target='_blank'>预览</a>"

End if
%></td>
</tr>

<%
i=i+1
next
%>
<tr>
<td height="30" colspan="4" align="right">
<a href="http://www.sem-cms.com/mobanzhanshi.html" target="_blank"><font color="#0000FF">模板下载</font></a>&nbsp;&nbsp;
</td>
</tr>
</table>


</body>
</html>
