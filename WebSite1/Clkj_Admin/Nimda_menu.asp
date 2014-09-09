<!--#include file="Nimda_function.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<link href="backdoor.css" type="text/css" rel="stylesheet">
<title>导航管理</title>
</head>

<body>
<table width="100%" cellpadding="0" cellspacing="0">
  <tr>
    <td height="39" class="tr_bg">&nbsp;导航管理 <font color="#ff0000"><%=request.querystring("Lei")%></font></td>
  </tr>
</table>
<% If request.querystring("Edit")="M_E" Then 
	Set Rs=server.createobject("ADODB.Recordset")
	Sql="select * from clkj_menu where clkj_menuid ="&request("clkj_menuid")
	Rs.open Sql,conn,1,1
%>
<table width="100%" cellpadding="0" cellspacing="0">
<form name="meun" action="Nimda_function.asp?Class=meun_edit" method="post">
<tr>
<td height="30" colspan="2" align="left" valign="middle" bgcolor="#F5F5F5" style="text-indent:2em;color:#ff0000;">修改导航栏目</td>
</tr>
<tr>
<td width="15%" height="30" align="center" valign="middle">栏目名称</td>
<td width="85%" align="left" class="h_td"><input name="meun_name" type="text" id="meun_name" value="<%=rs("clkj_menutitle")%>" size="15" /> 
排序
  <input name="meun_px" type="text" id="meun_px" value="<%=rs("clkj_paixu")%>" size="5" />
  缺省状态为&quot;0&quot;，由小到大排列：0，1，2，3</td>
</tr>
<tr>
<td width="15%" height="30" align="center" valign="middle">关键字<br>(keywords)</td>
<td width="85%" align="left" class="h_td"><input name="meun_key" type="text" id="meun_key" value="<%=rs("clkj_menukey")%>" size="25" />
如:衣服,帽子,关键词之间用逗号分开&quot;,&quot;关键词2到3个,首页关键词在参数设置里进行设置</td>
</tr>
<tr>
<td width="15%" height="30" align="center" valign="middle">描 述<br>(description)</td>
<td width="85%" align="left" class="h_td"><textarea name="meun_ms" cols="40" rows="3" id="meun_ms"><%=rs("clkj_menudes")%></textarea>
字数控制在100之内,首页关描述在参数设置里进行设置</td>
</tr><tr>
<td width="15%" height="30" align="center" valign="middle">链接地址</td>
<td width="85%" align="left" class="h_td"><input name="meun_dress" type="text" id="meun_dress" value="<%=rs("clkj_menuurl")%>" size="30" />
  尽量不要修改</td>
</tr>

<tr>
<td width="15%" height="30" align="center" valign="middle">&nbsp;</td>
<td width="85%" align="left" class="h_td"><input type="submit" name="Submit" value="修改" /><input type="hidden" name="clkj_menuid" value="<%=request("clkj_menuid")%>" /></td>
</tr>
</form>
</table>
<%
Rs.close
Else
%>
<table width="100%" cellpadding="0" cellspacing="0">
<form name="meun" action="Nimda_function.asp?Class=meun" method="post">
<tr>
<td height="30" colspan="2" align="left" valign="middle" bgcolor="#F5F5F5" style="text-indent:2em;color:#ff0000;">增加导航栏目</td>
</tr>
<tr>
<td width="15%" height="30" align="center" valign="middle">栏目名称</td>
<td width="85%" align="left" class="h_td"><input name="meun_name" type="text" id="meun_name" size="15" /> 
排序
  <input name="meun_px" type="text" id="meun_px" value="0" size="5" />
  缺省状态为&quot;0&quot;，由小到大排列：0，1，2，3</td>
</tr>
<tr>
<td width="15%" height="30" align="center" valign="middle">关键字<br>(keywords)</td>
<td width="85%" align="left" class="h_td"><input name="meun_key" type="text" id="meun_key" size="25" />
如:衣服,帽子,关键词之间用逗号分开&quot;,&quot;关键词2到3个。首页关键词在参数设置里进行设置 </td>
</tr>
<tr>
<td width="15%" height="30" align="center" valign="middle">描 述<br>(description)</td>
<td width="85%" align="left" class="h_td"><textarea name="meun_ms" cols="40" rows="3" id="meun_ms"></textarea>
字数控制在100之内,首页关描述在参数设置里进行设置</td>
</tr>
<tr>
<td width="15%" height="30" align="center" valign="middle">链接地址</td>
<td width="85%" align="left" class="h_td"><input name="meun_dress" type="text" id="meun_dress" size="30" />
  尽量不要修改</td>
</tr>

<tr>
<td width="15%" height="30" align="center" valign="middle">&nbsp;</td>
<td width="85%" align="left" class="h_td"><input type="submit" name="Submit" value="增加" /></td>
</tr>
</form>
</table>
<%End If%>
<br>
<table width="100%" cellpadding="0" cellspacing="0" >
<tr>
<td height="39" colspan="5" class="tr_bg" style="color:#FF0000;">&nbsp;导航列表 &nbsp;[<a href="Nimda_menu.asp">增加栏目</a>]</td>
</tr>

<%Call Menu_menu()%> 

</table>
</body>
</html>
