<!--#include file="Nimda_function.asp"-->
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<link href="backdoor.css" type="text/css" rel="stylesheet">
<title>分类管理</title>
</head>

<body>
<table width="100%" cellpadding="0" cellspacing="0">
  <tr>
    <td height="39" class="tr_bg">&nbsp;<strong>分类管理</strong> <font color="#0033ff"><%=request.querystring("Lei")%></font></td>
  </tr>
</table>

<!-------------------------------------------------------   一级分类管理 -------------------------------------------->
<% If request.querystring("Edit")="B_E" Then 
	Set Rs=server.createobject("ADODB.Recordset")
	Sql="select * from clkj_BigClass where clkj_BigClassID ="&request("clkj_BigClassID")
	Rs.open Sql,conn,1,1
%>

<table width="100%" border="1" cellpadding="0" cellspacing="0">
<form name="big_class_E" action="Nimda_function.asp?Class=big_Edit" method="post">
<tr>
<td height="30" colspan="2" align="left" valign="middle" bgcolor="#F5F5F5" style="text-indent:2em;color:#FF0000;" id="xy">修改一级分类</td>
</tr>
<tr>
<td width="25%" height="30" align="center" valign="middle">分类名称</td>
<td width="75%" align="left" class="h_td"><p>&nbsp;
  </p>
  <p>
    <input name="big_name" type="text" id="big_name" value="<%=rs("clkj_BigClassName")%>" size="70" />
  </p>
  <p>必须唯一</p>
  <p>&nbsp;</p></td>
</tr>
<tr>
<td width="25%" height="30" align="center" valign="middle">栏目文件链接名称</td>
<td width="75%" align="left" class="h_td"><p>&nbsp;
  </p>
  <p>
    <input name="big_wj_name" type="text" id="big_wj_name2" value="<%=rs("clkj_BigClassurl")%>" size="70" />
  </p>
  <p>用于生成静态作为文件名用。必须唯一</p>
  <p>&nbsp;</p></td>
</tr>
<tr>
<td width="25%" height="30" align="center" valign="middle">关键词<br>(keywords)</td>
<td width="75%" align="left" class="h_td"><p>&nbsp;
  </p>
  <p>
    <input name="big_key_name" type="text" id="big_key_name" value="<%=rs("clkj_BigClasskey")%>" size="70" />
  </p>
  <p>空格分隔</p>
  <p>&nbsp;</p></td>
</tr>
<tr>
<td width="25%" height="30" align="center" valign="middle">描述<br>(description)</td>
<td width="75%" align="left" class="h_td"><p>&nbsp;
  </p>
  <p>
    <textarea name="big_ms_name" cols="75" rows="12" id="big_ms_name"><%=rs("clkj_BigClassdes")%></textarea>
  </p>
  <p>建议控制在500个字以内</p></td>
</tr>
<tr>
<td width="25%" height="30" align="center" valign="middle">序号</td>
<td width="75%" align="left" class="h_td"><input name="big_paixu" type="text" id="big_paixu" value="<%=rs("big_paixu")%>" size="5">
  多个产品中,数字越大该类越靠前</td>
</tr>
<tr>
<td width="25%" height="30" align="center" valign="middle">&nbsp;</td>
<td width="75%" align="left" class="h_td"><input type="submit" name="Submit" value="确认修改" />
<input type="hidden" name="clkj_BigClassID" value="<%=request("clkj_BigClassID")%>" /></td>
</tr>
</form>
</table>
<%
Rs.close
Else If request.querystring("Edit")="B_Z" then
%>
<table width="100%" cellpadding="0" cellspacing="0">
<form name="big_class" action="Nimda_function.asp?Class=big" method="post">
<tr>
<td height="30" colspan="2" align="left" valign="middle" bgcolor="#F5F5F5" style="text-indent:2em;color:#ff0000;" id="zy">增加一级分类</td>
</tr>
<tr>
<td width="12%" height="30" align="center" valign="middle">分类名称</td>
<td width="88%" align="left" class="h_td"><input name="big_name" type="text" id="big_name" size="30" />
  必须唯一</td>
</tr>
<tr>
<td width="12%" height="30" align="center" valign="middle">栏目文件链接名称</td>
<td width="88%" align="left" class="h_td"><input name="big_wj_name" type="text" id="big_wj_name" size="30" />
  用于生成静态作为文件名用。必须唯一</td>
</tr>
<tr>
<td width="12%" height="30" align="center" valign="middle">关键词<br>(keywords)</td>
<td width="88%" align="left" class="h_td"><input name="big_key_name" type="text" id="big_key_name" size="30" />
  如：衣服，帽子，2到3个</td>
</tr>
<tr>
<td width="12%" height="30" align="center" valign="middle">描述<br>(description)</td>
<td width="88%" align="left" class="h_td"><textarea name="big_ms_name" cols="40" rows="3" id="big_ms_name"></textarea>
  建议控制在100个字以内</td>
</tr>
<tr>
<td width="12%" height="30" align="center" valign="middle">排序</td>
<td width="88%" align="left" class="h_td"><input name="big_paixu" type="text" id="big_paixu" value="10000" size="5">
  从小到大排0,1,2,3,4.默认为10000</td>
</tr>
<tr>
<td width="12%" height="30" align="center" valign="middle">&nbsp;</td>
<td width="88%" align="left" class="h_td"><input type="submit" name="Submit" value="增加" /></td>
</tr>
</form>
</table>
<%
End If
End If
%>
<br>
<!-------------------------------------------------------   二级分类管理  -------------------------------------------->
<br>
<% If request.querystring("Edit")="S_E" Then 
	Set Rs=server.createobject("ADODB.Recordset")
	Sql="select * from clkj_SmallClass where clkj_SmallClassID ="&request("clkj_SmallClassID")
	Rs.open Sql,conn,1,1
%>
<table width="100%" cellpdding="0" cellspacing="0">
<form name="small_class_E" action="Nimda_function.asp?Class=small_Edit" method="post">
<tr>
<td height="30" colspan="2" align="left" valign="middle" bgcolor="#F5F5F5" style="text-indent:2em;color:#ff0000;" id="xr">修改二级分类</td>
</tr>
<tr>
<td width="12%" height="30" align="center" valign="middle">一级分类名称</td>
<td align="left" valign="middle" class="h_td"><select name="big_name" id="big_name">
  <%
  n=rs("clkj_SmallClassName")
  s=rs("clkj_SmallClassurl")
  k=rs("clkj_SmallClasskey")
  m=rs("clkj_smallClassdes")
  sp=rs("small_paixu")
  tj=rs("clkj_tj")
  %>
  <option value="<%=rs("clkj_BigClassID")%>"><%=rs("clkj_BigClassName")%></option><%Call Big_input()%></select>
请选择一级分类</td>

</tr>
<tr>
<td width="12%" height="30" align="center" valign="middle">二级分类名称</td>
<td><input name="small_name" type="text" id="small_name" value="<%=n%>" size="30" />
  <span class="h_td">必须唯一</span></td>
</tr>
<tr>
<td width="12%" height="30" align="center" valign="middle">栏目文件链接名称</td>
<td><input name="small_wj_name" type="text" id="small_wj_name" value="<%=s%>" size="30" />
  <span class="h_td"> 用于生成静态作为文件名用。必须唯一</span></td>
</tr>
<tr>
<td width="12%" height="30" align="center" valign="middle">关键词<br>(keywords)</td>
<td><input name="small_key_name" type="text" id="small_key_name" value="<%=k%>" size="30" />
  <span class="h_td">如：衣服，帽子，2到3个</span></td>
</tr>
<tr>
<td width="12%" height="30" align="center" valign="middle">描述<br>(description)</td>
<td><textarea name="small_ms_name" cols="40" rows="3" id="small_ms_name"><%=m%></textarea>
  <span class="h_td">建议控制在100个字以内</span></td>
</tr>
<tr>
<td width="12%" height="30" align="center" valign="middle">排序</td>
<td width="88%" align="left" class="h_td"><input name="small_paixu" type="text" id="small_paixu" value="<%=sp%>" size="5">
  从小到大排0,1,2,3,4.默认为0</td>
</tr>
<tr>
  <td height="30" align="center" valign="middle">搜索推荐</td>
  <td align="left" class="h_td"><input type="checkbox" name="clkj_tj" id="clkj_tj" value="yes" <%if tj=1 then response.write"checked"%>>
    *推荐在搜索边上的内容</td>
</tr>
<tr>
<td width="12%" height="30" align="center" valign="middle">&nbsp;</td>
<td width="88%" align="left" class="h_td"><input type="submit" name="Submit" value="修改" />
<input type="hidden" name="clkj_SmallClassID" value="<%=request("clkj_SmallClassID")%>" /></td>
</tr>
</form>
</table>

<%
Rs.close
Else  If request.querystring("Edit")="S_Z" Then 
%>
<table width="100%" cellpdding="0" cellspacing="0">
<form name="small_class" action="Nimda_function.asp?Class=small" method="post">
<tr>
<td height="30" colspan="2" align="left" valign="middle" bgcolor="#F5F5F5" style="text-indent:2em;color:#ff0000;" id="zr">增加二级分类</td>
</tr>
<tr>
<td width="12%" height="30" align="center" valign="middle">一级分类名称</td>
<td align="left" valign="middle" class="h_td"><select name="big_name" id="big_name">
 <option value="">请选择</option><%Call Big_input()%></select> 
请选择一级分类</td>
</tr>
<tr>
<td width="12%" height="30" align="center" valign="middle">二级分类名称</td>
<td><input name="small_name" type="text" id="small_name" size="30" />
  <span class="h_td">必须唯一</span></td>
</tr>
<tr>
<td width="12%" height="30" align="center" valign="middle">栏目文件链接名称</td>
<td><input name="small_wj_name" type="text" id="small_wj_name" size="30" />
  <span class="h_td">用于生成静态作为文件名用。必须唯一</span></td>
</tr>
<tr>
<td width="12%" height="30" align="center" valign="middle">关键词<br>(keywords)</td>
<td><input name="small_key_name" type="text" id="small_key_name" size="30" />
  <span class="h_td">如：衣服，帽子，2到3个</span></td>
</tr>
<tr>
<td width="12%" height="30" align="center" valign="middle">描述<br>(description)</td>
<td><textarea name="small_ms_name" cols="40" rows="3" id="small_ms_name"></textarea>
  <span class="h_td">建议控制在100个字以内</span></td>
</tr>
<tr>
<td width="12%" height="30" align="center" valign="middle">排序</td>
<td width="88%" align="left" class="h_td"><input name="small_paixu" type="text" id="small_paixu" value="10000" size="5">
  从小到大排0,1,2,3,4.默认为10000</td>
</tr>
<tr>
  <td height="30" align="center" valign="middle">搜索推荐</td>
  <td align="left" class="h_td"><input name="clkj_tj" type="checkbox" id="clkj_tj"  value="yes" >
    *推荐在搜索边上的内容</td>
</tr>
<tr>
<td width="12%" height="30" align="center" valign="middle">&nbsp;</td>
<td width="88%" align="left" class="h_td"><input type="submit" name="Submit" value="增加" /></td>
</tr>
</form>
</table>
<%
End If
End If
%>
<br>
</body>
</html>
