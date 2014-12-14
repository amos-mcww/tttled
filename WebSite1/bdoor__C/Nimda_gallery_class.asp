<!--#include file="Nimda_function.asp"-->
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<link href="backdoor.css" type="text/css" rel="stylesheet">
<title>Gallery分类管理</title>
</head>

<body>
<table width="100%" cellpadding="0" cellspacing="0">
  <tr>
    <td height="39" class="tr_bg">&nbsp;
      <p><strong>Gallery分类管理     </strong> <font color="red"><%=request.querystring("Lei")%></font></p>
    </td>
  </tr>
</table>

<!-------------------------------------------------------   一级分类管理 -------------------------------------------->
<%
' 修改一级分类
If request.querystring("Edit")="gallery_big_edit" Then
	Set Rs=server.createobject("ADODB.Recordset")
	Sql="select * from clkj_gallery_BigClass where clkj_BigClassID ="&request("clkj_BigClassID")
	Rs.open Sql,conn,1,1
%>

<table width="100%" border="0" cellpadding="0" cellspacing="0">
<form name="big_class_E" action="Nimda_function.asp?Class=gallery_big_edit" method="post">
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
  <td width="25%" height="30" align="center" valign="middle">关键词<br>(keywords)</td>
  <td width="75%" align="left" class="h_td"><p>
    <input name="big_key_name" type="text" id="big_key_name" value="<%=rs("clkj_BigClasskey")%>" size="70" />
    </p>
    <p>空格隔开</p>
    <p>&nbsp;</p></td>
</tr>
<tr>
<td width="25%" height="30" align="center" valign="middle">描述<br>(description)</td>
<td width="75%" align="left" class="h_td"><p>
  <textarea name="big_ms_name" cols="75" rows="12" id="big_ms_name"><%=rs("clkj_BigClassdes")%></textarea>
  </p>
  <p>建议控制在500个字以内</p>
  <p>&nbsp;</p></td>
</tr>
<tr>
<td width="25%" height="30" align="center" valign="middle">序号</td>
<td width="75%" align="left" class="h_td"><input name="big_paixu" type="text" id="big_paixu" value="<%=rs("big_paixu")%>" size="5">
  多个gallery项目中,数字越大该类越靠前</td>
</tr>
<tr>
<td width="25%" height="30" align="center" valign="middle">&nbsp;</td>
<td width="75%" align="left" class="h_td"><input type="submit" name="Submit" value="确认修改" />
<input type="hidden" name="clkj_BigClassID" value="<%=request("clkj_BigClassID")%>" /></td>
</tr>
</form>
</table>
<%
' 添加一级分类 '
Rs.close
Else If request.querystring("Edit")="gallery_big_add" then
%>
<table width="100%" border="0" cellpadding="0" cellspacing="0">
<form name="big_class" action="Nimda_function.asp?Class=gallery_big_add" method="post">
<tr>
<td height="30" colspan="2" align="left" valign="middle" bgcolor="#F5F5F5" style="text-indent:2em;color:#ff0000;" id="zy">增加一级分类</td>
</tr>
<tr>
<td width="25%" height="30" align="center" valign="middle">分类名称</td>
<td width="75%" align="left" class="h_td"><p>&nbsp;
  </p>
  <p>
    <input name="big_name" type="text" id="big_name" size="70" />
  </p>
  <p>必须唯一</p>
  <p>&nbsp;</p></td>
</tr>
<tr>
  <td width="25%" height="30" align="center" valign="middle">关键词<br>(keywords)</td>
  <td width="75%" align="left" class="h_td"><p>
    <input name="big_key_name" type="text" id="big_key_name" size="70" />
    </p>
    <p>空格隔开</p>
    <p>&nbsp;</p></td>
</tr>
<tr>
<td width="25%" height="30" align="center" valign="middle">描述<br>(description)</td>
<td width="75%" align="left" class="h_td"><p>
  <textarea name="big_ms_name" cols="75" rows="12" id="big_ms_name"></textarea>
  </p>
  <p>建议控制在500个字以内</p></td>
</tr>
<tr>
<td width="25%" height="30" align="center" valign="middle">序号</td>
<td width="75%" align="left" class="h_td"><input name="big_paixu" type="text" id="big_paixu" value="0" size="5">
  多个gallery项目中,数字越大该类越靠前</td>
</tr>
<tr>
<td width="25%" height="30" align="center" valign="middle">&nbsp;</td>
<td width="75%" align="left" class="h_td"><input type="submit" name="Submit" value="确认修改" />
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
<%
' 修改二级分类'
  if request.querystring("Edit")="gallery_small_edit" Then
	Set Rs=server.createobject("ADODB.Recordset")
	Sql="select * from clkj_gallery_SmallClass where clkj_SmallClassID ="&request("clkj_SmallClassID")
	Rs.open Sql,conn,1,1
%>
<table width="100%" cellpdding="0" cellspacing="0">
<form name="small_class_E" action="Nimda_function.asp?Class=gallery_small_edit" method="post">
<tr>
<td height="22" colspan="2" align="left" valign="middle" bgcolor="#F5F5F5" style="text-indent:2em;color:#ff0000;" id="xr">修改二级分类</td>
</tr>
<tr>
<td width="25%" height="30" align="center" valign="middle">一级分类名称</td>
<td align="left" valign="middle" class="h_td"><p>
  <select name="big_name" id="big_name">
    <%
  n=rs("clkj_SmallClassName")
  s=rs("clkj_SmallClassurl")
  k=rs("clkj_SmallClasskey")
  m=rs("clkj_smallClassdes")
  sp=rs("small_paixu")
  %>
    <option value="<%=rs("clkj_BigClassID")%>"><%=rs("clkj_BigClassName")%></option>
    <%Call gallery_Big_input()%>
  </select>
  请选择一级分类</p>
  <p>&nbsp;</p></td>

</tr>
<tr>
<td width="25%" height="30" align="center" valign="middle">二级分类名称</td>
<td><p>
  <input name="small_name" type="text" id="small_name" value="<%=n%>" size="70" />
</p>
  <p>    <span class="h_td">必须唯一</span></p>
  <p>&nbsp;</p></td>
</tr>
<tr>
  <td width="25%" height="30" align="center" valign="middle">关键词<br>(keywords)</td>
  <td><p>
    <input name="small_key_name" type="text" id="small_key_name" value="<%=k%>" size="70" />
    </p>
    <p><span class="h_td">空格隔开</span></p>
    <p>&nbsp;</p></td>
</tr>
<tr>
<td width="25%" height="30" align="center" valign="middle">描述<br>(description)</td>
<td><textarea name="small_ms_name" cols="75" rows="12" id="small_ms_name"><%=m%></textarea>
  <span class="h_td">建议控制在500个字以内</span></td>
</tr>
<tr>
<td width="25%" height="30" align="center" valign="middle">序号</td>
<td width="88%" align="left" class="h_td"><input name="small_paixu" type="text" id="small_paixu" value="<%=sp%>" size="5">
  多个gallery项目中,数字越大该类越靠前</td>
</tr>
<tr>
<td width="25%" height="30" align="center" valign="middle">&nbsp;</td>
<td width="88%" align="left" class="h_td"><input type="submit" name="Submit" value="修改" />
<input type="hidden" name="clkj_SmallClassID" value="<%=request("clkj_SmallClassID")%>" /></td>
</tr>
</form>
</table>

<%
' 添加二级分类'
Rs.close
Else  If request.querystring("Edit")="small_class_add" Then
%>
<table width="100%" cellpdding="0" cellspacing="0">
<form name="small_class" action="Nimda_function.asp?Class=gallery_small_add" method="post">
<tr>
<td height="30" colspan="2" align="left" valign="middle" bgcolor="#F5F5F5" style="text-indent:2em;color:#ff0000;" id="zr">添加二级分类</td>
</tr>
<tr>
<td width="25%" height="30" align="center" valign="middle">一级分类名称</td>
<td align="left" valign="middle" class="h_td"><p>
  <select name="big_name" id="big_name">
    <option value="">请选择</option>
    <%Call gallery_Big_input()%>
  </select>
  请选择一级分类</p>
  <p>&nbsp;</p></td>
</tr>
<tr>
<td width="25%" height="30" align="center" valign="middle">二级分类名称</td>
<td><p>
  <input name="small_name" type="text" id="small_name" size="70" />
  </p>
  <p><span class="h_td">必须唯一</span></p>
  <p>&nbsp;</p></td>
</tr>
<tr>

<td width="25%" height="30" align="center" valign="middle">关键词<br>(keywords)</td>
<td><p>
  <input name="small_key_name" type="text" id="small_key_name" size="70" />
  </p>
  <p><span class="h_td">空格隔开</span></p>
  <p>&nbsp;</p></td>
</tr>
<tr>

  <td width="25%" height="30" align="center" valign="middle">描述<br>(description)</td>
  <td><p>
    <textarea name="small_ms_name" cols="75" rows="12" id="small_ms_name"></textarea>
    </p>
    <p><span class="h_td">建议控制在500个字以内</span></p>
    <p>&nbsp;</p></td>
</tr>
<tr>
<td width="25%" height="30" align="center" valign="middle">序号</td>
<td width="88%" align="left" class="h_td"><input name="small_paixu" type="text" id="small_paixu" value="0" size="5">
  多个gallery项目中,数字越大该类越靠前</td>
</tr>
<tr>
<td width="25%" height="30" align="center" valign="middle">&nbsp;</td>
<td width="88%" align="left" class="h_td"><input type="submit" name="Submit" value=" 增加 " /></td>
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
