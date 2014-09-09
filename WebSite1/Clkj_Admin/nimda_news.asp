<!--#include file="../Clkj_Inc/inc.asp"-->
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<link href="backdoor.css" type="text/css" rel="stylesheet">
<title>信息添加</title>
</head> 
		<script charset="utf-8" src="../Clkj_Edit/kindeditor.js"></script>
        <script charset="utf-8" src="../Clkj_Edit/lang/zh_CN.js"></script>
		<script>
			KindEditor.ready(function(K) {
				K.create('#content', {
					allowFileManager : true
				});
			});
		</script>
<body>
<table width="100%" cellpadding="0" cellspacing="0">
  <tr>
    <td height="39" class="tr_bg">&nbsp;<strong>信息管理</strong> <font color="#0033ff"><%=request.querystring("Lei")%></font></td>
  </tr>
</table>
<%
if request.QueryString("Edit")="N_E" Then 
	Set Rs=server.createobject("ADODB.Recordset")
	Sql="select * from clkj_News where clkj_newsid="&request("clkj_newsid")
	Rs.open Sql,conn,1,1
		%>
        <form name="news" action="nimda_function.asp?Class=content_Edit" method="post">
<table width="100%" cellpadding="0" cellspacing="0">
<tr><td height="25" colspan="2" bgcolor="#F5F5F5" style="text-indent:2em; color:#FF0000">信息修改</td>
</tr>

<tr><td width="15%" align="center" valign="middle" height="30">类别选择</td>
  <td><select name="Lei">
    <option value="news" <%if rs("clkj_new_lei")="news" then%>selected<%end if%>>公司新闻</option>
    <option value="About us" <%if rs("clkj_new_lei")="About us" then%>selected<%end if%>>关于我们</option>
	<option value="faq" <%if rs("clkj_new_lei")="faq" then%>selected<%end if%>>常见问题</option>
  </select>
    <font color="#FF0000">*</font>请选择新闻类别.FAQ请选择</td>
</tr>
<tr>
  <td align="center" valign="middle" height="30">新闻标题</td>
  <td><input name="c_title" type="text" id="c_title" value="<%=rs("clkj_news_Title")%>" size="50">
    <font color="#FF0000">*</font>不能为空</td>
</tr>
<tr>
  <td align="center" valign="middle" height="30">关键词</td>
  <td><input name="clkj_news_key" type="text" id="clkj_news_key" value="<%=rs("clkj_news_key")%>" size="30">
   </td>
</tr>
<!--<tr>
  <td align="center" valign="middle" height="45">新闻页面顶部信息</td>
  <td><textarea name="clkj_news_db" cols="50" rows="5" id="clkj_db"><%=rs("clkj_news_db")%></textarea></td>
</tr>-->
<tr>
  <td width="15%" align="center" valign="middle">内容</td>
<td>
<textarea id="content" name="content" cols="100" rows="8" style="width:95%;height:300px;visibility:hidden;"><%=Server.HtmlEncode(rs("clkj_news_content"))%></textarea></td>
  </tr>

  <tr>
  <td align="center" valign="middle" height="30">录入员</td>
  <td><input name="c_ru" type="text" id="c_ru" value="<%=session("username")%>" size="15" >

    <input name="time" type="hidden" id="time" value="<%=Date()%>" size="20" readonly="readonly" ></td>
</tr>
  <tr>
    <td align="center" valign="middle" height="30">&nbsp;</td>
    <td><input type="submit" name="Submit" value="修改"><input type="hidden" value=<%=request.querystring("clkj_newsid")%> name="clkj_newsid"></td>
  </tr>
  
</table>
</form>
<%
Else
%>
<form name="news" action="nimda_function.asp?Class=content" method="post">
<table width="100%" cellpadding="0" cellspacing="0">
<tr><td height="25" colspan="2" bgcolor="#F5F5F5" style="text-indent:2em; color:#FF0000">信息添加</td>
</tr>

<tr><td width="15%" align="center" valign="middle" height="30">类别选择</td><td><select name="Lei">
  <option value="">请选择</option>
  <option value="news">公司新闻</option>
  <option value="About us">关于我们</option>
  <option value="faq">常见问题</option>
  
</select>
  <font color="#FF0000">*</font>请选择新闻类别,FAQ请选择</td>
</tr>
<tr>
  <td align="center" valign="middle" height="30">新闻标题</td>
  <td><input name="c_title" type="text" id="c_title" size="50">
  <font color="#FF0000">*</font>不能为空</td>
</tr>
<tr>
  <td align="center" valign="middle" height="30">关键词</td>
  <td><input name="clkj_news_key" type="text" id="clkj_news_key" size="30">
   </td>
</tr>
<tr>
  <td width="15%" align="center" valign="middle">内容</td>
<td>
<textarea id="content" name="content" cols="100" rows="8" style="width:95%;height:300px;visibility:hidden;"></textarea></td>
  </tr>

  <tr>
  <td align="center" valign="middle" height="30">录入员</td>
  <td><input name="c_ru" type="text" id="c_ru" value="<%=session("username")%>" size="15" >

    <input name="time" type="hidden" id="time" value="<%=Date()%>" size="20" readonly="readonly"></td>
</tr>
  <tr>
    <td align="center" valign="middle" height="30">&nbsp;</td>
    <td><input type="submit" name="Submit" value="提交"></td>
  </tr>
</table>
 </form>
<%end if%>
</body>
</html>
