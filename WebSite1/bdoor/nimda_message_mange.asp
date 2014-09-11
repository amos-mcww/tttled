<!--#include file="../Clkj_Inc/inc.asp"-->
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<link href="backdoor.css" type="text/css" rel="stylesheet">
<title>留言管理</title>
</head>
<script type="text/JavaScript">
<!--
function MM_openBrWindow(theURL,winName,features) 
{ //v2.0
  window.open(theURL,winName,features);
}
//-->
</script>
<body>
<table width="100%" cellpadding="0" cellspacing="0">
  <tr>
    <td height="39" class="tr_bg">&nbsp;<strong>询盘管理</strong> <font color="#ff0000"><%=request.querystring("Lei")%></font></td>
  </tr>
</table>
	<table width="100%" border="0" cellpadding="0" cellspacing="0">
      <!--DWLayoutTable-->
      <tr>
        <td width="57" height="25" align="center" valign="middle" bgcolor="#F5F5F5">编号</td>
        <td width="208" align="left" valign="middle" bgcolor="#F5F5F5">标题</td>
        <td width="318" align="left" valign="middle" bgcolor="#F5F5F5">所属地区</td>
        <td width="190" align="center" valign="middle" bgcolor="#F5F5F5">联系人</td>
        <td width="170" align="center" valign="middle" bgcolor="#F5F5F5">电话</td>
        <td width="188" align="center" valign="middle" bgcolor="#F5F5F5">E-mail</td>
        <td width="120" align="center" valign="middle" bgcolor="#F5F5F5">操作</td>
      </tr>
    </table>
	</td>
  </tr>
  <tr>
    <td height="20" valign="top"><table width="100%" cellpadding="0" cellspacing="0" style="border-top:1px solid #A5B9C9;">
	
	<%
	sql="select * from clkj_message order by id desc"
	set rs=server.CreateObject("adodb.recordset")
	rs.open sql,conn,1,3
IF Not rs.eof Then
proCount=rs.recordcount
	rs.PageSize=17			  '定义每页显示数目
		if not IsEmpty(Request("ToPage")) then
	       ToPage=CInt(Request("ToPage"))
		   if ToPage>rs.PageCount then
					rs.AbsolutePage=rs.PageCount
					intCurPage=rs.PageCount
		   elseif ToPage<=0 then
					rs.AbsolutePage=1
					intCurPage=1
				else
					rs.AbsolutePage=ToPage
					intCurPage=ToPage
				end if
			else
			        rs.AbsolutePage=1
					intCurPage=1
		  end if
	intCurPage=CInt(intCurPage)
	For i = 1 to rs.PageSize
	if rs.EOF then     
	Exit For 
	end if
%> 

       <tr>
        <td width="58" height="25" align="center" valign="middle" style="border-bottom:1px solid #A5B9C9;border-right:1px solid #A5B9C9;"><%=rs("id")%></td>
        <td width="207" align="left" valign="middle" style="border-bottom:1px solid #A5B9C9;border-right:1px solid #A5B9C9; text-indent:0.5em;">&nbsp;<%=rs("title")%></td>
        <td width="317" align="left" valign="middle" style="border-bottom:1px solid #A5B9C9;border-right:1px solid #A5B9C9; text-indent:0.5em;">&nbsp;<%=rs("contry")%></td>
        <td width="192" align="center" valign="middle" style="border-bottom:1px solid #A5B9C9;border-right:1px solid #A5B9C9;" >&nbsp;<%=rs("name")%></td>
        <td width="172" align="center" valign="middle" style="border-bottom:1px solid #A5B9C9;border-right:1px solid #A5B9C9;">&nbsp;<%=rs("phone")%></td>
        <td width="185" align="center" valign="middle" style="border-bottom:1px solid #A5B9C9;border-right:1px solid #A5B9C9;">&nbsp;<%=rs("email")%></td>
        <td width="118" align="center" valign="middle" style="border-bottom:1px solid #A5B9C9;border-right:1px solid #A5B9C9;"><a href="javascript:" onClick="MM_openBrWindow('nimda_message_view.asp?id=<%=rs("id")%>','客户信息','width=640,height=480 toolbar=no,location=no,directories=no,status=no,menubar=no,resizable=no,scrollbars=yes,top=176,left=161')">查看</a> <a href=nimda_message_mange.asp?id=<%=rs("id")%>&del=del>删除</a></td>
      </tr>
		<%
 rs.MoveNext
 next
 end if
 %><tr><td height=25 colspan="8" align="right" valign="middle" bgcolor="#F5F5F5"><div> 共&nbsp;<font color="#ff0000"><%=proCount%></font>&nbsp;条, 当前第：<font color="#ff0000"><%=intCurPage%></font>/<font color="#ff0000"><%=rs.PageCount%></font>页
                <% if intCurPage<>1 then %>
                <a href="nimda_message_mange.asp?cla=<%=nb%>&amp;ToPage=1" class="redlink">首页</a>&nbsp;|&nbsp;<a href="nimda_message_mange.asp?cla=<%=nb%>&amp;ToPage=<%=intCurPage-1%>" class="redlink">上一页</a>&nbsp;|
                <%else%>
                首页&nbsp;|&nbsp;上一页&nbsp;|
                <% end if%>
                <%if intCurPage<>rs.PageCount then %>
                <a href="nimda_message_mange.asp?cla=<%=nb%>&amp;ToPage=<%=intCurPage+1%>" class="redlink">下一页</a>&nbsp;| <a href="nimda_message_mange.asp?cla=<%=nb%>&amp;ToPage=<%=rs.PageCount%>" class="redlink"> 末页</a>&nbsp;|&nbsp;
                <%else%>
                <font color="#999999">下一页&nbsp;|&nbsp;末页&nbsp;|&nbsp; </font>
                <% end if%></div></td></tr>
    </table>
	<%
	if request.querystring("del")="del" then
	Sql="delete * from clkj_message where id="&request("id")  
		set Rs=server.CreateObject("adodb.recordset")
		conn.execute Sql
		Response.Redirect "nimda_message_mange.asp?Lei=提示：删除成功"
  End IF	
	
	%>
	
</body>
</html>
