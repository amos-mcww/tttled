<!--#include file="../Clkj_Inc/inc.asp"-->
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<link href="backdoor.css" type="text/css" rel="stylesheet">
<title>信息管理</title>
</head>

<body>

<table width="100%" cellpadding="0" cellspacing="0">
  <tr>
    <td height="39" class="tr_bg">&nbsp;<strong>信息管理</strong> <font color="#0033ff"><%=request.querystring("Lei")%></font> [<a href="nimda_news.asp">增加信息</a>]</td>
  </tr>
</table>
<table cellpadding="0" border="0" width="100%">
<tr>
	<td height="30" bgcolor="#F7FCFF"><strong>信息分类筛选：</strong> <a href="?newsclass=news">新闻信息</a> | <a href="?newsclass=About us">关于我们</a> | <a href="?newsclass=faq">常见问题</a></td>
</tr>
</table>
	<table width="100%" border="0" cellpadding="0" cellspacing="0">
      <!--DWLayoutTable-->
      <tr>
        <td width="5%" height="25" align="center" valign="middle" bgcolor="#F5F5F5">编号</td>
        <td width="10%" align="center" valign="middle" bgcolor="#F5F5F5">文章分类</td>
        <td align="left" valign="middle" bgcolor="#F5F5F5">文章标题</td>
        <td width="10%" align="center" valign="middle" bgcolor="#F5F5F5">录入员</td>
        <td width="10%" align="center" valign="middle" bgcolor="#F5F5F5">点击量</td>
        <td width="10%" align="center" valign="middle" bgcolor="#F5F5F5">操作</td>
      </tr>
    </table>
	</td>
  </tr>
  <tr>
    <td height="20" valign="top"><table width="100%" cellpadding="0" cellspacing="0" style="border-top:1px solid #A5B9C9;">
	
	<%
	if request.QueryString("newsclass")<>"" then
	sql="select * from clkj_News where clkj_new_lei='"&request.QueryString("newsclass")&"' order by clkj_newsid desc"
	else
		sql="select * from clkj_News order by clkj_newsid desc"
	end if
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
        <td width="5%" height="25" align="center" valign="middle" style="border-bottom:1px solid #A5B9C9;border-right:1px solid #A5B9C9;"><%=rs("clkj_newsid")%></td>
        <td width="10%" align="center" valign="middle" style="border-bottom:1px solid #A5B9C9;border-right:1px solid #A5B9C9;">
		
		<%
		if rs("clkj_new_lei")="faq" then
		response.write "常见问题"
		else if rs("clkj_new_lei")="news" then
		response.write "公司新闻"
		else
		response.write "关于我们"
		end if
		end if
		
		%></td>
        <td align="left" valign="middle" style="border-bottom:1px solid #A5B9C9;border-right:1px solid #A5B9C9; text-indent:1em;"><%=rs("clkj_news_Title")%></td>
        <td width="10%" align="center" valign="middle" style="border-bottom:1px solid #A5B9C9;border-right:1px solid #A5B9C9;"><%=rs("clkj_news_user")%></td>
        <td width="10%" align="center" valign="middle" style="border-bottom:1px solid #A5B9C9;border-right:1px solid #A5B9C9;"><%=rs("clkj_news_time_hits")%></td>
        <td width="10%" align="center" valign="middle" style="border-bottom:1px solid #A5B9C9;border-right:1px solid #A5B9C9;"><a href=nimda_news.asp?clkj_newsid=<%=rs("clkj_newsid")%>&Edit=N_E>修改</a>|<a href=nimda_function.asp?clkj_newsid=<%=rs("clkj_newsid")%>&Class=news_Del onClick="return confirm('是否将此删除?');">删除</a></td>
      </tr>
		<%
 rs.MoveNext
 next
 end if
 %><tr><td height=25 colspan="7" align="right" valign="middle" bgcolor="#F5F5F5"><div> 共&nbsp;<font color="#ff0000"><%=proCount%></font>&nbsp;条, 当前第：<font color="#ff0000"><%=intCurPage%></font>/<font color="#ff0000"><%=rs.PageCount%></font>页
                <% if intCurPage<>1 then %>
                <a href="nimda_news_mange.asp?cla=<%=nb%>&amp;ToPage=1&newsclass=<%=request.QueryString("newsclass")%>" class="redlink">首页</a>&nbsp;|&nbsp;<a href="nimda_news_mange.asp?cla=<%=nb%>&amp;ToPage=<%=intCurPage-1%>&newsclass=<%=request.QueryString("newsclass")%>" class="redlink">上一页</a>&nbsp;|
                <%else%>
                首页&nbsp;|&nbsp;上一页&nbsp;|
                <% end if%>
                <%if intCurPage<>rs.PageCount then %>
                <a href="nimda_news_mange.asp?cla=<%=nb%>&amp;ToPage=<%=intCurPage+1%>&newsclass=<%=request.QueryString("newsclass")%>" class="redlink">下一页</a>&nbsp;| <a href="nimda_news_mange.asp?cla=<%=nb%>&amp;ToPage=<%=rs.PageCount%>&newsclass=<%=request.QueryString("newsclass")%>" class="redlink"> 末页</a>&nbsp;|&nbsp;
                <%else%>
                <font color="#999999">下一页&nbsp;|&nbsp;末页&nbsp;|&nbsp; </font>
                <% end if%></div></td></tr>
    </table>
</body>
</html>
