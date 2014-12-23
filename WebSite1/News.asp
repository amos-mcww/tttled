<!--#include file="Clkj_Inc/clkj_inc.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<meta http-equiv="X-UA-Compatible" content="IE=EmulateIE7" />
<link rel="shortcut icon" href="Clkj_Images/back_image/favicon.ico" mce_href="Clkj_Images/back_image/favicon.ico" type="image/x-icon" />
<link rel="icon" href="Clkj_Images/back_image/favicon.ico"  mce_href="Clkj_Images/back_image/favicon.ico" type="image/x-icon" />
<title>News_<%=clkj_co_name%></title>
<meta name="keywords" content="<% call navdes(27,0)%>" />
<meta name="description" content="<% call navdes(27,1)%>" />
<link href="Clkj_Template/Clkj_moban_1/Clkj_css/Trade.css" rel="stylesheet" type="text/css" />
</head>
<body>

<div class="sem">
<!--#include file="Clkj_Template/Clkj_moban_1/Clkj_Include/Trade_top.asp"-->
<div class="sem-tag"><a href="/">Home</a> » News</div>
<div class="cb"></div>
    <div class="sem-mid">
 <!--#include file="Clkj_Template/Clkj_moban_1/Clkj_Include/Trade_left.asp"-->
        <div class="sem-mid-cont">
        	<div class="sem-mid-cont-bt">News</div>
            <div class="cb"></div>
            <div class="sem-mid-cont-new">
            <ul>

              <%
set rs=server.CreateObject("adodb.recordset")
sql="select * from clkj_News where clkj_new_lei='news' order by clkj_newsid desc"
rs.open sql,conn,1,1
	If rs.eof Then
	Else
		rs.pagesize=20
		page=request("page")
	If Not Isnumeric(page) or page="" Then
		page=1
	else
		page=cint(page)
	End if
if page<1 then page=1
	if page>rs.pagecount then page=rs.pagecount
	rs.AbsolutePage = page
	for i=1 to rs.pagesize
%>

  <li class="sem-mid-left-1-1-2"><span class="sl"><a href="N_view.asp?nid=<%=rs("clkj_newsid")%>" target="_blank"><%=rs("clkj_news_Title")%></a></span> <span class="sr"><%=rs("clkj_news_time")%></span></li>

  <%
rs.movenext
	if rs.eof then
	Exit For
	End if
	next
end if
%>
<li><span class="sr"> <%if page<>1 then %>
    <a href='?page=<%=page+1%>'><img src="Clkj_Template/Clkj_moban_1/Clkj_Images/prev.gif" align="absmiddle" border="0"></a>
    <%end if%>
    <%
					if page>2 then s1=page-2 else s1=1
					if page<rs.pagecount-2 then s2=page+1 else s2=rs.pagecount
					if s1>=2 then response.write "<a href='?page="&(rs.pagecount-rs.pagecount)+1&"'>1</a>.."
					for j=s1 to s2
					if j=page then
					response.write "<a href=?'page="&j&"'><span class='sp_1'>"&j&"</span></a>&nbsp;&nbsp;"
					else
					response.write "<a href='?page="&j&"'><span class='sp_2'>"&j&"</span></a>&nbsp;&nbsp;"
					end if
					next
					if s2<rs.pagecount then response.write "..<a href='?page="&rs.pagecount&"'><span class='sp_2'>"&rs.pagecount&"</span></a>"
					%>
    <%if page<rs.pagecount then %>
    <a href='?page=<%=page-1%>'><img src="Clkj_Template/Clkj_moban_1/Clkj_Images/next.gif" border="0" align="absmiddle"></a>
    <%
					end if
					rs.close
					%></span></li>
  </ul>
            </div>
            <div class="cb"></div>
        </div>

    </div>
    <div class="cb"></div>
 <!--#include file="Clkj_Template/Clkj_moban_1/Clkj_Include/Trade_bot.asp"-->
</div>


<!-- 统计代码开始 -->
<script language="JavaScript" src="http://mmmled.com/__ip_stat++/smartstat.asp?siteID=1" type="text/JavaScript"></script>
<!-- 统计代码结束 -->



</body>
</html>