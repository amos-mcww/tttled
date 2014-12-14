<!--#include file="Nimda_function.asp"-->
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<link href="backdoor.css" type="text/css" rel="stylesheet">
<title>产品管理</title>

</head>
<script type="text/javascript">
function CheckAll(form)
{
   for (var i=0;i<form.elements.length;i++)
   {
      var e = form.elements[i];
      e.checked = true;
   }
}
function CAll(form)
{
   for (var i=0;i<form.elements.length;i++)
   {
      var e = form.elements[i];
      e.checked =false;
   }
}
function ConfirmDel(message)
{
   if (confirm(message))
   {
      document.formDel.submit()
   }
}
function open4()
{
window.location.href="nimda_products.asp?search="+document.formDel.searchs.value;
}

</script>
<body>
  <form name="formDel" method="post" action="nimda_function.asp?Class=P_Del&clkj_SmallClassID=<%=request("clkj_SmallClassID")%>&clkj_BigClassID=<%=request("clkj_BigClassID")%>&clkj_BigClassName=<%=request.QueryString("clkj_BigClassName")%>&clkj_SmallClassName=<%=request.QueryString("clkj_SmallClassName")%>&sf=<%=request.QueryString("sf")%>&ToPage=<%if request("ToPage")="" then response.write "1"%><%if request("ToPage")<>"" then response.write request("ToPage")%>">
<table width="100%" cellpadding="0" cellspacing="0">
  <tr>
    <td height="39" align="left" valign="middle" class="tr_bg"><strong style="float:left">&nbsp;产品管理</strong><span style="float:right">
<input onClick="CheckAll(this.form)" type="button" id="submitAllSearch" value="全选"/> <input onClick="CAll(this.form)" type="button" id="submitAllSearch" value="取消"/>
<input name="batch" type="submit" value="删除所选" onClick="ConfirmDel('是否确定删除？删除后不能恢复!');" /> </span> <span style=" float:right;">产品搜索：<input name="searchs" type="text" value="请输入产品标题" onMouseOver="this.select()" style="line-height:18px; height:20px;"> <a href="#" onClick="open4()">查询</a>&nbsp; </span></td>
  </tr>
    <tr>
    <td height="30"bgcolor="#E7EFF8">&nbsp;栏目: &nbsp;<%call big_lie()%>
    &nbsp;&nbsp;[<a href="nimda_product.asp">增加产品</a>]</td>
  </tr>
  <%if request("clkj_BigClassID")<>"" then%>
  <tr><td width="100%" height="25" bgcolor="#E7E7E7" style="text-indent:2em;" align="left" valign="middle"><%call small_lie()%></td>
  </tr>
  <%end if%>
</table>

<table width="100%" border="0" cellpadding="0" cellspacing="0">
      <!--DWLayoutTable-->
      <tr>
        <td width="77" height="25" align="center" valign="middle" bgcolor="#F5F5F5"><strong>编号</strong></td>
		<td width="77" height="25" align="center" valign="middle" bgcolor="#F5F5F5"><strong>排序</strong></td>
        <td width="136" align="center" valign="middle" bgcolor="#F5F5F5"><strong>产品分类</strong></td>
        <td width="420" align="left" valign="middle" bgcolor="#F5F5F5" style="text-indent:0.5em;"><strong>产品名称</strong></td>
        <td width="134" align="center" valign="middle" bgcolor="#F5F5F5"><strong>是否推荐</strong></td>
        <td width="240" align="center" valign="middle" bgcolor="#F5F5F5"><strong>时间</strong></td>
        <td width="158" align="center" valign="middle" bgcolor="#F5F5F5"><strong>操作</strong>
		</td>
  </tr>
    </table>
	</td>
  </tr>

  <tr>
    <td height="20" valign="top"><table width="100%" cellpadding="0" cellspacing="0" style="border-top:1px solid #A5B9C9;">

	<%
	if request("clkj_BigClassID")<>""and request.QueryString("clkj_BigClassName")<>"" and request.QueryString("clkj_SmallClassName")<>"" and request("clkj_SmallClassID")<>"" then
	sql="select * from clkj_Products where clkj_SmallClassID ="&request("clkj_SmallClassID")&" order by clkj_paixu,clkj_prid asc"
	Elseif request("clkj_BigClassID")<>""and request.QueryString("clkj_BigClassName")<>"" and request.QueryString("clkj_SmallClassName")="" and request("clkj_SmallClassID")="" then
	sql="select * from clkj_Products where clkj_BigClassID ="&request("clkj_BigClassID")&" order by clkj_paixu,clkj_prid asc"
	Elseif request("hot")=1 then
	sql="select * from clkj_Products where clkj_hot=1 order by clkj_paixu,clkj_prid asc"
	else
	sql="select * from clkj_Products where clkj_prtitle like '%"&request.QueryString("search")&"%' order by clkj_paixu,clkj_prid asc"

	End IF
	set rs=server.CreateObject("adodb.recordset")
	rs.open sql,conn,1,3
IF Not rs.eof Then
proCount=rs.recordcount
	rs.PageSize=17		  '定义每页显示数目
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
        <td width="78" height="25" align="center" valign="middle" bgcolor="#F5F5F5" style="border-bottom:1px solid #A5B9C9;border-right:1px solid #A5B9C9;"><%=rs("clkj_prid")%></td>
		<td width="78" height="25" align="center" valign="middle" style="border-bottom:1px solid #A5B9C9;border-right:1px solid #A5B9C9; color:#ff0000;"><%=rs("clkj_paixu")%></td>
        <td width="133" align="center" valign="middle" style="border-bottom:1px solid #A5B9C9;border-right:1px solid #A5B9C9;">
		<% if rs("clkj_SmallClassName")="" then
		    response.write rs("clkj_SmallClassName")
			else
			response.write rs("clkj_BigClassName")
			end if
		%></td>
        <td width="422" align="left" valign="middle" style="border-bottom:1px solid #A5B9C9;border-right:1px solid #A5B9C9; text-indent:0.5em;"><%=rs("clkj_prtitle")%></td>
        <td width="132" align="center" valign="middle" style="border-bottom:1px solid #A5B9C9;border-right:1px solid #A5B9C9;"><%if rs("clkj_hot")=1 then %><font color="#FF0000">是</font><%else%>否<%end if%>
         </td>
        <td width="240" align="center" valign="middle" style="border-bottom:1px solid #A5B9C9;border-right:1px solid #A5B9C9;"><%=rs("clkj_time")%></td>
        <td width="158" align="center" valign="middle" style="border-bottom:1px solid #A5B9C9;border-right:1px solid #A5B9C9;"><a href=nimda_product.asp?clkj_prid=<%=rs("clkj_prid")%>&Edit=P_E&clkj_BigClassID=<%=rs("clkj_BigClassID")%>&clkj_SmallClassID=<%=rs("clkj_SmallClassID")%>&ToPage=<%=Request.QueryString("ToPage")%>&sf=<%=request.QueryString("sf")%>>修改</a>
          <input type="checkbox" name="delid" value="<%=rs("clkj_prid")%>">
        </td>
      </tr>
		<%
 rs.MoveNext
 next
 end if
 %>

 <%
dim link
if request("clkj_BigClassID")<>"" and request("clkj_SmallClassID")<>"" then

	link="?clkj_BigClassID="&request("clkj_BigClassID")&"&clkj_BigClassName="&request("clkj_BigClassName")&"&clkj_SmallClassID="&request("clkj_SmallClassID")&"&clkj_SmallClassName="&request("clkj_SmallClassName")&"&sf=s&"

elseif request("clkj_BigClassID")<>"" and request("clkj_SmallClassID")="" then

	link="?clkj_BigClassID="&request("clkj_BigClassID")&"&clkj_BigClassName="&request("clkj_BigClassName")&"&sf=b&"
elseif request("hot")=1 then
	link="?hot=1&"
else
	link="?"
end if

%>
 <tr><td height=25 colspan="7" align="right" valign="middle" bgcolor="#F5F5F5"><div> 共&nbsp;<font color="#ff0000"><%=proCount%></font>&nbsp;条, 当前第：<font color="#ff0000"><%=intCurPage%></font>/<font color="#ff0000"><%=rs.PageCount%></font>页
                <% if intCurPage<>1 then %>
                <a href="<%=link%>ToPage=1" class="redlink">首页</a>&nbsp;|&nbsp;<a href="<%=link%>ToPage=<%=intCurPage-1%>" class="redlink">上一页</a>&nbsp;|
                <%else%>
                首页&nbsp;|&nbsp;上一页&nbsp;|
                <% end if%>
              <%if intCurPage<>rs.PageCount then %>
                <a href="<%=link%>ToPage=<%=intCurPage+1%>" class="redlink">下一页</a>&nbsp;| <a href="<%=link%>ToPage=<%=rs.PageCount%>" class="redlink"> 末页</a>&nbsp;|&nbsp;
                <%else%>
                <font color="#999999">下一页&nbsp;|&nbsp;末页&nbsp;|&nbsp; </font>
                <% end if%></div></td></tr>
    </table>
</form>
<form name="form1" method="post">
<table width="100%" cellpadding="0" cellspacing="0"><tr><td align="right" height="20">
跳转到第<input name="go" type="text" id="go" style="width:80px;height:20px;text-align:center;vertical:middle" size="80">
页
<input type="button" name="skip" value="查询" onClick="window.location.href='<%=link%>ToPage='+document.getElementById('go').value">


</td></tr></table>
</form>
</body>
</html>
