<!--#include file="Clkj_Inc/clkj_inc.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<meta http-equiv="X-UA-Compatible" content="IE=EmulateIE7" />
<link rel="shortcut icon" href="Clkj_Images/back_image/favicon.ico" mce_href="Clkj_Images/back_image/favicon.ico" type="image/x-icon" />
<link rel="icon" href="Clkj_Images/back_image/favicon.ico"  mce_href="Clkj_Images/back_image/favicon.ico" type="image/x-icon" />
<title><%call ileibie2(1)%> - <%=clkj_co_name%></title>
<meta name="keywords" content="<%call ileibie2(2)%>" />
<meta name="description" content="<%call ileibie2(3)%>" />
<link href="Clkj_Template/Clkj_moban_1/Clkj_css/Trade.css" rel="stylesheet" type="text/css" />
</head>
<body>
<div class="sem">
<!--#include file="Clkj_Template/Clkj_moban_1/Clkj_Include/Trade_top.asp"-->
<div class="sem-tag"><a href="/">Home</a> » <%call ileibie()%></div>
<div class="cb"></div>
    <div class="sem-mid">
 <!--#include file="Clkj_Template/Clkj_moban_1/Clkj_Include/Trade_left.asp"-->
        <div class="sem-mid-cont">
         <%
dim big,small
big=request.QueryString("big")
small=request.QueryString("small")
set rs=server.CreateObject("adodb.recordset")
	if big="" and small="" then
		sql="select top 50 * from clkj_Products order by  clkj_prid desc"
	else if big<>"" and small="" then
		sql="select * from clkj_Products where clkj_BigClassID="&big&" order by clkj_paixu,clkj_prid asc"
	else if small<>"" and big="" then
		sql="select * from clkj_Products where clkj_SmallClassID="&small&" order by clkj_paixu,clkj_prid asc"
	end if
	end if
	end if
rs.open sql,conn,1,1
	If rs.eof Then
	response.write"<div style='margin-top:120px; height:100px; font-size:18px; font-weight:bold; width:760px; text-align:center;'><img src='Clkj_Template/Clkj_moban_1/Clkj_Images/sorry.gif' align='absmiddle'> Sorry, no related products!</div>"

	Else
		rs.pagesize=clkj_config_sl
		page=request("page")
	If Not Isnumeric(page) or page="" Then
		page=1
	else
		page=cint(page)
	End if
if page<1 then page=1
	if page>rs.pagecount then page=rs.pagecount
	rs.AbsolutePage = page
	proCount=rs.recordcount

%>

        	<div class="sem-mid-cont-bt"><span class="sl"><%call ileibie()%></span>
            <span class="sr"><%if page<>1 then %><a href="<%if big=""and small="" then%>?page=<%=page-1%><%elseif  big<>"" and small="" then%>?big=<%=big%>&page=<%=page-1%><%elseif small<>"" and big="" then%>?small=<%=small%>&page=<%=page-1%><%end if%>"><img src="Clkj_Template/Clkj_moban_1/Clkj_Images/prev.gif" border="0" align="absmiddle"></a> <%end if%>

            <%if page<rs.pagecount then %> <a href="<%if big=""and small="" then%>?page=<%=page+1%><%elseif  big<>"" and small="" then%>?big=<%=big%>&page=<%=page+1%><%elseif small<>"" and big="" then%>?small=<%=small%>&page=<%=page+1%><%end if%>"><img src="Clkj_Template/Clkj_moban_1/Clkj_Images/next.gif" border="0" align="absmiddle"></a><%end if%>
            </span>
           </div>
            <div class="cb"></div>
            <div class="sem-mid-cont-1">
			<%
		for i=1 to rs.pagesize
		j=1
		b=Split(rs("clkj_prpic"),",")'取一条图片记录
		For Each CStRss in b
			if j<=1 then
				IF isobjinstalled("Persits.Jpeg") Then

					clkj_prpic=Replace(CStRss,"upfile/","upfile/smallpic/")
				Else
					clkj_prpic=CStRss
				End IF
			End if
			j=j+1
		Next
		pn=rs("clkj_prtitle")
		If Len(pn) <= 24  Then
        	SEMCMS_Br = "<br/><br/>"
        ElseIf Len(pn) <= 49 Then
        	SEMCMS_Br = "<br/>"
        ElseIf Len(pn) > 66 Then
        	SEMCMS_Br = ""
            pn = Left(pn,66) & "..."
        Else
        	SEMCMS_Br= ""
        End if
		%>

           <div class="pic"><div class="pic-div"><dt class="pic-dt"><a href="P_view.asp?pid=<%=rs("clkj_prid")%>" target="_blank"><img src="<%=clkj_prpic%>"  alt="<%=rs("clkj_prtitle")%>" border="0"/></a></dt>
            <dd>
            <dd><a href="P_view.asp?pid=<%=rs("clkj_prid")%>" target="_blank"><%=pn%><%=SEMCMS_Br%></a></dd>
            <dd><a href="Feedback.asp?name=<%=rs("clkj_prid")%>"><img src="Clkj_Template/Clkj_moban_1/Clkj_images/inq.gif" alt="Inquire Now" border="0" /></a></dd>
          </div></div>
          <%
		clkj_prpic=""
		pn=""
		rs.movenext
			if rs.eof then
			Exit For
			End if
			next
		%>
            </div>
            <div class="cb"></div>
            <div class="sem-mid-cont-1"><span class="sr"> Total products <b><%=proCount%></b>, Page <b><%=page%></b> of <b><%=rs.pagecount%></b> <%if page<>1 and  page<>0 then %>
<a href="<%if big=""and small="" then%>?page=<%=page-1%><%elseif  big<>"" and small="" then%>?big=<%=big%>&page=<%=page-1%><%elseif small<>"" and big="" then%>?small=<%=small%>&page=<%=page-1%><%end if%>"><img src="Clkj_Template/Clkj_moban_1/Clkj_Images/prev.gif" border="0" align="absmiddle"></a> <%end if%>
            <%'分页
		if big=""and small="" then

			if page>2 then s1=page-2 else s1=1
			if page<rs.pagecount-2 then s2=page+1 else s2=rs.pagecount
			if s1>=2 then response.write "<a href=?page="&(rs.pagecount-rs.pagecount)+1&"><span class='sp_2'>1</span></a>.."
			for j=s1 to s2
			if j=page then
			response.write "<a href='?page="&j&"'><span class='sp_1'>"&j&"</span></a>&nbsp;"
			else
			response.write "<a href='?page="&j&"'><span class='sp_2'>"&j&"</span></a>&nbsp;"
			end if
			next
			if s2<rs.pagecount then response.write "..<a href='?page="&rs.pagecount&"'><span class='sp_2'>"&rs.pagecount&"</span></a>"

		elseif big<>""and small="" then

			if page>2 then s1=page-2 else s1=1
			if page<rs.pagecount-2 then s2=page+1 else s2=rs.pagecount
			if s1>=2 then response.write "<a href=?big="&big&"&page="&(rs.pagecount-rs.pagecount)+1&"><span class='sp_2'>1</span></a>.."
			for j=s1 to s2
			if j=page then
			response.write "<a href='?big="&big&"&page="&j&"'><span class='sp_1'>"&j&"</span></a>&nbsp;"
			else
			response.write "<a href='?big="&big&"&page="&j&"'><span class='sp_2'>"&j&"</span></a>&nbsp;"
			end if
			next
			if s2<rs.pagecount then response.write "..<a href='?big="&big&"&page="&rs.pagecount&"'><span class='sp_2'>"&rs.pagecount&"</span></a>"

		elseif big=""and small<>"" then

			if page>2 then s1=page-2 else s1=1
			if page<rs.pagecount-2 then s2=page+1 else s2=rs.pagecount
			if s1>=2 then response.write "<a href='?small="&small&"&page="&(rs.pagecount-rs.pagecount)+1&"'><span class='sp_2'>1</span></a>.."
			for j=s1 to s2
			if j=page then
			response.write "<a href='?small="&small&"&page="&j&"'><span class='sp_1'>"&j&"</span></a>&nbsp;"
			else
			response.write "<a href='?small="&small&"&page="&j&"'><span class='sp_2'>"&j&"</span></a>&nbsp;"
			end if
			next
			if s2<rs.pagecount then response.write "..<a href='?small="&small&"&page="&rs.pagecount&"'><span class='sp_2'>"&rs.pagecount&"</span></a>"

		end if
		%>
<%if page<rs.pagecount then %> <a href="<%if big=""and small="" then%>?page=<%=page+1%><%elseif  big<>"" and small="" then%>?big=<%=big%>&page=<%=page+1%><%elseif small<>"" and big="" then%>?small=<%=small%>&page=<%=page+1%><%end if %>"><img src="Clkj_Template/Clkj_moban_1/Clkj_Images/next.gif" border="0" align="absmiddle"></a>
        <%
		end if
		rs.close
		%></span></div>
            <div class="cb"></div>
            <%end if%>
        </div>

    </div>
    <div class="cb"></div>
 <!--#include file="Clkj_Template/Clkj_moban_1/Clkj_Include/Trade_bot.asp"-->
</div>

<!-- 统计代码开始 -->
<script language="JavaScript" src="http://mmmled.com/smartstat/smartstat.asp?siteID=1" type="text/JavaScript"></script>
<!-- 统计代码结束 -->

</body>
</html>
