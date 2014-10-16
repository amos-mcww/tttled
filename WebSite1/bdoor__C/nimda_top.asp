<!--#include file="../Clkj_Inc/inc.asp"-->
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>tttled Top</title>
<LINK href="backdoor.css" type="text/css" rel="stylesheet">
</head>
<body bgcolor="#CCCCCC">
<table width="100%" height="91" border="0" cellpadding="0" cellspacing="0" background="../Clkj_Images/back_image/back_images_top.gif" >
  <!--DWLayoutTable-->
<tr>
<td width="255" rowspan="2" align="center" valign="middle"><img src="../Clkj_Images/back_image/logo.gif" alt="SEMCMS"></td>
<td width="996" height="52">&nbsp;<font color="#FF9900"><b><%=session("username")%></b></font>&nbsp;[<a href="nimda_function.asp?class=Out" target="_parent">安全退出</a>，<%if session("clkj_flg")=1 then%><a href="nimda_user.asp" target="mainFrame">用户管理</a><%else%><a href="nimda_user.asp?clkj_id=<%=session("clkj_id")%>&Edit=B_E#xy" target="mainFrame">用户管理</a><%end if%>] <font color="#FF9900">今天是:<%=year(now())%>年<%=month(now())%>月<%=day(now())%>号 星期<% if WeekDay (date())-1 =0 then %>天<% else if WeekDay (date())-1 =1 then %>一<% else if WeekDay (date())-1 =2 then %>二<% else if WeekDay (date())-1 =3 then %>三<% else if WeekDay (date())-1 =4 then %>四<% else if WeekDay (date())-1 =5 then %>五<% else if WeekDay (date())-1 =6 then %>六<% end if%><% end if%><% end if%><% end if%><% end if%><% end if%><% end if%></font></td>
</tr>
<tr>
  <td height="39" align="left" valign="middle" ><table width="700" border="0" cellpadding="0" cellspacing="0">
  <!--DWLayoutTable-->
  <tr>

    <td width="90" align="center" valign="middle" background="../Clkj_Images/back_image/back_top_1.png" height="39" ><a href="nimda_products.asp" target="mainFrame"><font color="#FFFFFF">产品管理</font></a></td>
	    <td width="10" valign="top"><!--DWLayoutEmptyCell-->&nbsp;</td>
    <td width="90" align="center" valign="middle" background="../Clkj_Images/back_image/back_top_1.png"><a href="http://www.sem-cms.com/" target="mainFrame"><font color="#FFFFFF">帮助中心</font></a></td>
	<td width="10" valign="top"><!--DWLayoutEmptyCell-->&nbsp;</td>
        <td width="90" align="center" valign="middle" background="../Clkj_Images/back_image/back_top_1.png"><a href="http://www.sem-cms.com/guanyu.html" target="mainFrame"><font color="#FFFFFF">系统说明</font></a></td>
	<td width="10" valign="top"><!--DWLayoutEmptyCell-->&nbsp;</td>
    <td width="400">&nbsp;</td>
  </tr>
</table></td>
  </tr>
</table>

</body>
</html>
