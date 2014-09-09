<!--#include file="Clkj_Inc/clkj_inc.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<meta http-equiv="X-UA-Compatible" content="IE=EmulateIE7" />
<title><%=clkj_news_Title%> - <%=clkj_co_name%></title>
<meta name="description" content="<% call navdes(27,1)%>"/>
<link href="Clkj_Template/Clkj_moban_1/Clkj_css/Trade.css" rel="stylesheet" type="text/css" />
</head>
<body>
<div class="sem">
<!--#include file="Clkj_Template/Clkj_moban_1/Clkj_Include/Trade_top.asp"-->
<div class="sem-tag"><a href="/">Home</a> » <%=clkj_news_Title%></div>
<div class="cb"></div>
    <div class="sem-mid">
 <!--#include file="Clkj_Template/Clkj_moban_1/Clkj_Include/Trade_left.asp"-->
        <div class="sem-mid-cont">
        	<div class="sem-mid-cont-bt"><%=clkj_news_Title%></div>
            <div class="cb"></div>
            <div class="sem-mid-cont-1"><%=clkj_news_content%></div>
            <div class="cb"></div>
            <div class="sem-mid-cont-1"><a href="javascript:window.close()"><img src="Clkj_Template/Clkj_moban_1/Clkj_images/x.gif" border="0" align="absmiddle" /></a>&nbsp;&nbsp;<a href="javascript:window.print()"><img src="Clkj_Template/Clkj_moban_1/Clkj_images/print.gif" border="0" align="absmiddle" /></a> <img src="Clkj_Template/Clkj_moban_1/Clkj_images/clock.gif" align="absmiddle" /> <%=clkj_news_time%></div>
            
            <div class="sem-mid-cont-bt">Related news</div>
            <div class="cb"></div>
            <div class="sem-mid-cont-newx"><ul><%=RndNews%></ul></div>
            <div class="cb"></div>
        </div>
    
    </div>
    <div class="cb"></div>
 <!--#include file="Clkj_Template/Clkj_moban_1/Clkj_Include/Trade_bot.asp"-->
</div>


</body>
</html>
