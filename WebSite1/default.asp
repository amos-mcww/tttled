<%
clkj_new_lei="About us"
clkj_new_les="index"
%>
<!--#include file="Clkj_Inc/clkj_inc.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="X-UA-Compatible" content="IE=EmulateIE7" />
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title><%=clkj_index_key%> - <%=clkj_co_name%></title>
<meta name="keywords" content="<%=clkj_index_key%>" />
<meta name="description" content="<%=clkj_index_des%>" />
<%=clkj_meta%>
<link href="Clkj_Template/Clkj_moban_1/Clkj_css/Trade.css" rel="stylesheet" type="text/css" />
<script src="Clkj_Inc/js/adv-2.js" type="text/javascript"></script>
<script src="Clkj_Inc/js/adv-1.js" type="text/javascript"></script>
</head>

<body>
<div class="sem">
<!--#include file="Clkj_Template/Clkj_moban_1/Clkj_Include/Trade_top.asp"-->
    <div class="sem-mid">
 <!--#include file="Clkj_Template/Clkj_moban_1/Clkj_Include/Trade_left.asp"-->
        <div class="sem-mid-mid">
        	<div class="sem-mid-mid-1"> <!--#include file="Clkj_Template/Clkj_moban_1/Clkj_Include/Trade_adv.asp"--></div>
            <div class="cb"></div>
            <div class="sem-mid-mid-bt">New Products</div>
            <div class="cb"></div>
            <div class="sem-mid-mid-2">
			 <%=index_new%>             
            </div>
        </div>
        <div class="sem-mid-right">
        	<div class="sem-mid-left-bt">Hot products</div>
            <div class="cb"></div>
            <div class="sem-mid-right-1">
     <%=index_Hot%>
            </div>
           
        
        </div>
    
    </div>
    <div class="cb"></div>

<!--#include file="Clkj_Template/Clkj_moban_1/Clkj_Include/Trade_bot.asp"-->


</div>



</body>
</html>
