<!--#include file="Clkj_Inc/clkj_inc.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<meta http-equiv="X-UA-Compatible" content="IE=EmulateIE7" />
<link rel="shortcut icon" href="Clkj_Images/back_image/favicon.ico" mce_href="Clkj_Images/back_image/favicon.ico" type="image/x-icon" /> 
<link rel="icon" href="Clkj_Images/back_image/favicon.ico"  mce_href="Clkj_Images/back_image/favicon.ico" type="image/x-icon" />
<title><%=clkj_prtitle%>_<%=clkj_co_name%></title>
<meta name="keywords" content="<%=clkj_prkey%>" />
<meta name="description" content="<%=clkj_prprdes%>" />
<link href="Clkj_Template/Clkj_moban_1/Clkj_css/Trade.css" rel="stylesheet" type="text/css" />
<link rel="stylesheet" href="Clkj_Inc/js/lightbox.css" type="text/css" media="screen" />
</head>
<body>
<div class="sem"> 
  <!--#include file="Clkj_Template/Clkj_moban_1/Clkj_Include/Trade_top.asp"-->
  <div class="sem-tag"><a href="/">Home</a> » <%=clkj_prtitle%></div>
  <div class="cb"></div>
  <div class="sem-mid"> 
    <!--#include file="Clkj_Template/Clkj_moban_1/Clkj_Include/Trade_left.asp"-->
    <div class="sem-mid-cont">
      <div class="sem-mid-view-1 sem-mid-view-1bg">
        <div class="sem-mid-view-1-c">
          <div class="sem-mid-view-1-left">
            <div class="sem-mid-view-1-left-1">
              <dt class="pic-dtb">
                <%
			js=1
			ProductsP=Split(clkj_prpic,",")
			For Each PicStrss In ProductsP
			IF  trim(PicStrss)<>"" Then
				if js<2 then
			%>
                <%IF isobjinstalled("Persits.Jpeg") Then'判断是否支持图片组件%>
                <a href="<%=Replace(PicStrss,"smallpic","Bigpic")%>" rel="lightbox[roadtrip]"><img src="<%=Replace(PicStrss,"smallpic","Bigpic")%>" alt="<%=clkj_prtitle%>" border="0" /></a>
                <%Else%>
                <a href="<%=PicStrss%>" rel="lightbox[roadtrip]"><img src="<%=PicStrss%>" alt="<%=clkj_prtitle%>" border="0" /></a>
                <%
		    End IF
			End IF       
			End IF
			js=js+1
			next
			%>
              </dt>
            </div>
            <div class="cb"></div>
            <div class="sem-mid-view-1-left-2">
              <ul>
                <%
			ProductsP=Split(clkj_prpic,",")
			For Each PicStrss In ProductsP
			IF trim(PicStrss)<>""  Then
			%>
                <%IF isobjinstalled("Persits.Jpeg") Then'判断是否支持图片组件%>
                <li>
                  <dt class="pic-dts"><a href="<%=Replace(PicStrss,"smallpic","Bigpic")%>" rel="lightbox[roadtrip]"><img src="<%=PicStrss%>" alt="<%=clkj_prtitle%>" border="0" /></a></dt>
                </li>
                <%Else%>
                <li>
                  <dt class="pic-dts"><a href="<%=PicStrss%>" rel="lightbox[roadtrip]"><img src="<%=PicStrss%>" alt="<%=clkj_prtitle%>" border="0" /></a></dt>
                </li>
                <%
			End IF       
			End IF
			next
			%>
              </ul>
            </div>
          </div>
          <div class="sem-mid-view-1-right">
            <div class="sem-mid-view-1-right-1">
              <h1><%=clkj_prtitle%></h1>
            </div>
            <div class="cb"></div>
            <div class="sem-mid-view-1-right-2"><%=clkj_prprdes%></div>
            <div class="cb"></div>
            <div class="sem-mid-view-1-right-2"><a href="Feedback.asp?name=<%=request("pid")%>"><img src="Clkj_Template/Clkj_moban_1/Clkj_images/contactnow.jpg" border="0"></a></div>
            <div class="cb"></div>
            <div class="sem-mid-view-1-right-2"><%=clkj_addthis%></div>
            <div class="cb"></div>
          </div>
        </div>
        <div class="cb"></div>
      </div>
      <div class="cb"></div>
      <div class="sem-mid-view-1-m">
        <div class="sem-mid-view-1-m-1">Detailed description</div>
        <div class="cb"></div>
        <div class="sem-mid-view-1-m-2"><%=clkj_prcontent%></div>
      </div>
      <div class="cb"></div>
      <div class="sem-mid-view-1-m">
        <div class="sem-mid-view-1-m-1"> <span class="sl">Related products </span><span class="sr"><a href="Product.asp?big=<%=clkj_BigClassID%>" target="_blank">More»</a></span></div>
        <div class="cb"></div>
        <div class="sem-mid-view-1-m-2">
          <%call reproduct(clkj_BigClassID,request("pid"))%>
        </div>
      </div>
      <div class="cb"></div>
    </div>
  </div>
  <div class="cb"></div>
  <!--#include file="Clkj_Template/Clkj_moban_1/Clkj_Include/Trade_bot.asp"--> 
  
</div>
<script src="Clkj_Inc/js/prototype.js" type="text/javascript"></script> 
<script src="Clkj_Inc/js/scriptaculous.js?load=effects,builder" type="text/javascript"></script> 
<script src="Clkj_Inc/js/lightbox.js" type="text/javascript"></script>
</body>
</html>
