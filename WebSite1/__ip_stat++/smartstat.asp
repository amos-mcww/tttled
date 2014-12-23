<!--#include file="inc/inc.config.asp"-->
<%
Function finddir(filepath)
	finddir=left(filepath,instrRev(filepath,"/"))
End Function

style = Request("style")
FromUrl = replace(Trim(Request("fromurl")&""),",","")
theurl	= "http://" & Request.ServerVariables("http_host") & finddir(Request.ServerVariables("url"))
%>
var smartstat_siteid = '<%=siteid%>'
var smartstat_style	= '<%=style%>';
var smartstat_fromurl = '<%=FromUrl%>';
var smartstat_url = '<%=theurl%>';
var smartstat_ndate	= new Date();
var smartstat_tzone	= 0 - smartstat_ndate.getTimezoneOffset()/60;
var smartstat_tcolor = screen.colorDepth;
var smartstat_sSize	= screen.width + ',' + screen.height;
var smartstat_referrer = escape(document.referrer);
var smartstat_outstr = '<script language=javascript src=' + smartstat_url 
			  + 'stat.asp?style=' + smartstat_style 
			  + '&siteid=' + smartstat_siteid 
			  + '&tzone=' + smartstat_tzone 
			  + '&tcolor=' + smartstat_tcolor
			  + '&sSize=' + smartstat_sSize
			  + '&fromurl=' + smartstat_fromurl
			  + '&referrer=' + smartstat_referrer
			  + '><\/script>';
document.write(smartstat_outstr);
<%
if smart_SaveOnline then
%>
function smartstatimgon(reftime){
	var ttime=new Date();
	var smartstat_img=new Image();
	smartstat_img.src=smartstat_url+'inc.online.asp?siteid='+smartstat_siteid+'&o='+ttime.getDate+ttime.getMinutes +ttime.getSeconds;
	var smartstatimgtimeout=setTimeout('smartstatimgon('+reftime+');',reftime);
}
smartstatimgon(<%=smart_CheckOnline*1000%>);
<%
end if
%>