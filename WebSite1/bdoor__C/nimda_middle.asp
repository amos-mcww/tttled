<!--#include file="../Clkj_Inc/inc.asp"-->
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>SemCms外贸网站管理系统</title>
<LINK href="backdoor.css" type="text/css" rel="stylesheet">

<script type="text/javascript">
//弹出层
function openLayer(objId,conId){
var arrayPageSize   = getPageSize();//调用getPageSize()函数
var arrayPageScroll = getPageScroll();//调用getPageScroll()函数
if (!document.getElementById("popupAddr")){
//创建弹出内容层
var popupDiv = document.createElement("div");
//给这个元素设置属性与样式
popupDiv.setAttribute("id","popupAddr")
popupDiv.style.position = "absolute";
popupDiv.style.border = "1px solid #ccc";
popupDiv.style.background = "#fff";
popupDiv.style.zIndex = 99;
//创建弹出背景层
var bodyBack = document.createElement("div");
bodyBack.setAttribute("id","bodybg")
bodyBack.style.position = "absolute";
bodyBack.style.width = "100%";
bodyBack.style.height = (arrayPageSize[1] + 35 + 'px');
bodyBack.style.zIndex = 98;
bodyBack.style.top = 0;
bodyBack.style.left = 0;
bodyBack.style.filter = "alpha(opacity=50)";
bodyBack.style.opacity = 0.5;
bodyBack.style.background = "#ddf";
//实现弹出(插入到目标元素之后)
var mybody = document.getElementById(objId);
insertAfter(popupDiv,mybody);//执行函数insertAfter()
insertAfter(bodyBack,mybody);//执行函数insertAfter()
}
//显示背景层
document.getElementById("bodybg").style.display = "";
//显示内容层
var popObj=document.getElementById("popupAddr")
popObj.innerHTML = document.getElementById(conId).innerHTML;
popObj.style.display = "";
//让弹出层在页面中垂直左右居中(统一)
// popObj.style.width  = "600px";
// popObj.style.height = "400px";
// popObj.style.top  = arrayPageScroll[1] + (arrayPageSize[3] - 35 - 400) / 2 + 'px';
// popObj.style.left = (arrayPageSize[0] - 20 - 600) / 2 + 'px';
//让弹出层在页面中垂直左右居中(个性)
var arrayConSize=getConSize(conId)
popObj.style.top  = arrayPageScroll[1] + (arrayPageSize[3] - arrayConSize[1]) / 2-50 + 'px';
popObj.style.left = (arrayPageSize[0] - arrayConSize[0]) / 2 -30 + 'px';
}
//获取内容层内容原始尺寸
function getConSize(conId){
var conObj=document.getElementById(conId)
conObj.style.position = "absolute";
conObj.style.left=-1000+"px";
conObj.style.display="";
var arrayConSize=[conObj.offsetWidth,conObj.offsetHeight]
conObj.style.display="none";
return arrayConSize;
}
function insertAfter(newElement,targetElement){//插入
var parent = targetElement.parentNode;
if(parent.lastChild == targetElement){
parent.appendChild(newElement);
}
else{
parent.insertBefore(newElement,targetElement.nextSibling);
}
}
//获取滚动条的高度
function getPageScroll(){
var yScroll;
if (self.pageYOffset) {
yScroll = self.pageYOffset;
} else if (document.documentElement  &&  document.documentElement.scrollTop){
yScroll = document.documentElement.scrollTop;
} else if (document.body) {
yScroll = document.body.scrollTop;
}
arrayPageScroll = new Array('',yScroll)
return arrayPageScroll;
}
//获取页面实际大小
function getPageSize(){
var xScroll,yScroll;
if (window.innerHeight  &&  window.scrollMaxY){
xScroll = document.body.scrollWidth;
yScroll = window.innerHeight + window.scrollMaxY;
} else if (document.body.scrollHeight > document.body.offsetHeight){
sScroll = document.body.scrollWidth;
yScroll = document.body.scrollHeight;
} else {
xScroll = document.body.offsetWidth;
yScroll = document.body.offsetHeight;
}
var windowWidth,windowHeight;
if (self.innerHeight) {
windowWidth = self.innerWidth;
windowHeight = self.innerHeight;
} else if (document.documentElement  &&  document.documentElement.clientHeight) {
windowWidth = document.documentElement.clientWidth;
windowHeight = document.documentElement.clientHeight;
} else if (document.body) {
windowWidth = document.body.clientWidth;
windowHeight = document.body.clientHeight;
}
var pageWidth,pageHeight
if(yScroll < windowHeight){
pageHeight = windowHeight;
} else {
pageHeight = yScroll;
}
if(xScroll < windowWidth) {
pageWidth = windowWidth;
} else {
pageWidth = xScroll;
}
arrayPageSize = new Array(pageWidth,pageHeight,windowWidth,windowHeight)
return arrayPageSize;
}
//关闭弹出层
function closeLayer(){
document.getElementById("popupAddr").style.display = "none";
document.getElementById("bodybg").style.display = "none";
return false;
}
</script>
<script type="text/javascript">
//对"拖动点"定义：onMousedown="StartDrag(this)" onMouseup="StopDrag(this)" onMousemove="Drag(this)"即可
var move=false,oldcolor,_X,_Y;
function StartDrag(obj){  //定义准备拖拽的函数
obj.setCapture(); //对当前对象的鼠标动作进行跟踪
oldcolor=obj.style.backgroundColor;
obj.style.background="#999";
move=true;
//获取鼠标相对内容层坐标
var parentwin=document.getElementById("popupAddr");
_X=parentwin.offsetLeft-event.clientX
_Y=parentwin.offsetTop-event.clientY
}
function Drag(obj){        //定义拖拽函数
if(move){
var parentwin=document.getElementById("popupAddr");
parentwin.style.left=event.clientX+_X;
parentwin.style.top=event.clientY+_Y;
}
}
function StopDrag(obj){   //定义停止拖拽函数
obj.style.background=oldcolor;
obj.releaseCapture(); //停止对当前对象的鼠标跟踪
move=false;
}
</script>
</head>
<body bgcolor="#E9F1FE">
<table width="100%" cellspacing="0" cellpadding="0">
  <tr>
    <td height="39" background="../Clkj_Images/back_image/back_images_title.gif" widht=100%>&nbsp;<b>欢迎使SemCms外贸网站管理系统</b></td>
  </tr>
  <tr>
    <td height="30" bgcolor="#E9F1FE" widht=100%>&nbsp;版本号: SemCms 4.3 - 2013.12.16</td>
  </tr>
</table>
<br>
<table width="100%" cellspacing="0" cellpadding="0">
  <tr>
  </tr>
</table>
<br>
<table width="100%" cellspacing="0" cellpadding="0">
  <tr>
    <td height="39" background="../Clkj_Images/back_image/back_images_title.gif" widht=100% colspan=2>&nbsp;<b>系统组件信息</b></td>
  </tr>
  <tr>
    <td height="30" bgcolor="#E9F1FE" widht=100% width=50%>&nbsp;Fso组件
    	<%IF isobjinstalled("Scripting.FileSystemObject")Then
	 response.write"<font color=red>->>支持</font>"
	 Else
	 response.write"<font color=blue>不支持</font>"
	 End IF
	 %></td>
    <td height="30" bgcolor="#E9F1FE" widht=100% width=50%>&nbsp;AspJpeg组件  
	<%
	
	IF isobjinstalled("Persits.Jpeg") Then
	 response.write"<font color=red>->>支持</font>"
	 Else
	 response.write"<font color=blue>不支持</font>"
	 End IF
	 %>  <a href="http://tradecms.zjxyk.com/TradeCmsdown/AspJpeg.rar">下载</a></td>
  </tr>
  <tr>
    <td height="30" bgcolor="#E9F1FE" widht=100% width=50%>adodb.stream组件<a href="http://www.speed-web.cn/"></a>	
	<%IF isobjinstalled("adodb.stream") Then
	 response.write"<font color=red>->>支持</font>"
	 Else
	 response.write"<font color=blue>不支持</font>"
	 End IF
	 %></td>
    <td height="30" bgcolor="#E9F1FE" widht=100% width=50%>&nbsp;Jmail组件<a href="http://www.tradedoor.net/"></a>	
	<%IF isobjinstalled("JMail.SmtpMail") Then
	 response.write"<font color=red>->>支持</font>"
	 Else
	 response.write"<font color=blue>不支持</font>"
	 End IF
	 %> <a href="http://tradecms.zjxyk.com/TradeCmsdown/jmail44_free.zip">下载</a> &nbsp;&nbsp;&nbsp; 系统自带CDO邮件组件 <%IF isobjinstalled("CDO.Message") Then
	 response.write"<font color=red>->>支持</font>"
	 Else
	 response.write"<font color=blue>不支持</font>"
	 End IF
	 %></td>
     
 
  </tr>
</table>

<br>
<table width="100%" cellspacing="0" cellpadding="0">
  <tr>
    <td height="39" background="../Clkj_Images/back_image/back_images_title.gif" widht=100% colspan=2>&nbsp;<b>联系我们</b></td>
  </tr>
  <tr>
    <td height="30" bgcolor="#E9F1FE" widht=100% width=50%>&nbsp;产品名称：SemCms</td>
    <td height="30" bgcolor="#E9F1FE" widht=100% width=50%>&nbsp;项目开发：QQ 1181698019(黑蚂蚁)</td>
  </tr>
  <tr>
    <td height="30" bgcolor="#E9F1FE" widht=100% width=50%>&nbsp;官方网站：<a href="http://www.sem-cms.com/">http://www.sem-cms.com/</a><a href="http://www.speed-web.cn/"> </a></td>
    <td height="30" bgcolor="#E9F1FE" widht=100% width=50%>&nbsp;帮助中心：<a href="http://www.sem-cms.com/">http://www.sem-cms.com/talk/</a></td>
  </tr>
</table>
<script type="text/javascript"> 
<%
Set fso=Server.CreateObject("Scripting.FileSystemObject")
IF Not fso.FolderExists(Server.MapPath("../clkj_admin")) Then
response.write""
else
%>
window.onload=function(){openLayer('test3','test_con3');} 
<%end if%>
 
 <%
Set fso=Server.CreateObject("Scripting.FileSystemObject")
IF Not fso.FolderExists(Server.MapPath("../Clkj_DaTa")) Then
response.write""
else
%>
window.onload=function(){openLayer('test3','test_con3');} 
<%end if%> 
</script> 

<br>
<br>
<br>
<br>
<br>
</body>
</html>
