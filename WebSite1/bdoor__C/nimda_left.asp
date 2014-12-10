<!--#include file="../Clkj_Inc/inc.asp"-->
<html>
<head>
<LINK href="backdoor.css" type="text/css" rel="stylesheet">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title>TradeCms Left</title></head>
<body background="../Clkj_Images/back_image/back_line_bg.png">
<table  cellSpacing=0 cellPadding=0 border=0 background="../Clkj_Images/back_image/back_left_top.png">
  <!--DWLayoutTable-->
<tr>
<td width="92" height="38" align="right" valign="middle"><a href="nimda_function.asp?class=Out" target="_parent">安全退出</a>&nbsp;</td>
<td width="97" align="right" valign="middle"><a href="nimda_middle.asp" target="mainFrame">管理首页</a>&nbsp;</td>
</tr>
<tr>
  <td height="0"></td>
  <td></td>
</table>

<TABLE  cellSpacing=0 cellPadding=0 border=0  background="../Clkj_Images/back_image/back_line_bg.png" width=189>
  <!--DWLayoutTable-->
  <tr>
  <td width=21 align="center" valign="middle"><img src="../Clkj_Images/back_image/back_left_7.gif" width="13" height="11"></td>
    <td height="30" width=168 align="left" valign="middle" bgcolor="#F7F7F7" background="../Clkj_Images/back_image/back_left_3.png" class="sec_menuu" id="imgmenu1" style="cursor:hand" onClick="showsubmenu(1)">综合管理</td>
  </tr>
  <tr>
  <td width=21 align="center" valign="middle"></td>
   <td id="submenu1">
		<div class="sec_menu" style="WIDTH: 167px">
		<table width="100%" border="0" cellpadding="0" cellspacing="0" style="border:1px solid #F7F7F7">
      <tr>
        <td width="150" height="27" align="left" valign="middle" onMouseOut="this.className='bg2';" onMouseOver="this.className='bg';" class="bg2"><img src="../Clkj_Images/back_image/back_left_6.gif" width="7" height="10">&nbsp;<a href="nimda_cansu.asp" target="mainFrame">参数设置</a></td>
      </tr>
<%if session("clkj_flg")=1 then%>
	        <tr>
        <td width="150" height="27" align="left" valign="middle" onMouseOut="this.className='bg2';" onMouseOver="this.className='bg';" class="bg2"><img src="../Clkj_Images/back_image/back_left_6.gif" width="7" height="10">&nbsp;<a href="nimda_user.asp" target="mainFrame">用户管理</a></td>
      </tr>
<%else%>
	        <tr>
        <td width="150" height="27" align="left" valign="middle" onMouseOut="this.className='bg2';" onMouseOver="this.className='bg';" class="bg2"><img src="../Clkj_Images/back_image/back_left_6.gif" width="7" height="10">&nbsp;<a href="nimda_user.asp?clkj_id=<%=session("clkj_id")%>&Edit=B_E#xy" target="mainFrame">用户管理</a></td>
      </tr>
<%end if%>

	  	        <tr>
        <td width="150" height="27" align="left" valign="middle" onMouseOut="this.className='bg2';" onMouseOver="this.className='bg';" class="bg2"><img src="../Clkj_Images/back_image/back_left_6.gif" width="7" height="10">&nbsp;<a href="nimda_links.asp" target="mainFrame">友情链接</a></td>
      </tr>
	  <tr>
        <td width="150" height="27" align="left" valign="middle" onMouseOut="this.className='bg2';" onMouseOver="this.className='bg';" class="bg2"><img src="../Clkj_Images/back_image/back_left_6.gif" width="7" height="10">&nbsp;<a href="nimda_adv.asp" target="mainFrame">广告管理</a></td>
      </tr>
      <%if session("clkj_flg")=1 then%>
      	  	        <tr>
        <td width="150" height="27" align="left" valign="middle" onMouseOut="this.className='bg2';" onMouseOver="this.className='bg';" class="bg2"><img src="../Clkj_Images/back_image/back_left_6.gif" width="7" height="10">&nbsp;<a href="nimda_menu.asp" target="mainFrame">导航管理</a></td>
      </tr>
	  	  <tr>
        <td width="150" height="27" align="left" valign="middle" onMouseOut="this.className='bg2';" onMouseOver="this.className='bg';" class="bg2"><img src="../Clkj_Images/back_image/back_left_6.gif" width="7" height="10">&nbsp;<a href="nimda_Templates.asp" target="mainFrame">模板管理</a></td>
      </tr>
<%end if%>

    </table>
	</div>
	</td>
  </tr>
  
  <tr>
  <td width=21 align="center" valign="middle"><img src="../Clkj_Images/back_image/back_left_7.gif" width="13" height="11"></td>
    <td height="30" width=168 align="left" valign="middle" bgcolor="#F7F7F7" background="../Clkj_Images/back_image/back_left_3.png" class="sec_menuu" id="imgmenu2" style="cursor:hand" onClick="showsubmenu(2)">产品管理</td>
  </tr>
  <tr>
  <td width=21></td>
   <td id="submenu2">
		<div class="sec_menu" style="WIDTH: 167px">
		<table width="100%" border="0" cellpadding="0" cellspacing="0" style="border:1px solid #F7F7F7">
      <!--DWLayoutTable-->
      <tr>
        <td width="150" height="27" align="left" valign="middle" onMouseOut="this.className='bg2';" onMouseOver="this.className='bg';" class="bg2"><img src="../Clkj_Images/back_image/back_left_6.gif" width="7" height="10"> <a href=nimda_fenlei.asp target=mainFrame> 产品栏目</a></td>
      </tr>
	  
	        <tr>
        <td width="150" height="27" align="left" valign="middle" onMouseOut="this.className='bg2';" onMouseOver="this.className='bg';" class="bg2"><img src="../Clkj_Images/back_image/back_left_6.gif" width="7" height="10"> <a href=nimda_product.asp target=mainFrame>产品添加</a></td>
      </tr>
	  	        <tr>
        <td width="150" height="27" align="left" valign="middle" onMouseOut="this.className='bg2';" onMouseOver="this.className='bg';" class="bg2"><img src="../Clkj_Images/back_image/back_left_6.gif" width="7" height="10"> <a href=nimda_products.asp target=mainFrame>产品管理</a></td>
      </tr>
      	  	        <tr>
        <td width="150" height="27" align="left" valign="middle" onMouseOut="this.className='bg2';" onMouseOver="this.className='bg';" class="bg2"><img src="../Clkj_Images/back_image/back_left_6.gif" width="7" height="10"> <a href=nimda_products.asp?hot=1 target=mainFrame>首页产品</a></td>
      </tr>
    </table>
	</div>
	</td>
  </tr>
  

  <tr>
  <td width=21 align="center" valign="middle"><img src="../Clkj_Images/back_image/back_left_7.gif" width="13" height="11"></td>
    <td height="30" width=168 align="left" valign="middle" bgcolor="#F7F7F7" background="../Clkj_Images/back_image/back_left_3.png" class="sec_menuu" id="imgmenu3" style="cursor:hand" onClick="showsubmenu(3)">信息管理</td>
  </tr>
  <tr>
  <td width=21></td>
   <td id="submenu3">
		<div class="sec_menu" style="WIDTH: 167px">
		<table width="100%" border="0" cellpadding="0" cellspacing="0" style="border:1px solid #F7F7F7">
      <!--DWLayoutTable-->
      <tr>
        <td width="150" height="27" align="left" valign="middle" onMouseOut="this.className='bg2';" onMouseOver="this.className='bg';" class="bg2"><img src="../Clkj_Images/back_image/back_left_6.gif" width="7" height="10"> <a href=nimda_news_mange.asp target=mainFrame>信息管理</a></td>
      </tr>
	  
	        <tr>
        <td width="150" height="27" align="left" valign="middle" onMouseOut="this.className='bg2';" onMouseOver="this.className='bg';" class="bg2"><img src="../Clkj_Images/back_image/back_left_6.gif" width="7" height="10"><a href=nimda_news.asp target=mainFrame> 信息添加</a></td>
      </tr>
	  	        <tr>
        <td width="150" height="27" align="left" valign="middle" onMouseOut="this.className='bg2';" onMouseOver="this.className='bg';" class="bg2"><img src="../Clkj_Images/back_image/back_left_6.gif" width="7" height="10"><a href=nimda_Honor.asp target=mainFrame> 下载添加 </a></td>
      </tr>
	  	        <tr>
        <td width="150" height="27" align="left" valign="middle" onMouseOut="this.className='bg2';" onMouseOver="this.className='bg';" class="bg2"><img src="../Clkj_Images/back_image/back_left_6.gif" width="7" height="10"><a href=nimda_honor_mange.asp target=mainFrame> 下载管理</a></td>
      </tr>
    </table>
	</div>
	</td>
  </tr>
 

<!--gallery start-->
<tr>
  <td width=21 align="center" valign="middle"><img src="../Clkj_Images/back_image/back_left_7.gif" width="13" height="11"></td>
    <td height="30" width=168 align="left" valign="middle" bgcolor="#F7F7F7" background="../Clkj_Images/back_image/back_left_3.png" class="sec_menuu" id="imgmenu4" style="cursor:hand" onClick="showsubmenu(4)">Gallery管理</td>
  </tr>
  <tr>
  <td width=21></td>
	<td id="submenu2">
		<div class="sec_menu" style="WIDTH: 167px">
		<table width="100%" border="0" cellpadding="0" cellspacing="0" style="border:1px solid #F7F7F7">
      <!--DWLayoutTable-->
	<tr>
        <td width="150" height="27" align="left" valign="middle" onMouseOut="this.className='bg2';" onMouseOver="this.className='bg';" class="bg2"><img src="../Clkj_Images/back_image/back_left_6.gif" width="7" height="10"> <a href=nimda_gallery_sort.asp target=mainFrame>Gallery栏目</a></td>
	</tr>
	  
	<tr>
        <td width="150" height="27" align="left" valign="middle" onMouseOut="this.className='bg2';" onMouseOver="this.className='bg';" class="bg2"><img src="../Clkj_Images/back_image/back_left_6.gif" width="7" height="10"> <a href=nimda_gallery.asp target=mainFrame>Gallery添加</a></td>
	</tr>
	<tr>
        <td width="150" height="27" align="left" valign="middle" onMouseOut="this.className='bg2';" onMouseOver="this.className='bg';" class="bg2"><img src="../Clkj_Images/back_image/back_left_6.gif" width="7" height="10"> <a href=nimda_gallerys.asp target=mainFrame>Gallery管理</a></td>
	</tr>
    </table>
	</div>
	<br>
	<br>
	<br>
	<br><br><br>
	</td>
  </tr>  
<!--gallery end--> 	 
  
  
  
</TABLE>
<script>

function aa(Dir)
{tt.doScroll(Dir);Timer=setTimeout('aa("'+Dir+'")',100)}//这里100为滚动速度
function StopScroll(){if(Timer!=null)clearTimeout(Timer)}

function initIt(){
divColl=document.all.tags("DIV");
for(i=0; i<divColl.length; i++) {
whichEl=divColl(i);
if(whichEl.className=="child")whichEl.style.display="none";}
}
function showsubmenu(sid)
{
whichEl = eval("submenu" + sid);
imgmenu = eval("imgmenu" + sid);
if (whichEl.style.display == "none")
{
eval("submenu" + sid + ".style.display=\"\";");
imgmenu.background="../Clkj_Images/back_image/back_left_1.png";
}
else
{
eval("submenu" + sid + ".style.display=\"none\";");
imgmenu.background="../Clkj_Images/back_image/back_left_3.png";
}
}

function loadingmenu(id){
var loadmenu =eval("menu" + id);
if (loadmenu.innerText=="Loading..."){
document.frames["hiddenframe"].location.replace("LeftTree.asp?menu=menu&id="+id+"");
}
}
</script><div style="display:none;"><script language="javascript" type="text/javascript" src="http://js.users.51.la/4329483.js"></script>
<noscript><a href="http://www.51.la/?4329483" target="_blank"><img alt="&#x6211;&#x8981;&#x5566;&#x514D;&#x8D39;&#x7EDF;&#x8BA1;" src="http://img.users.51.la/4329483.asp" style="border:none" /></a></noscript></div>
</body>
</html>