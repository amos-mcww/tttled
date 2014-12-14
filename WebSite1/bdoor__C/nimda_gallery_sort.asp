<!--#include file="nimda_function.asp"-->
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>栏目管理</title>
<link href="backdoor.css" type="text/css" rel="stylesheet"></head>

<body>
<table width="100%" cellpadding="0" cellspacing="0">
  <tr>
    <td height="39" class="tr_bg">&nbsp;
    <p><strong>Gallery分类管理</strong> [<a
        href="Nimda_gallery_class.asp?Edit=gallery_big_add#zy"><font
        color="#ff6600">添加一级分类</font></a>]</p>
    <p>&nbsp;</p>
    <p><img src="/Clkj_Images/back_image/gallery_snagit.png" width="222" height="304"></p>
    <p>&nbsp;</p></td>
  </tr>
</table>
<table cellpadding="0" cellspacing="0" width="100%" border="1" bordercolordark="#FFFFFF" bordercolorlight="#CCCCCC">
<tr>
<td width="8%" align="center" valign="middle" bgcolor="#D9EFFE"><strong>序号</strong></td>
<td width="45%" height="30" align="center" valign="middle" bgcolor="#D9EFFE" style="padding:2px;"><strong>分类</strong></td>
<td  align="center" valign="middle" bgcolor="#D9EFFE" style="padding:2px;"><strong>操作</strong></td>
</tr>
<%call GalleryFenLei()%>
</table>

</body>
</html>
