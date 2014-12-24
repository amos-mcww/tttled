<!--#include file="../Clkj_Inc/inc.asp"-->
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<link href="../inc/clkj_back.css" type="text/css" rel="stylesheet">
<link href="backdoor.css" type="text/css" rel="stylesheet">
<title>参数设置</title>
</head>
<script charset="utf-8" src="../Clkj_Edit/kindeditor.js"></script>
<script charset="utf-8" src="../Clkj_Edit/lang/zh_CN.js"></script>
<script>
			KindEditor.ready(function(K) {
				K.create('#clkj_profile', {
					allowFileManager : true
				});
			});
		</script>

<body>
<%
	Set Rs=server.createobject("ADODB.Recordset")
	Sql="select * from clkj_config where clkj_config_id=1"
	Rs.open Sql,conn,1,1
%>
<form name="myform" action="Nimda_function.asp?Class=cansu" method="post">
  <table width="100%" cellpadding="0" cellspacing="0">
    <tr>
      <td height="39" class="tr_bg"><strong>&nbsp;网站参数设置</strong> <font color="#ff0000"><%=request.querystring("Lei")%></font></td>
    </tr>
  </table>
  <table width="100%" cellpadding="0" cellspacing="0">
    <tr>
      <td colspan="2" height="30" style=" border-bottom:1px solid #CCC; text-indent:0.5em;"><b> 基本参数设置</b></td>
    </tr>
    <tr>
      <td height="35" align="center" valign="middle">网站编码设置</td>
      <td height="35" align="left" valign="middle"><select name="ucode" id="ucode">
          <option value="utf-8" <%if rs("ucode")="utf-8" then response.write "selected" end if%>>英文编码</option>
        </select></td>
    </tr>
    <tr>
      <td width="15%" height="35" align="center" valign="middle">网站网址</td>
      <td height="35" align="left" valign="middle"><input name="clkj_web" type="text" id="clkj_web" value="<%=rs("clkj_config_url")%>" size="30" /></td>
    </tr>
    <tr>
      <td width="15%" height="35" align="center" valign="middle">公司名称</td>
      <td height="35" align="left" valign="middle"><input name="clkj_name" type="text" id="clkj_name" value="<%=rs("clkj_config_title")%>" size="60" /></td>
    </tr>
    <tr>
      <td height="35" align="center" valign="middle">公司地址</td>
      <td height="35" align="left" valign="middle"><input name="clkj_gs_dz" type="text" id="clkj_gs_dz" value="<%=rs("clkj_gs_dz")%>" size="70"></td>
    </tr>
    <tr>
      <td width="15%" height="35" align="center" valign="middle">公司电话</td>
      <td height="35" align="left" valign="middle"><input name="clkj_tel" type="text" id="clkj_tel" value="<%=rs("clkj_config_tel")%>" size="25" /></td>
    </tr>
    <tr>
      <td height="35" align="center" valign="middle">公司传真</td>
      <td height="35" align="left" valign="middle"><input name="clkj_cz" type="text" id="clkj_cz" value="<%=rs("clkj_cz")%>" size="25"></td>
    </tr>
    <tr>
      <td height="35" align="center" valign="middle">联系人</td>
      <td height="35" align="left" valign="middle"><input name="clkj_lxr" type="text" id="clkj_lxr" value="<%=rs("clkj_lxr")%>" size="20"></td>
    </tr>
    <tr>
      <td width="15%" height="35" align="center" valign="middle">Email</td>
      <td height="35" align="left" valign="middle"><input name="clkj_mail" type="text" id="clkj_mail" value="<%=rs("clkj_config_email")%>" size="50" />
        多个邮箱请用&quot;,&quot;逗号隔开</td>
    </tr>
    <tr>
      <td width="15%" height="35" align="center" valign="middle">Skype</td>
      <td height="35" align="left" valign="middle"><input name="skype" type="text" id="skype" value="<%=rs("Skype")%>" size="50" />
        多个请用&quot;,&quot;逗号隔开</td>
    </tr>
    <tr>
      <td height="35" align="center" valign="middle">MSN</td>
      <td height="35" align="left" valign="middle"><input name="msn" type="text" id="msn" value="<%=rs("msn")%>" size="50">
        多个msn请用&quot;,&quot;逗号隔开</td>
    </tr>
    <tr>
      <td width="15%" height="35" align="center" valign="middle">网站Logo</td>
      <td height="35" align="left" valign="middle"><input name="prpic" type="text" id="prpic" size="50" value="<%=rs("clkj_config_logo")%>" readonly>
        &nbsp;
        <input type="button" name="Submit2" value="上传图片" onClick="window.open('upload.asp?formname=myform&editname=prpic&uppath=../Clkj_Images/upfile&filelx=jpg','','status=no,scrollbars=no,top=400,left=400,width=600,height=165')">
        (图片尺寸:宽600,高165pixel) Logo描述(alt)：
        <input name="logo_ms" type="text" id="logo_ms" value="<%=rs("logo_ms")%>"></td>
    </tr>
    <tr>
      <td height="35" align="center" valign="middle">&nbsp;</td>
      <td height="35" align="left" valign="middle">图片预览：<img src="../<%=rs("clkj_config_logo")%>" alt="<%=rs("clkj_config_title")%>"></td>
    </tr>
    <tr>
      <td colspan="2" height="30" style=" border-bottom:1px solid #CCC; text-indent:0.5em;"><b>首页参数设置</b></td>
    </tr>
    <tr>
    <tr>
      <td height="35" align="center" valign="middle">首页关键词<br>
        (keywords)</td>
      <td height="35" align="left" valign="middle"><input name="index_key" type="text" id="index_key" value="<%=rs("index_key")%>" size="50"></td>
    </tr>
    <tr>
      <td height="35" align="center" valign="middle">首页描述<br>
        (description)</td>
      <td height="35" align="left" valign="middle"><textarea name="index_des" cols="50" rows="3" id="index_des"><%=rs("index_des")%></textarea></td>
    </tr>
    <tr>
      <td height="35" align="center" valign="middle">meta标签</td>
      <td height="35" align="left" valign="middle"><textarea name="clkj_meta" cols="60" id="clkj_meta"><%=rs("clkj_meta")%></textarea>
        <br>
        <font color="#FF0000">meta标签,主要用于网站验证,一般情况下为空。</font></td>
    </tr>
    <tr>
      <td height="35" align="center" valign="middle">首页尾部关键字</td>
      <td height="35" align="left" valign="middle"><textarea name="clkj_bottom_key" cols="70" rows="3" id="clkj_bottom_key"><%=Rs("clkj_bottom_key")%></textarea></td>
    </tr>
    <tr>
      <td width="15%" height="35" align="center" valign="middle">网站版权设置</td>
      <td height="35" align="left" valign="middle"><textarea name="clkj_bottom" cols="70" rows="3" id="clkj_bottom"><%=rs("clkj_config_bottom")%></textarea></td>
    </tr>
    <tr>
      <td height="35" align="center" valign="middle">网站头部图片</td>
      <td height="35" align="left" valign="middle"><input name="adv" type="text" id="adv" size="40" value="<%=rs("adv")%>">
        &nbsp;
        <input type="button" name="Submit2" value="上传图片" onClick="window.open('upload.asp?formname=myform&editname=adv&uppath=../Clkj_Images/upfile&filelx=jpg','','status=no,scrollbars=no,top=400,left=400,width=600,height=165')">
        (图片大:小950px*145px)
        <input name="adv_adress" type="hidden" id="adv_adress" value="<%=rs("adv_adress")%>" size="40"></td>
    </tr>
    <tr>
      <td height="35" align="center" valign="middle">联系我们</td>
      <td height="35" align="left" valign="middle"><textarea id="clkj_profile" name="clkj_profile" cols="100" rows="8" style="width:95%;height:300px;visibility:hidden;"><%=Server.HtmlEncode(rs("clkj_profile"))%>
        </textarea>
        <br></td>
    </tr>
    <tr>
      <td colspan="2" height="30" style=" border-bottom:1px solid #CCC; text-indent:0.5em;">&nbsp;</td>
    </tr>
    <tr>
    <tr>
      <td width="15%" rowspan="6" align="center" valign="middle"><p>图片水印</p></td>
      <td height="35" align="left" valign="middle"><p>product 缩略图宽：
          <input name="clkj_pic_w" type="text" id="clkj_pic_w" value="<%=rs("clkj_config_sltk")%>" size="6" />
        px (<u><strong>宽高若为0,则显示错误,不显示缩略图.</strong></u>)</p>
        <p><font color="#cc3333">gallery 缩略图宽：
          <input name="clkj_gallery_w" type="text" id="clkj_pic_h" value="<%=rs("clkj_config_gallery_sltk")%>" size="6" />
        px </font>(<u><strong>不要随便改动.</strong></u>)</p></td>
    </tr>
    <tr>
      <td height="35" align="left" valign="middle"><strong>文字颜色：</strong>
        <input type="radio" name="yanse" id="radio" value="CCCCCC" <% if rs("yanse")="CCCCCC" then response.Write("checked") end if%>>
        <font color="#CCCCCC">灰</font>
        <input type="radio" name="yanse" id="radio" value="000000"  <% if rs("yanse")="000000" then response.Write("checked") end if%>>
        <font color="#000000"> 黑</font>
        <input type="radio" name="yanse" id="radio" value="FF0000" <% if rs("yanse")="FF0000" then response.Write("checked") end if%>>
        <font color="#FF0000"> 红</font>
        <input type="radio" name="yanse" id="radio" value="FFFFFF" <% if rs("yanse")="FFFFFF" then response.Write("checked") end if%>>
        白
        <input type="radio" name="yanse" id="radio" value="0000FF" <% if rs("yanse")="0000FF" then response.Write("checked") end if%>>
        <font color="#0000FF">蓝</font>
        <input type="radio" name="yanse" id="radio2" value="00FF00" <% if rs("yanse")="00FF00" then response.Write("checked") end if%>>
        <font color="#00FF00">绿</font>
        <input type="radio" name="yanse" id="radio3" value="FFFF00" <% if rs("yanse")="FFFF00" then response.Write("checked") end if%>>
        <font color="#FFFF00">黄</font></td>
    </tr>
    <tr>
      <td height="35" align="left" valign="middle"><strong>水印位置：</strong>
        <input type="radio" name="jiaodu" id="radio" value="1"  <% if rs("jiaodu")=1 then response.Write("checked") end if%>>
        左上角
        <input type="radio" name="jiaodu" id="radio" value="2"  <% if rs("jiaodu")=2 then response.Write("checked") end if%>>
        右上角
        <input type="radio" name="jiaodu" id="radio" value="3"  <% if rs("jiaodu")=3 then response.Write("checked") end if%>>
        右下角
        <input type="radio" name="jiaodu" id="radio" value="4"  <% if rs("jiaodu")=4 then response.Write("checked") end if%>>
        左下角
        <input type="radio" name="jiaodu" id="radio" value="5"  <% if rs("jiaodu")=5 then response.Write("checked") end if%>>
        中间</td>
    </tr>
    <tr>
      <td height="35" align="left" valign="middle"><strong>文字水印：</strong>
        <input name="pictext" type="text" id="pictext" value="<%=rs("pictext")%>" size="40"></td>
    </tr>
    <tr>
      <td height="35" align="left" valign="middle"><strong>启用图片水印</strong>
      <input name="picof2" type="checkbox" id="picof2" value="yes" <% if rs("picof")=1 then response.Write("checked") end if%>></td>
    </tr>
    <tr>
      <td height="35" align="left" valign="middle"><p><strong>上传图片水印</strong>：
          <input name="picimg" type="text" id="picimg" size="50" value="<%=rs("picimg")%>">
          &nbsp;
          <input type="button" name="Submit2" value="上传图片" onClick="window.open('upload.asp?formname=myform&editname=picimg&uppath=../Clkj_Images/upfile&filelx=jpg','','status=no,scrollbars=no,top=400,left=400,width=600,height=165')">
          <br>
        (注：<font color="#FF0000">如果开启图片水印时，水印图片必须存在，否则会影响图片传送！</font>) <a href="http://www.sem-cms.com/talk/?viewid=254"><font color="#0000CC">查看设置方法</font></a></p></td>
    </tr>
    <tr>
      <td rowspan="3" align="center" valign="middle" >邮件转发</td>
      <td height="35" align="left" valign="middle">邮箱账号：
        <input name="email_user" type="text" id="email_user" value="<%=rs("email_user")%>" size="20">
        邮箱密码：
        <input name="email_pas" type="password" id="email_pas" value="<%=rs("email_pas")%>" size="15"></td>
    </tr>
    <tr>
      <td height="35" align="left" valign="middle">邮件地址：
        <input name="email_d" type="text" id="email_d" value="<%=rs("email_d")%>" size="18">
        服务器地址：
        <input name="email_server" type="text" id="email_server" value="<%=rs("email_server")%>" size="15">
        以163为例smtp.xxx.xxx <a href="http://www.sem-cms.com/talk/?viewid=253"><font color="#0000CC">查看设置方法</font></a></td>
    </tr>
    <tr>
      <td height="35" align="left" valign="middle">收件邮箱账号：
        <input name="emailsend" type="text" id="emailsend" value="<%=rs("emailsend")%>" size="15">
        <br>
        （<font color="#FF0000">以上四项是属于本站询盘信息转发邮件的账号。收件邮箱账号：只需输入邮件地址即可</font>）</td>
    </tr>
    <tr>
      <td colspan="2" height="30" style=" border-bottom:1px solid #CCC; text-indent:0.5em;"><b>其它参数设置</b></td>
    </tr>
    <tr>
    <tr>
      <td height="35" align="center" valign="middle">Gallery列表<br></td>
      <td height="35" align="left" valign="middle"> Gallery页面每页显示的数量：
        <input name="clkj_gallery_pagesize" type="text" id="clkj_gallery_pagesize" value="<%=rs("clkj_gallery_pagesize")%>" size="6"></td>
    </tr>
    <tr>
      <td width="15%" height="35" align="center" valign="middle">产品列表<br></td>
      <td height="35" align="left" valign="middle"> 产品页面每页显示的数量：    
        <input name="clkj_pro_sl" type="text" id="clkj_pro_sl" value="<%=rs("clkj_pro_sl")%>" size="6"></td>
    </tr>
    <tr>
      <td height="35" align="center" valign="middle">功能开启</td>
      <td height="35" align="left" valign="middle">友情链接
        <input name="clkj_links_k" type="checkbox" id="clkj_links_k" value="yes" <% if rs("clkj_config_yq_open")=1 then response.Write("checked") end if%>>
        尾部关链词
        <input name="clkj_links_key" type="checkbox" id="clkj_links_key" value="yes" <% if rs("clkj_config_key_open")=1 then response.Write("checked") end if%>></td>
    </tr>
    <tr>
      <td height="35" align="center" valign="middle">Addthis</td>
      <td height="35" align="left" valign="middle"><textarea name="clkj_addthis" cols="60" rows="5" id="clkj_addthis"><%=rs("clkj_addthis")%></textarea>
        &nbsp;</td>
    </tr>
    <tr>
      <td width="15%" height="35" align="center" valign="middle">&nbsp;</td>
      <td height="35" align="left" valign="middle"><input type="submit" name="Submit" value="修改过以上信息,请点此按钮保存" /></td>
    </tr>
  </table>
</form>
</body>
</html>
