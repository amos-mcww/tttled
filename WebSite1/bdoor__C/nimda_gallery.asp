<!--#include file="../Clkj_Inc/inc.asp"-->
<%
IF Request.QueryString("picdel")="picdel" Then
Call Delpic(cint(request("clkj_prid")),request.QueryString("purl"),"?clkj_prid="&request("clkj_prid")&"&Edit="&request.QueryString("Edit")&"&clkj_BigClassID="&request("clkj_BigClassID")&"&clkj_SmallClassID="&request("clkj_SmallClassID")&"&sf="&request.QueryString("sf")&"&ToPage="&request("Topage")&"#pi")
End IF
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<link href="backdoor.css" type="text/css" rel="stylesheet">
<title>Gallery添加</title>
		<script charset="utf-8" src="../Clkj_Edit/kindeditor.js"></script>
        <script charset="utf-8" src="../Clkj_Edit/lang/zh_CN.js"></script>
		<script>
			KindEditor.ready(function(K) {
				K.create('#clkj_prcontent', {
					allowFileManager : true
				});
			});
		</script>
</head>


<script type="text/JavaScript">
<!--
function MM_jumpMenu(targ,selObj,restore){ //v3.0
  eval("location='"+selObj.options[selObj.selectedIndex].value+"'");
  if (restore) selObj.selectedIndex=0;
}
//-->
</script>
<script language="javascript"> 
<!-- 
var subcat = new Array(); 
<%
Dim subcat,rsClass2,sqlClass2
set rsClass2=server.createobject("adodb.recordset")
sqlClass2="select * from clkj_SmallClass"
rsClass2.open sqlClass2,conn,1,1
%>
var subval2 = new Array();
<%
count2 = 0
do while not rsClass2.eof
%>
subcat[<%=count2%>] = new Array('<%=rsClass2("clkj_BigClassID")%>','<%=rsClass2("clkj_SmallClassName")%>','<%=rsClass2("clkj_SmallClassID")%>')
<%
count2 = count2 + 1
rsClass2.movenext
loop
rsClass2.close
%>
function changeselect1(locationid) 
{ 
document.products.small_Lei.length = 0;          
document.products.small_Lei.options[0] = new Option('请选择','');    
for (i=0; i <subcat.length; i++) 
{ 
if (subcat[i][0] == locationid)                  
{document.products.small_Lei.options[document.products.small_Lei.length] = new Option(subcat[i][1], subcat[i][2]);}  
} 
} 
//--> 
</script>
<body>
<table width="100%" cellpadding="0" cellspacing="0">
  <tr>
    <td height="39" class="tr_bg">&nbsp;<strong>Gallery管理</strong> <font color="#ff0000"><%=request.querystring("Lei")%></font> </td>
  </tr>
</table>
<%
IF Request.querystring("Edit")="gallery_edit" Then 
%>
<form name="small_class" action="Nimda_function.asp?Class=Pro_Edit&clkj_BigClassID=<%=request("clkj_BigClassID")%>" method="post">
<table width="100%" cellpdding="0" cellspacing="0">

<tr>
<td height="30" colspan="2" align="left" valign="middle" bgcolor="#F5F5F5" style="text-indent:2em;color:#ff0000;" id="zr"><strong>Gallery</strong>修改</td>
</tr>
<tr>
<td width="15%" height="30" align="center" valign="middle">类别选择</td>
<td align="left" valign="middle"><p><span class="h_td">
  <select name="big_Lei" onChange="MM_jumpMenu('_slef',this,0)">
    <%
	Set Rs=server.createobject("ADODB.Recordset")
	Sql="select * from clkj_gallery_BigClass"
	Rs.open Sql,conn,1,1
	Do while not Rs.Eof
	%>
    <option value="?clkj_BigClassID=<%=rs("clkj_BigClassID")%>&clkj_prid=<%=request("clkj_prid")%>&Edit=P_E" <%IF cint(request("clkj_BigClassID"))=rs("clkj_BigClassID") then%>selected<%end if%>><%=rs("clkj_BigClassName")%></option>
    <%
	Rs.Movenext
	Loop	
	%>
  </select>
  <select name="small_Lei">
    <option>暂无栏目</option>
    <%
	Set Rs=server.createobject("ADODB.Recordset")
	Sql="select * from clkj_gallery_SmallClass where clkj_BigClassID="&request("clkj_BigClassID")
	Rs.open Sql,conn,1,1
	IF Rs.bof or Rs.eof then
	response.write "<option>暂无栏目</option>"
	Else
	Do while not Rs.Eof
	%>
    <option value="<%=rs("clkj_SmallClassID")%>" <%IF cint(request("clkj_SmallClassID"))=rs("clkj_SmallClassID") then%>selected<%end if%>><%=rs("clkj_SmallClassName")%></option>
    <%
	Rs.Movenext
	Loop
	End IF	
	%>
  </select>
</span></p>
  <p><span class="h_td"><font color="#ff000">*</font></span>一级类别必须选择 </p></td>
</tr>
<%
	Set Rs=server.createobject("ADODB.Recordset")
	Sql="select * from clkj_Products where clkj_prid ="&cint(request("clkj_prid"))
	Rs.open Sql,conn,1,1
%>
<tr>
<td width="15%" height="30" align="center" valign="middle">产品名称</td>
<td><input name="clkj_prtitle" type="text" id="clkj_prtitle" value="<%=rs("clkj_prtitle")%>" size="50">
  <font color="#ff000">*</font>标题</td>
</tr>
<tr>
  <td align="center" valign="middle" height="30">产品排序</td>
  <td>
    <input name="clkj_paixu" type="text" id="textfield" value="<%=rs("clkj_paixu")%>" size="5">
    <font color="#ff000">*</font>从小到大排序,0,1,2,3,4,5.相同数字按id排序.默认下为'10000'</td>
</tr>
<tr>
<td width="15%" height="30" align="center" valign="middle">产品关键字<br>(keywords)</td>
<td><input name="clkj_prkey" type="text" id="textfield2" value="<%=rs("clkj_prkey")%>">
  关键字2-5个</td>
</tr>
<tr>
<td width="15%" height="30" align="center" valign="middle">产品页面描述<br>(description)</td>
<td><textarea name="clkj_prprdes" cols="50" id="clkj_prprdes"><%=rs("clkj_prprdes")%></textarea>
  描述100个字符以内</td>
</tr>



<tr>
  <td height="30" align="center" valign="middle">产品详细描述</td>
  <td><textarea id="clkj_prcontent" name="clkj_prcontent" cols="100" rows="8" style="width:95%;height:300px;visibility:hidden;"><%=Server.HtmlEncode(rs("clkj_prcontent"))%></textarea>
  </td>
</tr>
<%
Arayypic=rs("clkj_prpic")
pis=1
Arayypic=split(Arayypic,",")
for each astrss in Arayypic

if astrss<>"" and astrss<>" " then'空字符不执行
%>
<tr>
  <td align="center" valign="middle" height="30">上传产品图片</td>
    <td>
    <input name="clkj_prpic" type="text" id="clkj_prpic_<%=pis%>" value="<%=trim(astrss)%>" size="50" readonly>
    <input type="button" name="Submit2" value="修改图片" onClick="window.open('upload.asp?formname=small_class&editname=clkj_prpic_<%=pis%>&uppath=../Clkj_Images/upfile&filelx=jpg','','status=no,scrollbars=no,top=400,left=400,width=600,height=165')"><a href="?clkj_prid=<%=request("clkj_prid")%>&Edit=<%=request.QueryString("Edit")%>&clkj_BigClassID=<%=request("clkj_BigClassID")%>&clkj_SmallClassID=<%=request("clkj_SmallClassID")%>&picdel=picdel&purl=<%=trim(astrss)%>&sf=<%=request.QueryString("sf")%>&ToPage=<%=request("Topage")%>">删除</a>
    <font color="#ff000">*</font>图片控制于200k内</td>
</tr>
<%
end if
pis=pis+1
next
%>
<%IF request.QueryString("picadd")<>"add" Then%> 
<tr>
 <td align="center" valign="middle" height="30"></td> <td align="center"><a href="?clkj_prid=<%=request("clkj_prid")%>&Edit=<%=request.QueryString("Edit")%>&clkj_BigClassID=<%=request("clkj_BigClassID")%>&clkj_SmallClassID=<%=request("clkj_SmallClassID")%>&sf=<%=request.QueryString("sf")%>&ToPage=<%=request("Topage")%>&picadd=add#pic"><strong><font color="#0000FF">增加图片</font></strong></a></td>
</tr>
<%else%>
<tr>
 <td align="center" valign="middle" height="30" ></td> <td align="center"><a href="?clkj_prid=<%=request("clkj_prid")%>&Edit=<%=request.QueryString("Edit")%>&clkj_BigClassID=<%=request("clkj_BigClassID")%>&clkj_SmallClassID=<%=request("clkj_SmallClassID")%>&sf=<%=request.QueryString("sf")%>&ToPage=<%=request("Topage")%>&picadd=edd#pi"><strong><font color="#FF0000">隐藏上传图片</font></strong></a></td>
</tr>
<%end if%>
   <%IF request.QueryString("picadd")="add" Then%>
    <tr>
    <td rowspan="4" align="center" valign="middle" id="pic">上传产品图片</td>
    <td height="30"><input name="clkj_prpic" type="text" id="clkj_prpic_o" size="50"> <input type="button" name="Submit2" value="上传图片" onClick="window.open('upload.asp?formname=small_class&editname=clkj_prpic_o&uppath=../Clkj_Images/upfile&filelx=jpg','','status=no,scrollbars=no,top=400,left=400,width=600,height=165')">
      <font color="#ff000">*</font>图片控制于200k内</td>
  </tr> 
    <tr>
    <td height="30"><input name="clkj_prpic" type="text" id="clkj_prpic_t" size="50"> <input type="button" name="Submit2" value="上传图片" onClick="window.open('upload.asp?formname=small_class&editname=clkj_prpic_t&uppath=../Clkj_Images/upfile&filelx=jpg','','status=no,scrollbars=no,top=400,left=400,width=600,height=165')"></td>
  </tr>
      <tr>
    <td height="30"><input name="clkj_prpic" type="text" id="clkj_prpic_s" size="50"> <input type="button" name="Submit2" value="上传图片" onClick="window.open('upload.asp?formname=small_class&editname=clkj_prpic_s&uppath=../Clkj_Images/upfile&filelx=jpg','','status=no,scrollbars=no,top=400,left=400,width=600,height=165')"></td>
  </tr> 
      <tr>
    <td height="30"><input name="clkj_prpic" type="text" id="clkj_prpic_f" size="50"> <input type="button" name="Submit2" value="上传图片" onClick="window.open('upload.asp?formname=small_class&editname=clkj_prpic_f&uppath=../Clkj_Images/upfile&filelx=jpg','','status=no,scrollbars=no,top=400,left=400,width=600,height=165')"></td>
  </tr> 
<%
End IF
%>
  
  <tr>
    <td align="center" valign="middle" height="30" id="pi">首页推荐产品</td>
    <td>
      推荐
      <input name="clkj_hot" type="checkbox" id="clkj_hot" value="yes" <% if rs("clkj_hot")=1 then response.Write("checked") end if%>>
      <font color="#ff000">*</font>首页显示 </td>
  </tr>
  <tr>
    <td align="center" valign="middle" height="30">产品文件名</td>
    <td><input name="clkj_pr_url" type="text" id="clkj_pr_url" value="<%=rs("clkj_pr_url")%>">
      <font color="#ff000">*默认为为空；生成静态时用的，必须唯一(文件名如:RF-Hotel-Lock)单词之前用&quot;-&quot;链接</font></td>
  </tr>
  <tr>
  <td align="center" valign="middle" height="30">录入员</td>
  <td><input name="c_ru" type="text" id="c_ru" value="<%=session("username")%>" size="15" >
  录入时间<input name="clkj_time" type="text" id="clkj_time" value="<%=now()%>" size="20" readonly></td></tr>
<tr>
<td width="15%" height="30" align="center" valign="middle">&nbsp;</td>
<td width="85%" align="left" class="h_td"><input type="submit" name="Submit" value="修改" /><input type="hidden" name="clkj_prid" value="<%=request("clkj_prid")%>"/><input type="hidden" name="ToPage" value="<%if request.QueryString("ToPage")="" then response.write "1"%><%if request.QueryString("ToPage")<>"" then response.write request.QueryString("ToPage")%>"/><input type="hidden" name="sf" value="<%=request.QueryString("sf")%>"/><input type="checkbox" name="lscp" id="checkbox" value="yes"><input type="submit" name="Submit" value="发布类似产品" /><font color="red">在点"发布类似产品"按钮时，前面的沟必须钩上！</td>
</tr>

</table></form>

<%Else%>
<form name="products" action="Nimda_function.asp?Class=gallery_add" method="post">
<table width="100%" cellpadding="0" cellspacing="0">

<tr>
<td height="30" colspan="2" align="left" valign="middle" bgcolor="#F5F5F5" style="text-indent:2em;color:#ff0000;" id="zr"><p> Gallery增加     [<a href="Nimda_gallery_class.asp?Edit=gallery_big_add#zy"><font color="#3333CC">点这里增加一级分类</font></a>]</p></td>
</tr>
<tr>
<td width="15%" align="center" valign="middle" height="30">类别选择</td><td>
<%
Dim count1,rsClass1,sqlClass1
set rsClass1=server.createobject("adodb.recordset")
sqlClass1="select * from clkj_gallery_BigClass" 
rsClass1.open sqlClass1,conn,1,1
%>
<select name="big_Lei" id="big_Lei" onChange="changeselect1(this.value)">
<option value="">请选择一级分类</option>
<%
count1 = 0
do while not rsClass1.eof
response.write"<option value="&rsClass1("clkj_BigClassID")&">"&rsClass1("clkj_BigClassName")&"</option>"
count1 = count1 + 1
rsClass1.movenext
loop
rsClass1.close
%>
</select>
<select name="small_Lei" id="small_Lei"> 
<option value="" selected="selected">请选择二级分类</option>
</select>
<font color="#ff000">*</font>必选项
</td>
</tr>
<tr>
  <td align="center" valign="middle" height="30">Gallery名称</td>
  <td><input name="clkj_prtitle" type="text" id="clkj_prtitle" size="70">
    <font color="#ff000">*</font>标题 </td>
</tr>
<tr>
  <td align="center" valign="middle" height="30">产品排序</td>
  <td>
    <input name="clkj_paixu" type="text" id="textfield" value="1" size="6">
    <font color="#ff000">*</font>多个gallery项目中,数字越大该类越靠前</td>
</tr>
<tr>
  <td align="center" valign="middle" height="30">产品关键字<br>(keywords)</td>
  <td>
    <input name="clkj_prkey" type="text" id="textfield" size="70">
    <font color="#ff000">*</font>空格隔开</td>
</tr>
<tr>
  <td align="center" valign="middle" height="45"><p>产品页面描述<br>(description)</p>
    <p>&nbsp;</p></td>
  <td><p>
    <textarea name="clkj_prprdes" cols="75" rows="5" id="clkj_prprdes"></textarea>
    <font color="#ff000">*</font>建议500个字符以内</p></td>
</tr>
<!--<tr>
  <td align="center" valign="middle" height="45">产品页面顶部信息</td>
  <td><textarea name="clkj_db" cols="50" rows="5" id="clkj_db"></textarea></td>
</tr>-->
<tr>
  <td width="15%" align="center" valign="middle">产品详细描述</td>
  <td>
<textarea id="clkj_prcontent" name="clkj_prcontent" cols="100" rows="15" style="width:96%;height:400px;visibility:hidden;"></textarea></td>
  </tr>
  <tr>
    <td rowspan="4" align="center" valign="middle"><p>上传产品图片</p>
      <p>(第一张为标题图)</p></td>
    <td height="30"><p>
      <input name="clkj_prpic" type="text" id="clkj_prpic_1" size="70"> 
      <input type="button" name="Submit2" value="上传图片" onClick="window.open('upload.asp?formname=products&editname=clkj_prpic_1&uppath=../Clkj_Images/upfile&filelx=jpg','','status=no,scrollbars=no,top=400,left=400,width=600,height=165')">
    </p></td>
  </tr> 
    <tr>
    <td height="30"><input name="clkj_prpic" type="text" id="clkj_prpic_2" size="70"> <input type="button" name="Submit2" value="上传图片" onClick="window.open('upload.asp?formname=products&editname=clkj_prpic_2&uppath=../Clkj_Images/upfile&filelx=jpg','','status=no,scrollbars=no,top=400,left=400,width=600,height=165')"></td>
  </tr>
      <tr>
    <td height="30"><input name="clkj_prpic" type="text" id="clkj_prpic_3" size="70"> <input type="button" name="Submit2" value="上传图片" onClick="window.open('upload.asp?formname=products&editname=clkj_prpic_3&uppath=../Clkj_Images/upfile&filelx=jpg','','status=no,scrollbars=no,top=400,left=400,width=600,height=165')"></td>
  </tr> 
      <tr>
        <td height="30"><input name="clkj_prpic" type="text" id="clkj_prpic_4" size="70"> <input type="button" name="Submit2" value="上传图片" onClick="window.open('upload.asp?formname=products&editname=clkj_prpic_4&uppath=../Clkj_Images/upfile&filelx=jpg','','status=no,scrollbars=no,top=400,left=400,width=600,height=165')"></td>
    </tr>  
  <tr>
    <td align="center" valign="middle" height="30">产品文件名</td>
    <td><p>
      <input name="clkj_pr_url" type="text" id="clkj_pr_url">
      </p>
      <p><font color="#ff000">生成静态时用的，必须唯一(文件名如:RF-Hotel-Lock)单词之前用&quot;-&quot;链接</font></p></td>
  </tr>
  <tr>
  <td align="center" valign="middle" height="30">录入员</td>
  <td><input name="c_ru" type="text" id="c_ru" value="<%=session("username")%>" size="30" >
  <!--录入时间-->
    <input name="clkj_time" type="hidden" id="clkj_time" value="<%=now()%>" size="20" readonly></td>
</tr>
  <tr>
    <td align="center" valign="middle" height="30">&nbsp;</td>
    <td><input type="submit" name="Submit" value="添加产品" ></td>
  </tr>

</table>
  </form>
<%End IF%>

</body>
</html>
