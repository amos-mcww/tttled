<!--#include file="../Clkj_Inc/inc.asp"-->
<%
uppath=request("uppath")&"/"			
filelx=request("filelx")				
formName=request("formName")			
EditName=request("EditName")		
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<link href="backdoor.css" type="text/css" rel="stylesheet">
<script language="javascript">
<!--
function mysub()
{
		esave.style.visibility="visible";
}
-->
</script>
</head>
<body>
<form name="form1" method="post" action="upfile.asp" enctype="multipart/form-data" >
  <div id="esave" style="position:absolute; top:18px; left:40px; z-index:10; visibility:hidden">
    <TABLE WIDTH=340 BORDER=0 CELLSPACING=0 CELLPADDING=0>
      <TR>
        <td width=20%></td>
        <TD bgcolor=#ff0000 width="60%"><TABLE WIDTH=100% height=120 BORDER=0 CELLSPACING=1 CELLPADDING=0>
            <TR>
              <td bgcolor=#ffffff align=center><font color=red>正在上传文件，请稍候...</font></td>
            </tr>
          </table></td>
        <td width=20%></td>
      </tr>
    </table>
  </div>
  <table width="95%" border="0" align="center" cellspacing="1" bgcolor="#FFFFFF">
    <tr>
      <td align="center" height="50">
	  <strong>图片上传</strong>
        <input type="hidden" name="filepath" value="<%=uppath%>">
        <input type="hidden" name="filelx" value="<%=filelx%>">
        <input type="hidden" name="EditName" value="<%=EditName%>">
        <input type="hidden" name="FormName" value="<%=formName%>">
        <input type="hidden" name="act" value="uploadfile">
         </td>
    </tr>
    <tr >
      <td align="center" id="upid" height="30">自定义图片名称
        <input name="imgname" type="text" id="imgname" size="20" class="tx1" >
        <font color="#FF0000">注意：此项为空即,随机生成图片文件名</font> </td>
    </tr>
    <tr >
      <td align="center" id="upid" height="50">选择文件:
        <input type="file" name="file1" size="45" class="tx1" value="">
        <input type="submit" name="Submit" value="开始上传" onClick="javascript:mysub()">
      </td>
    </tr>
  </table>
</form>
</body>
</html>
