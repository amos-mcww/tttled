<!--#include file="inc/plus.inc.asp"-->
<% 
'**************************************************************
' Software name: SmartStat
' Web: http://www.cactussoft.cn
' Copyright (C) 2008－2010 仙人掌软件 版权所有
'**************************************************************
if Request.ServerVariables("REQUEST_METHOD") = "POST" then

	if strleach(trim(request("username")))="" then
		Message=Message & "<li>用户名不能为空</li>"
		FoundErr = True
	End If

	if strLeach(trim(request("Password")))="" then
		Message=Message & "<li>密码不能为空</li>"
		FoundErr = True
	End If

	if strCLng(request.form("strCAPTCHA")) <> Cint(Session("GetCAPTCHA")) then
		FoundErr=True
		Message=Message & "<li>验证码输入错误</li>"
	end if

	if FoundErr<>True then

		set rs=server.createobject("adodb.recordset")
		sql="select * from tblSite where UserName='"&strLeach(trim(request("username")))&"' and UserPassword='"&strLeach(md5(trim(request("Password"))))&"'"
		rs.open sql,conn,1,3
		if not (rs.eof or err) then
			session("strUserName") = Trim(rs("UserName"))
			session("strPassword") = Trim(rs("UserPassword"))
			rs.close
			set rs=nothing
			response.redirect "home.asp"
		else
			Message = Message & "用户名或密码错误，请重新输入"
			FoundErr = True
		end if

	end if

end if

if FoundErr<>True then
%>
<html>
<head>
<title><%=smart_SystemName%></title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="style/style.css" rel="stylesheet" type="text/css">
<script type="text/javascript" src="inc/function.js"></script>
<script type="text/javascript" src="jquery/niftycube.js"></script>
<style type="text/css">
div#box{width: 33em;padding: 20px;margin:0 auto;
	background:#eff3ff;color:#000}
h1{font: lighter 200% "Trebuchet MS",Arial sans-serif;color: #208BE1}
h1,p{margin:0;padding:10px 20px}
</style>
<script type="text/javascript">
window.onload=function(){
Nifty("div#box","big");
}
</script>
</head>
<body>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td height="60" bgcolor="#00309c"><table width="100%" border="0" cellspacing="0" cellpadding="5">
      <tr>
        <td width="50%"><a href="http://www.cactussoft.cn" target="_blank"><img src="images/logo.gif" border="0"></a></td>
        <td><table border="0" align="right" cellpadding="5" cellspacing="0">
          <tr>
            <td><a href="mailto:avram@tom.com" class="smartLink">技术支持</a></td>
            <td class="SmartWhite">|</td>
            <td><a href="http://www.cactussoft.cn" title="仙人掌软件" target="_blank" class="smartLink">关于</a></td>
          </tr>
        </table></td>
      </tr>
    </table></td>
  </tr>
  <tr>
    <td height="22" background="images/smart-stat-bg.gif">&nbsp;</td>
  </tr>
  <tr>
    <td height="400"><div align="center"><div id="box">
<table width="100%" border="0" cellspacing="0" cellpadding="5">
            <form name="thisform" method="post" action="Login.asp" target="_parent"><tr>
              <td class="SmartTitle">用户登录</td>
            </tr>
            <tr>
              <td><table width="100%" border="0" cellspacing="0" cellpadding="2">
                <tr>
                  <td width="20%">用户名</td>
                  <td><input name="username" type="text" id="username" style="WIDTH:90%" maxlength="30"></td>
                </tr>
                <tr>
                  <td>密码</td>
                  <td><input name="password" type="password" id="password" style="WIDTH:90%" maxlength="30"></td>
                </tr>
                <tr>
                  <td>验证码</td>
                  <td><input name='strCAPTCHA' type='text' size='6' maxlength="6"><%=GetCAPTCHA()%></td>
                </tr>
              </table></td>
            </tr>
            <tr>
              <td><div align="center">
                <input name="submit" type="image" value="登录" src="images/smart-stat-submit.gif" />
              </div></td>
            </tr></form>
          </table>
</div>
	  </div></td>
  </tr>
</table>
<center>
<span class=smallTxt><%=smart_SystemCopyright%></span><br>
</center>
</body>
</html>
<%
end if

if FoundErr=True then
	call ErrMsg(Message)
end if
%>
