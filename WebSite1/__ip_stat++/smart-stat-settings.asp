<!--#include file="inc/plus.inc.asp"-->
<%
'**************************************************************
' Software name: SmartStat
' Web: http://www.cactussoft.cn
' Copyright (C) 2008－2010 仙人掌软件 版权所有
'**************************************************************
'判断后台登录是否过期
if GetIsAdmin()=False then
	Response.Write "<script language='javascript'>"
	Response.Write "parent.parent.parent.window.navigate('"&theURL&"login.asp');"
	Response.Write "</script>"
	Response.End
end if
dim FoundErr,Message
act=Trim(Request("act"))

if act="Conserve" then
	call Conserve()
else
	call smarthome()
end if

if FoundErr=True then
	call ErrMsg(Message)
end if

sub smarthome()
	set rs=server.createobject("adodb.recordset")
	sql="select * from tblSite"
	rs.open sql,conn,1,1
	'response.write sql
	'response.End
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<link href="style/style.css" rel="stylesheet" type="text/css">
<title>系统设置</title>
</head>
<body>
<table width="99%" border="0" align="center" cellpadding="3" cellspacing="0">
  <tr>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td class="SmartTitle">系统设置</td>
  </tr>
    <form name="thisform" method="post" action="smart-stat-settings.asp">  
    <tr>
      <td class="SmartText">网站名称</td>
    </tr>
    <tr>
      <td><input name="UserSiteName" type="text" value="<%=Trim(rs("UserSiteName"))%>" style="WIDTH:50%" maxlength="100"></td>
    </tr>
    <tr>
      <td class="SmartText">网站地址</td>
    </tr>
    <tr>
      <td><input name="UserSiteUrl" type="text" value="<%=Trim(rs("UserSiteUrl"))%>" style="WIDTH:50%" maxlength="100"></td>
    </tr>
    <tr>
      <td class="SmartText">管理员登录名</td>
    </tr>
    <tr>
      <td class="SmartText"><input name="UserName" type="text" value="<%=Trim(rs("UserName"))%>" style="WIDTH:50%" maxlength="100" /></td>
    </tr>
    <tr>
      <td class="SmartText">管理员密码</td>
    </tr>
    <tr>
      <td class="SmartText"><input name="UserPassword" type="password" value="<%=Trim(rs("UserPassword"))%>" style="WIDTH:50%" maxlength="100" /></td>
    </tr>
    <tr>
    <td class="SmartText">管理员电子邮件</td>
    </tr>
    <tr>
    <td><input name="UserEmail" type="text" value="<%=Trim(rs("UserEmail"))%>" style="WIDTH:50%" maxlength="100" /></td>
    </tr>
    <tr>
      <td class="SmartText">城市</td>
    </tr>
    <tr>
      <td><%=GetCity(Trim(rs("UserProvince")),Trim(rs("UserCity")))%></td>
    </tr>
    <tr>
      <td><span class="SmartText">网站描述</span></td>
    </tr>
    <tr>
      <td><textarea name="UserContent" cols="45" rows="5" style="WIDTH:50%"><%=Trim(rs("UserContent"))%></textarea></td>
    </tr>
    <tr>
      <td class="SmartText">开始统计日期</td>
    </tr>
    <tr>
      <td><input name="UserStartTime" type="text" value="<%=Trim(rs("UserStartTime"))%>" style="WIDTH:50%" maxlength="100" /></td>
    </tr>
    <tr>
      <td class="SmartText">管理员所有时区</td>
    </tr>
    <tr>
      <td><input name="UserMasterTimeZone" type="text" value="<%=Trim(rs("UserMasterTimeZone"))%>" style="WIDTH:50%" maxlength="100" /></td>
    </tr>
    <tr>
      <td class="SmartText">防刷指数（保存多少IP）</tr>
    <tr>
      <td><input name="UserDelRefresh" type="text" value="<%=Trim(rs("UserDelRefresh"))%>" style="WIDTH:50%" maxlength="100" /></td>
    </tr>
  <tr>
    <td class="SmartText">保存最近×条数据</td>
  </tr>
  <tr>
    <td><input name="UserSaveNum" type="text" value="<%=Trim(rs("UserSaveNum"))%>" style="WIDTH:50%" maxlength="100" /></td>
  </tr><tr>
    <td><input name="act" type="hidden" id="act" value="Conserve"><input type="submit" name="Submit" value="提交(S)" class="SmartButton"></td>
    </tr>
    <tr>
    <td>&nbsp;</td>
    </tr>
    </form>
</table>
<%
rs.Close
end sub

sub Conserve

	SQL="select * from tblSite where Site_ID=" & SiteID & ""
	set rs=conn.execute(SQL)
	if not(rs.bof and rs.eof) then
		strUserPassword=Trim(rs("UserPassword"))
	end if
	rs.close

	if Trim(Request("UserSiteName"))="" or Trim(Request("UserSiteUrl"))="" or Trim(Request("UserName"))="" or Trim(Request("UserPassword"))="" then
		FoundErr=True
		Message=Message & "<li>参数错误</li>"
	end if

	if FoundErr=True then
		Exit sub
	end if
	
	if Trim(Request("UserPassword"))<>strUserPassword then
		strUserPassword=md5(Trim(Request("UserPassword")))
	else
		strUserPassword=strUserPassword
	end if

	Set cmdTemp = Server.CreateObject("ADODB.Command")
	Set InsertCursor = Server.CreateObject("ADODB.Recordset")
	cmdTemp.CommandText = "SELECT * FROM tblSite"
	cmdTemp.CommandType = 1
	Set cmdTemp.ActiveConnection = conn
	InsertCursor.Open cmdTemp, , 1, 3
	InsertCursor("UserSiteName") =Trim(Request("UserSiteName"))
	InsertCursor("UserSiteUrl") =Trim(Request("UserSiteUrl"))
	InsertCursor("UserName") =Trim(Request("UserName"))
	InsertCursor("UserPassword") =strUserPassword
	InsertCursor("UserEmail") =Trim(Request("UserEmail"))
	InsertCursor("UserProvince") =Trim(Request("UserProvince"))
	InsertCursor("UserCity") =Trim(Request("UserCity"))
	InsertCursor("UserContent") =Trim(Request("UserContent"))
	InsertCursor("UserStartTime") =Trim(Request("UserStartTime"))
	InsertCursor("UserMasterTimeZone") =Trim(Request("UserMasterTimeZone"))
	InsertCursor("UserDelRefresh") =Trim(Request("UserDelRefresh"))
	InsertCursor("UserSaveNum") =Trim(Request("UserSaveNum"))
	InsertCursor.Update
	InsertCursor.close
	conn.close
	set InsertCursor=nothing
	set conn=nothing
	
	'成功提示	   
	Call performMsg("操作成功","")
End sub
%>
<center>
<span class=smallTxt><%=smart_SystemCopyright%></span><br>
</center>
</body>
</html>
