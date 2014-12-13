<!--#include file="../Clkj_Inc/inc.asp"-->
<%

IF Request.Querystring("Edit")="honor" Then
	Set Rs=server.createobject("ADODB.Recordset")
	Sql="select * from clkj_Honor where clkj_ryTitle='"&trim(request.Form("clkj_ryTitle"))&"'"
	Rs.open Sql,conn,1,3
  If Not (Rs.Eof and Rs.Bof) or request.Form("clkj_ryTitle")="" Then
response.write "<script language='javascript'>alert('标题不能为空或有重名,请返回查看');history.go(-1);</script>"
	Else
	Rs.Addnew
	Rs("clkj_ryTitle") = trim(Request.Form("clkj_ryTitle"))
	Rs("clkj_ryImg") = Replace(Request.Form("clkj_ryImg"),"../","")
	Rs("clkj_ryTime") = Request.Form("clkj_ryTime")
	Rs.update
	Response.Redirect "nimda_Honor.asp?Lei=提示,增加成功"
	Rs.close
  End If
Else IF Request.Querystring("Edit")="honor_Edit" Then
	Set Rs=server.createobject("ADODB.Recordset")
	Sql="select * from clkj_Honor where clkj_ryid="&request("clkj_ryid")
	Rs.open Sql,conn,1,3
    Rs("clkj_ryTitle") = trim(Request.Form("clkj_ryTitle2"))
	Rs("clkj_ryImg") = Replace(Request.Form("clkj_ryImg2"),"../","")
	Rs("clkj_ryTime") = Request.Form("clkj_ryTime2")
	Rs.update
	Response.Redirect "nimda_honor_mange.asp?Lei=提示,修改改成功"
	Rs.close
Else If request.QueryString("Class")="hone_Del" Then
		Call DelImage("hone_Del",cint(request("clkj_ryid")))'图片删除
		set Rs=server.CreateObject("adodb.recordset")'数据库记录删除
		Sql="delete * from clkj_Honor where clkj_ryid="&request("clkj_ryid")
		conn.execute Sql
		Response.Redirect "nimda_honor_mange.asp?Lei=提示,删除成功"
End If
End If
End If
%>