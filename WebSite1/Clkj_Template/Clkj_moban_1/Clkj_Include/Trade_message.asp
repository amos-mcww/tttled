<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%Response.Charset= "utf-8"%>
<!--#include file="Clkj_Md5.asp"-->
<!--#include file="Clkj_rehtml.asp"-->
<!--#include file="Clkj_back_Conn.asp"-->
<!--#include file="../Clkj_Template/$TemplateUrl$/Clkj_webconfig/Clkj_Config.asp"-->
<%	
IF request.QueryString("class")="message" then
		Set Rs=server.createobject("ADODB.Recordset")
		Sql="select * from clkj_message"
		Rs.open Sql,conn,1,3
		If  request.form("Name")=""  or request.form("mail")="" or request.Form("title")="" or request.Form("content")="" or request.Form("Phone")="" or request.Form("Region")="" Then
		response.write "<script language='javascript'>alert('Note:with * can not be empty!');history.go(-1);</script>"
		else if request.form("yzm")<>session("GetCode") Then
		response.write "<script language='javascript'>alert('Please enter the verification code!');history.go(-1);</script>"
Else
		rs.addnew
		rs("title")= request.form("title")
		rs("cotant") = request.form("content")
		rs("company") = request.form("Company")
		rs("name") = request.form("Name")
		rs("email") = request.form("mail")
		rs("phone") = request.form("Phone")
		rs("fax") = request.form("Fax")
		rs("contry") = request.form("Region")
		rs("homepage") = request.form("Home")
		rs("messageip")=getIP
		rs.update
		
		email_title="来自公司网站的询盘邮件"
		email_content="标题："&request.form("title")&"<BR>内容："&request.form("content")&"<BR>公司名称："&request.form("company")&"<BR>联系人："&request.form("name")&"<BR>联系邮箱："&request.form("mail")&"<BR>联系电话："&request.form("Phone")&"<BR>传真："&request.form("Fax")&"<BR>区域："&request.form("Region")&"<BR>网站主页："&request.form("Home")
	
		response.write"<script language='javascript'>alert('Thank you for your message, we will contact you as soon as possible!');window.location.href=' ../Feedback.asp?t="&email_title&"&n="&email_content&"';</script>"

End IF
End IF
End IF


IF request.QueryString("class")="mails" then
		Set Rs=server.createobject("ADODB.Recordset")
		Sql="select * from clkj_dy"
		Rs.open Sql,conn,1,3
		if request.Form("mails")<>"" and trim(request.Form("mails"))<>"Enter your E-mail" then
		rs.addnew
		rs("email")=request.Form("mails")
		rs.update
		response.write"<script language='javascript'>alert('Thank you!');history.back(-1);</script>"
		else
		response.write"<script language='javascript'>alert('Enter your e-mail!');history.back(-1);</script>"
		end if
End if


%>