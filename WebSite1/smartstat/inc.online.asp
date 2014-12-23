<!--#include file="inc/plus.inc.asp"-->
<%
isCan = smart_IsCan
if isCan then
	Response.Expires = 0

	vIP = Request.ServerVariables("Remote_Addr")
	vAgent = Request.ServerVariables("HTTP_USER_AGENT")
	vPage = Request.ServerVariables("HTTP_REFERER")
	FromUrl = REPLACE(TRIM(Request("fromurl")&""),",","")
	IF FromUrl <> "" THEN
		vPage = FromUrl
	end if
	vAgent = replace(vAgent,"'","''")
	vPage = replace(vPage,"'","''")
	truenow = dateadd("h",0 - smart_ZoneServer,now())
	onnnow = dateadd("s",-2.5 * smart_CheckOnline,truenow)
	set isOnline = conn.Execute("select * from tblOnline where LastTime>#"&onnnow&"# and UserIP='"&vIP&"' and Site_ID=" & SiteID)
	if isOnline.eof then
		set rsOd = conn.Execute("select top 1 id from tblOnline where LastTime<#"&onnnow&"# and Site_ID=" & SiteID & " order by LastTime")
		if rsOd.eof then
			conn.Execute "insert into tblOnline (Site_ID,UserIP,UserAgent,UserPage,OnTime,LastTime) VALUES("&SiteID&",'"&vIP&"','"&vAgent&"','"&vPage&"','"&truenow&"','"&truenow&"')"
		else
			conn.Execute "update tblOnline set Site_ID="&SiteID&",UserIP='"&vIP&"',UserAgent='"&vAgent&"',UserPage='"&vPage&"',Ontime='"&truenow&"',LastTime='"&truenow&"' where id=" & rsOd("id")
		end if
	else
		conn.Execute("update tblOnline set LastTime='"&truenow&"',UserPage='"&vPage&"' where LastTime>#"&onnnow&"# and UserIP='"&vIP&"' and Site_ID=" & SiteID)
	end if

end if'if isCan
Server.Transfer("images/unknown.gif")
%>