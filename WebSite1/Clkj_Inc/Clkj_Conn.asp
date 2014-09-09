<%

If Request.QueryString<>"" Then  	StopInjection(Request.QueryString)
If Request.Cookies<>"" Then   		StopInjection(Request.Cookies)

Function Stop_Inj(str)
   dim BadStr,myarry
   str=lcase(str)
   BadStr = "exec|insert|select|delete|update|count|chr|mid|master|truncate|char|declare"'此项按实际情况进行增减
   myarry=split(BadStr,"|")   
	for i=0 to ubound(myarry)
	  if instr(str,myarry(i))>0 then
	  	  'response.Write("sorry:"&str&"accc"&myarry(i)&"<br>以免影响你的正常访问")
		  response.Write("<Script Language='javascript'>alert('Excuse me:take a break!');history.back(-1);</Script>")
		  response.End()
	  end if
	next
end function

Sub StopInjection(Values)
	Dim sItem, sValue
    For Each sItem In Values
        sValue = Values(sItem)
        call Stop_Inj(sValue)
    Next
End Sub

%>

<%
'--------------数据链接--------------------------------------------------------------------------------------------------

Dim conn,connstr

on error resume next
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "driver={Microsoft Access Driver (*.mdb)};dbq="&Server.MapPath(""&Datasmdb&"")	
		if err then
			err.clear
			set conn=nothing
			response.write "数据库链接出错,请检查数据路径"
			response.End		 
		End IF

%>
