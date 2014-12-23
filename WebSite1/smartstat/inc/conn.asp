<%@ CODEPAGE=936 %> 
<% Response.CodePage=936 %> 
<% Response.Charset="gb2312" %>
<%
' Êý¾Ý¿âÄ¿Â¼
DBFolder = "_database/"
DBPath=DBFolder&"VPOBVdLK9pMl4IxG39MI.mdb"

public conn
public DBPath

sub openConn
  on error resume next
  set conn=server.createobject("adodb.connection")
  'conn.Open "driver={Microsoft Access Driver (*.mdb)};dbq=" & DBPath
  DBPath = Server.MapPath( DBPath )
  err=0
  conn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & DBPath
  if err<>0 then
    response.write "<font color=red><b>Error!</b></font><br>There was an error when the system try to open the database."
	response.end
  end if
end Sub

Call openConn

sub closeconn
	conn.Close
	set conn=nothing
end sub
%>