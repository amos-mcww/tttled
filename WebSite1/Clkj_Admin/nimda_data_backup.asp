<!--#include file="../Clkj_Inc/inc.asp"-->
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<link href="backdoor.css" type="text/css" rel="stylesheet">
<title>数据备份</title>
</head>
<body>
<% 
IF request.cookies("userpas")("upas")=md5(request("data_password"))Then
Function CopyTo(ByVal cFile,ByVal toFile)
		cFile=Server.MapPath(cFile) '所要备份的文件
		toFile=Server.MapPath(toFile)'备份文件
		Dim cFso,cf
		set cFso=Server.CreateObject("Scripting.FileSystemObject")
		cFso.fileexists(cFile)
		cFso.Copyfile cFile,toFile
End Function
'-------------------------------
Dim dbpath,bkfolder,bkdbname,fso,fso1
call main()
call main2()
conn.close
set conn=nothing
sub main()
IF request("action")="Backup" then
call backupdata()
Else
%>
<table width="100%" cellpadding="0" cellspacing="0">
  <tr>
    <td height="39" class="tr_bg">&nbsp;<strong>数据备份</strong></td>
  </tr>
</table>
<table cellspacing=1 cellpadding=1 align=center width="100%">
  <tr>
    <th height=25 align="left" > <B>数据库备份</B> </th>
  </tr>
  <form method="post" action="nimda_data_backup.asp?action=Backup">
    <tr>
      <td height=100  style="line-height:150%"> 当前数据库路径(绝对路径)：
        <input type=text size=35 name="DBpath" value="<%=Clkj_mdb%>" readonly="readonly">
        <BR>
        备份数据库目录(绝对路径)：
        <input type=text size=35 name="bkfolder" value="/clkj_Data_back">
        如目录不存在，程序将自动创建<BR>
        备份数据库名称(填写名称)：
        <input type=text size=20 name="bkDBname" value="clkjdateback.mdb">
		<input type="hidden" name="data_password" value="<%=request("data_password")%>">
        如备份目录有该文件，将覆盖，如没有，将自动创建<BR>
        <input type=submit value="备份数据">
      <hr align="center" width="100%" color="#E7EFF8"></td>
    </tr>
  </form>
</table>
<%
		End if
		End sub
		sub main2()
		if request("action")="Restore" then
		Dbpath=request.form("Dbpath")
		backpath=request.form("backpath")
		if Dbpath="" then
		response.write "请输入您要恢复成的数据库全名" 
		else
		Dbpath=server.mappath(Dbpath)		
		End if
		backpath=server.mappath(backpath)
		'Response.write backpath
		Set Fso=server.createobject("scripting.filesystemobject")
		if fso.fileexists(backpath) then       
		 fso.copyfile backpath,Dbpath
		  response.write "<font color=red>成功恢复数据</font>"
		else
		  response.write "<font color=red>备份目录下并无您的备份文件！</font>" 
		End if
		else
		%>
<table align=center cellspacing=1 cellpadding=1 width="100%">
  <tr>
    <th height=25 align="left" > <B>恢复数据库</B> </th>
  </tr>
  <form method="post" action="nimda_data_backup.asp?action=Restore">
    <tr>
      <td height=100 > 当前数据库路径(<span style="line-height:150%">绝对路径</span>)：
        <input type=text size=35 name="DBpath" value="<%=Clkj_mdb%>" readonly="readonly">
        <BR>
        备份数据库路径(<span style="line-height:150%">绝对路径</span>)：
        <input type=text size=35 name="backpath" value="/clkj_Data_back/clkjdateback.mdb">
        <BR>
		<input type="hidden" name="data_password" value="<%=request("data_password")%>">
        <input type="submit" value="恢复数据">
        <hr width="100%" align="center" color="#E7EFF8">
        <font color="#666666">·注意：所有路径都是<span style="line-height:150%">绝对路径</span></font></td>
    </tr>
  </form>
</table>
<%
	End if
	End sub
	sub backupdata()
			Dbpath=request.form("Dbpath")
			Dbpath=server.mappath(Dbpath)
			bkfolder=request.form("bkfolder")
			bkdbname=request.form("bkdbname")
			
			Set Fso=server.createobject("scripting.filesystemobject")
			if fso.fileexists(dbpath) then
						If CheckDir(bkfolder) = True Then
							fso.copyfile dbpath,bkfolder& "\\"& bkdbname
						else
							MakeNewsDir bkfolder
							fso.copyfile dbpath,bkfolder& "\\"& bkdbname
						End if
				response.write "<script language='javascript'>alert('备份数据库成功');history.back(-1);</script>"
			Else
				response.write "<script language='javascript'>alert('找不到您所需要备份的文件');history.back(-1);</script>"
			End if
	End sub
	
	'------------------检查某一目录是否存在-------------------
	Function CheckDir(FolderPath)
			folderpath=Server.MapPath("/")&"\\"&folderpath
			Set fso1 = CreateObject("Scripting.FileSystemObject")
				If fso1.FolderExists(FolderPath) then
					CheckDir = True
				Else
					CheckDir = False
				End if
			Set fso1 = nothing
	End Function
			'-------------根据指定名称生成目录---------
	Function MakeNewsDir(foldername)
			Dim f
			Set fso1 = CreateObject("Scripting.FileSystemObject")
			Set f = fso1.CreateFolder(foldername)
				MakeNewsDir = True
			Set fso1 = nothing
	End Function
Else
%>
<div style="height:100%; width:100%; text-alig:center; margin:auto;" align="center">
<form action="nimda_data_backup.asp" method="post" name="data_back" style="padding-top:240px; padding-bottom:240px; border:1px solid #CCCCCC;">
请输入登陆密码:
<input type="text" size="20" name="data_password">
<input type="submit" name="Submit" value="进入">
</form>
</div>
<%
End IF
%>

</body>
</html>
