<!--#include file="../Clkj_Inc/inc.asp"-->

<%

	''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'  联系人：黑蚂蚁                                        '
	'  联系QQ：1181698019                                   '
	'  msn: zjxyk@hotmail.com                              '
	'  网址：http:/www.sem-cms.com/                         '
	'                                                      ' 
	''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'================ 一级栏目 =========================================================================

If request.QueryString("Class")="big" Then '一级栏目添加
		Set Rs=server.createobject("ADODB.Recordset")
		Sql="select * from clkj_BigClass where clkj_BigClassName = '"&trim(request.Form("big_name"))&"' or clkj_BigClassurl= '"&trim(request.Form("big_wj_name"))&"'"   '判断是否有重名
		Rs.open Sql,conn,1,3
	If Not (Rs.Eof and Rs.Bof) or request.Form("big_name")="" or request.Form("big_wj_name")="" Then
		response.write "<script language='javascript'>alert('[重复添加]或[一级栏目名为空]或[文件类别名称为空]，请返回查看');history.go(-1);</script>"
	Else
			Rs.Addnew
			Rs("clkj_BigClassName") = trim(request.Form("big_name"))
			RS("clkj_BigClassurl") = trim(request.Form("big_wj_name"))
			Rs("clkj_BigClasskey") = request.Form("big_key_name")
			Rs("clkj_BigClassdes") = request.Form("big_ms_name")
			Rs("big_paixu") = request.Form("big_paixu")
			Rs.update
			Response.Redirect "Nimda_class.asp?Lei=增加成功&Edit=B_Z"
			RS.close
	End If
Else If request.QueryString("Class")="big_Edit" Then '一级栏目修改
		Set Rs=server.createobject("ADODB.Recordset")
		Sql="select * from clkj_BigClass where clkj_BigClassID = "&request("clkj_BigClassID") 
		Rs.open Sql,conn,1,3
		Rs("clkj_BigClassName") = trim(request.Form("big_name"))
		RS("clkj_BigClassurl") = trim(request.Form("big_wj_name"))
		Rs("clkj_BigClasskey") = request.Form("big_key_name")
		Rs("clkj_BigClassdes") = request.Form("big_ms_name")
		Rs("big_paixu") = request.Form("big_paixu")
		Rs.update
		
		'修改二级样目下相对应的一级栏目名称
			Set Rs_s=server.createobject("ADODB.Recordset")
			Sql_s="select * from clkj_SmallClass where clkj_BigClassID = "&request("clkj_BigClassID")
			Rs_s.open Sql_s,conn,1,3
			if not Rs_s.eof  then
				do while not rs_s.eof
				 Rs_s("clkj_BigClassName") = trim(request.Form("big_name"))
				 Rs_s.update
				 Rs_s.movenext
				Loop
			
			end if
			Rs_s.close
			
		 '修改产品内容中的一级样目名称
			Set Rs_p=server.createobject("ADODB.Recordset")
			Sql_p="select * from clkj_Products where clkj_BigClassID = "&request("clkj_BigClassID")
			Rs_p.open Sql_p,conn,1,3
			if not Rs_p.eof  then
				do while not rs_p.eof
				 Rs_p("clkj_BigClassName") = trim(request.Form("big_name"))
				 Rs_p.update 
				 Rs_p.movenext
				Loop
			
			end if
			Rs_p.close
		
		Response.Redirect "Nimda_class.asp?Lei=修改成功&Edit=B_E&clkj_BigClassID="&request("clkj_BigClassID")
		Rs.close
Else If request.QueryString("Class")="big_Del" Then'一级栏目删除
		Sql="delete * from clkj_BigClass where clkj_BigClassID = "&request("clkj_BigClassID")   
		conn.execute Sql
		
		Sql_1="delete * from clkj_SmallClass where clkj_BigClassID = "&request("clkj_BigClassID")   
		conn.execute Sql_1
		
		Set Rs_d=server.createobject("ADODB.Recordset")
		Sql_del="select * from clkj_Products where clkj_BigClassID = "&request("clkj_BigClassID")
		Rs_d.open Sql_del,conn,1,1
		
		do while not Rs_d.eof
		
		Call DelFenLeiImages(Rs_d("clkj_prpic"))
		
		Rs_d.Movenext
		Loop
		Rs_d.close
		
		Sql_s="delete * from clkj_Products where clkj_BigClassID = "&request("clkj_BigClassID")
		conn.execute Sql_s
		
		
		Response.Redirect "Nimda_fenlei.asp?Lei=一级栏目删除成功"
End If		
End If
End If

'================ 列出所有一级栏目内容 ====================================================================

Sub Big()
		Set Rs=server.createobject("ADODB.Recordset")
		Sql="select * from clkj_BigClass"
		Rs.open Sql,conn,1,1	
	If Rs.Eof and Rs.Bof Then
		response.write "<tr><td style='padding:5px; height:30px;'>暂无栏目</td></tr>"
	Else
		Do while not Rs.Eof
		n=n+1
		response.write "<td height='30' style='padding:4px;' style='border-bottom:1px dashed #d4d4d4;color:#FF0000'>"&Rs("clkj_BigClassName")&"&nbsp;[<a href='Nimda_class.asp?clkj_BigClassID="&Rs("clkj_BigClassID")&"&Edit=B_E#xy'>修改</a>&nbsp;<a href='Nimda_class.asp?clkj_BigClassID="&Rs("clkj_BigClassID")&"&Class=big_Del' onclick="&chr(34)&"return confirm('是否将此一级栏目删除?');"&chr(34)&">删除</a>]&nbsp;</td>"
		if n mod 5=0 then
		j=j+1
		response.write("<tr></tr>")
		if j=100 then exit do
		end if
		Rs.Movenext
		Loop
		Rs.close		
	End If				
End Sub

Sub Big_input()'所有一级栏目的栏目，表单形式
		Set Rs=server.createobject("ADODB.Recordset")
		Sql="select * from clkj_BigClass"
		Rs.open Sql,conn,1,1	
	If Rs.Eof and Rs.Bof Then
		response.write "暂无栏目"
	Else
		Do while not RS.Eof
		response.write "<option value="
		response.write rs("clkj_BigClassID")
		IF rs("clkj_BigClassID")=cint(request("clkj_BigClassID")) then
		response.write" selected=selected"
		end if
		response.write">"
		response.write rs("clkj_BigClassName")
		response.write "</option>"

		Rs.Movenext
		Loop
		Rs.close
	End If
End Sub

'============================== 二级栏目 ===============================================================================

If request.QueryString("Class")="small" Then '二级栏目添加
		Set Rs=server.createobject("ADODB.Recordset")
		Sql="select * from clkj_SmallClass where clkj_SmallClassName = '"&trim(request.Form("small_name"))&"' or clkj_SmallClassurl= '"&trim(request.Form("small_wj_name"))&"'"   '判断是否有重名
		Rs.open Sql,conn,1,3
	If Not (Rs.Eof and Rs.Bof) or request.Form("small_name")="" or request.form("big_name")="" or request.form("small_wj_name")=""Then
		response.write "<script language='javascript'>alert('[重复添加]或[二级栏目名为空]或[文件类别名称为空]，请返回查看');history.go(-1);</script>"
	Else
			Rs.Addnew
			Rs("clkj_BigClassID") = trim(request.Form("big_name"))
			'读出大类别名称，写入小类别表中开始
				Set Rss=server.createobject("ADODB.Recordset")
				Sqls="select * from clkj_BigClass where clkj_BigClassID ="&trim(request.Form("big_name"))
				Rss.open Sqls,conn,1,1
				big_name=Rss("clkj_BigClassName")
				
			'读出大类别名称，写入小类别表中结束
			Rs("clkj_BigClassName") = big_name
			Rs("clkj_smallClassName") = trim(request.Form("small_name"))
			RS("clkj_smallClassurl") = trim(request.Form("small_wj_name"))
			Rs("clkj_smallClasskey") = request.Form("small_key_name")
			Rs("clkj_smallClassdes") = request.Form("small_ms_name")
			Rs("small_paixu") = request.Form("small_paixu")
			if request.Form("clkj_tj")="yes" then Rs("clkj_tj")=1 else Rs("clkj_tj")=0
			Rs.update
			Response.Redirect "Nimda_class.asp?Lei=二级栏目增加成功&Edit=S_Z"
			RS.close
			Rss.close
	End If
Else If request.QueryString("Class")="small_Edit" Then '二级栏目修改
		Set Rs=server.createobject("ADODB.Recordset")
		Sql="select * from clkj_SmallClass where clkj_SmallClassID = "&request("clkj_SmallClassID") 
		Rs.open Sql,conn,1,3
		Rs("clkj_BigClassID") = trim(request.Form("big_name"))
				'读出大类别名称，写入小类别表中开始
				Set Rss=server.createobject("ADODB.Recordset")
				Sqls="select * from clkj_BigClass where clkj_BigClassID ="&trim(request.Form("big_name"))
				Rss.open Sqls,conn,1,1
				big_name=Rss("clkj_BigClassName")
		      '读出大类别名称，写入小类别表中结束
		Rs("clkj_BigClassName") = big_name
		Rs("clkj_smallClassName") = trim(request.Form("small_name"))
		RS("clkj_smallClassurl") = trim(request.Form("small_wj_name"))
		Rs("clkj_smallClasskey") = request.Form("small_key_name")
		Rs("clkj_smallClassdes") = request.Form("small_ms_name")
		Rs("small_paixu") = request.Form("small_paixu")
		if request.Form("clkj_tj")="yes" then Rs("clkj_tj")=1 else Rs("clkj_tj")=0
		Rs.update
		
		'修改二级栏目时相对应的修改产品中相对应的类别
			Set Rs_sp=server.createobject("ADODB.Recordset")
			Sql_sp="select * from clkj_Products where clkj_SmallClassID = "&request("clkj_SmallClassID")
			Rs_sp.open Sql_sp,conn,1,3
			if not rs_sp.eof then
				do while not rs_sp.eof
				 Rs_sp("clkj_SmallClassName") = trim(request.Form("small_name"))
				 Rs_sp("clkj_BigClassName")=big_name
				 Rs_sp.update
				 Rs_sp.movenext
				Loop
			 end if
			Rs_sp.close
			
		Response.Redirect "Nimda_class.asp?Lei=二级栏目修改成功&Edit=S_E&clkj_SmallClassID="&request("clkj_SmallClassID")
		Rs.close
		
Else If request.QueryString("Class")="small_Del" Then'二级栏目删除
		Sql="delete * from clkj_SmallClass where clkj_SmallClassID = "&request("clkj_SmallClassID")   	
		conn.execute Sql
		
		Set Rs_d=server.createobject("ADODB.Recordset")
		Sql_del="select * from clkj_Products where clkj_SmallClassID = "&request("clkj_SmallClassID") 
		Rs_d.open Sql_del,conn,1,1
		
		do while not Rs_d.eof
		
		Call DelFenLeiImages(Rs_d("clkj_prpic"))
		
		Rs_d.Movenext
		Loop
		Rs_d.close
		
		Sql_s="delete * from clkj_Products where clkj_SmallClassID = "&request("clkj_SmallClassID")   
		conn.execute Sql_s
		
		
		Response.Redirect "Nimda_fenlei.asp?Lei=二级删除成功"
End If		
End If
End If

Sub small()'列出所有二级栏目		
	Set Rs=server.createobject("ADODB.Recordset")
		Sql="select * from clkj_SmallClass"
		Rs.open Sql,conn,1,1	
	If Rs.Eof and Rs.Bof Then
		response.write "<tr><td style='padding:5px; height:30px;'>暂无栏目</td></tr>"
	Else
		Do while not Rs.Eof
		n=n+1
		response.write"<td height='30' style='padding:4px;' style='border-bottom:1px dashed #d4d4d4;color:#FF0000'>"&Rs("clkj_SmallClassName")&"&nbsp;[<a href='Nimda_class.asp?clkj_SmallClassID="&Rs("clkj_SmallClassID")&"&Edit=S_E#xr'>修改</a>&nbsp;<a href='Nimda_class.asp?clkj_SmallClassID="&Rs("clkj_SmallClassID")&"&Class=small_Del' onclick="&chr(34)&"return confirm('是否将此二级栏目删除?');"&chr(34)&">删除</a>]&nbsp;</td>"
		if n mod 5=0 then
		j=j+1
		response.write("<tr></tr>")
		if j=100 then exit do
		end if
		Rs.Movenext
		Loop
		Rs.close		
	End If
End Sub

Sub small_lie()'列出所有二级栏目
	Set Rs=server.createobject("ADODB.Recordset")
		Sql="select * from clkj_SmallClass where clkj_BigClassID="&cint(request("clkj_BigClassID"))
		Rs.open Sql,conn,1,1	
	If Rs.Eof or Rs.Bof Then
		response.write "暂无栏目"
	Else
		Do while not Rs.Eof
		IF rs("clkj_SmallClassID")=cint(request("clkj_SmallClassID"))Then
		response.write "<a href=nimda_products.asp?clkj_SmallClassID="&Server.UrlEncode(rs("clkj_SmallClassID"))&"&clkj_BigClassID="&request("clkj_BigClassID")&"&clkj_BigClassName="&Server.UrlEncode(request("clkj_BigClassName"))&"&clkj_SmallClassName="&Server.UrlEncode(RS("clkj_SmallClassName"))&"&sf=s><font color='#FF0000;'><b>"&Rs("clkj_SmallClassName")&"</b></font></a>&nbsp;&nbsp;&nbsp;"
		Else
		response.write "<a href=nimda_products.asp?clkj_SmallClassID="&Server.UrlEncode(rs("clkj_SmallClassID"))&"&clkj_BigClassID="&request("clkj_BigClassID")&"&clkj_BigClassName="&Server.UrlEncode(request("clkj_BigClassName"))&"&clkj_SmallClassName="&Server.UrlEncode(RS("clkj_SmallClassName"))&"&sf=s>"&Rs("clkj_SmallClassName")&"</a>&nbsp;&nbsp;&nbsp;"
		End if
		Rs.Movenext
		Loop
		Rs.close		
	End If
End Sub

Sub big_lie()'列出所有一级栏目		
	Set Rs=server.createobject("ADODB.Recordset")
		Sql="select * from clkj_BigClass"
		Rs.open Sql,conn,1,1	
	If Rs.Eof or Rs.Bof Then
		response.write "暂无栏目"
	Else
		Do while not Rs.Eof
		IF rs("clkj_BigClassID")=cint(request("clkj_BigClassID")) Then
		response.write "<a href=nimda_products.asp?clkj_BigClassID="&rs("clkj_BigClassID")&"&clkj_BigClassName="&Server.UrlEncode(rs("clkj_BigClassName"))&"&sf=b><font color='#FF0000;'><b>"&Rs("clkj_BigClassName")&"</b></font></a>&nbsp;&nbsp;"
		Else
		response.write "<a href=nimda_products.asp?clkj_BigClassID="&rs("clkj_BigClassID")&"&clkj_BigClassName="&Server.UrlEncode(rs("clkj_BigClassName"))&"&sf=b>"&Rs("clkj_BigClassName")&"</a>&nbsp;&nbsp;"
		End if
		Rs.Movenext
		Loop
		Rs.close		
	End If
End Sub

'============================== 导航管理 ===============================================================================

If request.QueryString("Class")="meun" Then '一级栏目添加
		Set Rs=server.createobject("ADODB.Recordset")
		Sql="select * from clkj_menu where clkj_menutitle = '"&trim(request.Form("meun_name"))&"'"   '判断是否有重名
		Rs.open Sql,conn,1,3
	If Not (Rs.Eof and Rs.Bof) or request.Form("meun_name")="" Then
		response.write "<script language='javascript'>alert('重复添加或栏目名为空，请返回查看');history.go(-1);</script>"
	Else
			Rs.Addnew
			Rs("clkj_menutitle") = trim(request.Form("meun_name"))
			RS("clkj_menuurl") = request.Form("meun_dress")
			Rs("clkj_menukey") = request.Form("meun_key")
			Rs("clkj_menudes") = request.Form("meun_ms")
			Rs("clkj_paixu") = request.Form("meun_px")
			
			Rs.update
			Response.Redirect "Nimda_menu.asp?Lei=栏目增加成功"
			RS.close
	End If
Else If request.QueryString("Class")="meun_edit" Then '栏目修改
		Set Rs=server.createobject("ADODB.Recordset")
		Sql="select * from clkj_menu where clkj_menuid = "&request("clkj_menuid")
		Rs.open Sql,conn,1,3
		Rs("clkj_menutitle") = trim(request.Form("meun_name"))
		RS("clkj_menuurl") = request.Form("meun_dress")
		Rs("clkj_menukey") = request.Form("meun_key")
		Rs("clkj_menudes") = request.Form("meun_ms")
		Rs("clkj_paixu") = request.Form("meun_px")
		
		Rs.update
		Response.Redirect "Nimda_menu.asp?Lei=栏目修改成功&Edit=M_E&clkj_menuid="&request("clkj_menuid")
		Rs.close
Else If request.QueryString("Class")="meun_Del" Then'栏目删除
		Sql="delete * from clkj_menu where clkj_menuid = "&request("clkj_menuid")   
		conn.execute Sql
		Response.Redirect "Nimda_menu.asp?Lei=栏目删除成功"
End If		
End If
End If

Sub Menu_menu()'列出所有菜单栏目
		Set Rs=server.createobject("ADODB.Recordset")
		Sql="select * from clkj_menu order by clkj_paixu,clkj_menuid asc"
		Rs.open Sql,conn,1,1	
	If Rs.Eof and Rs.Bof Then
		response.write "<tr><td style='padding:5px; height:30px;'>暂无栏目</td></tr>"
	Else
		Do while not Rs.Eof
		n=n+1
		response.write "<td height="&chr(34)&"30"&chr(34)&" style='padding:4px;border-bottom:1px dashed #d4d4d4;color:#FF0000'>"&Rs("clkj_menutitle")&"&nbsp;[<a href='Nimda_menu.asp?clkj_menuid="&Rs("clkj_menuid")&"&Edit=M_E'>修</a>&nbsp;<a href='Nimda_menu.asp?clkj_menuid="&Rs("clkj_menuid")&"&Class=meun_Del' onclick="&chr(34)&"return confirm('是否将此栏目删除?');"&chr(34)&">删</a>]&nbsp;</td>"
		if n mod 5=0 then
		j=j+1
		response.write("<tr></tr>")
		if j=100 then exit do
		end if
		Rs.Movenext
		Loop
		Rs.close
	End If
End Sub

'============================== 参数设置 ===============================================================================
IF Request.Querystring("Class")="cansu" and trim(Request.Form("clkj_name"))<>"" and trim(request.form("index_key"))<>"" and Request.Form("jiaodu")<>"" Then

	Set Rs=server.createobject("ADODB.Recordset")
	Sql="select * from clkj_config where clkj_config_id=1"
	Rs.open Sql,conn,1,3
	
	Rs("index_des") = request.form("index_des")
	Rs("index_key") = request.form("index_key")
	Rs("clkj_config_title") = Request.Form("clkj_name")
	Rs("clkj_config_url") = Request.Form("clkj_web")
	Rs("clkj_config_email") = Request.Form("clkj_mail")
	Rs("clkj_config_sltk") = Request.Form("clkj_pic_w")
	Rs("clkj_config_sltg") = Request.Form("clkj_pic_h")
	Rs("clkj_config_tel") = Request.Form("clkj_tel")
	Rs("clkj_config_bottom") = Request.Form("clkj_bottom")
	Rs("clkj_config_logo") = Replace(Request.Form("prpic"),"../","./")
	Rs("clkj_bottom_key") = Request.Form("clkj_bottom_key")
	Rs("clkj_gs_dz") = Request.Form("clkj_gs_dz")
	Rs("clkj_cz") = Request.Form("clkj_cz")
	Rs("clkj_lxr") = Request.Form("clkj_lxr")
	Rs("clkj_profile") = Request.Form("clkj_profile")
	Rs("clkj_pro_sl") = int(Request.Form("clkj_pro_sl"))
	Rs("logo_ms") =  Request.Form("logo_ms")
	Rs("adv") =  Replace(Request.Form("adv"),"../","")
	Rs("adv_adress") = Request.Form("adv_adress")
	Rs("msn") = Request.Form("msn")
	Rs("skype") = Request.Form("skype")
	Rs("ucode")= Request.Form("ucode")
	Rs("pictext") = Request.Form("pictext")
	Rs("picimg") = Request.Form("picimg")
	Rs("yanse") = Request.Form("yanse")
	Rs("jiaodu") = int(Request.Form("jiaodu"))
	
	IF Request.Form("picof")="yes" Then
		Rs("picof")=1
	Else
		Rs("picof")=0
	End IF
	
	
	''''''邮件参数设置''''''''''''''''''''''''''
	Rs("email_user") = trim(Request.Form("email_user"))
	Rs("email_pas") = trim(Request.Form("email_pas"))
	Rs("email_server") = trim(Request.Form("email_server"))
	Rs("email_d") = trim(Request.Form("email_d"))
	Rs("emailsend") = trim(Request.Form("emailsend"))'邮件接收地址
	
	IF Request.Form("clkj_links_k")="yes" Then
		rs("clkj_config_yq_open")=1
	Else
		rs("clkj_config_yq_open")=0
	End IF
	
	IF Request.Form("clkj_links_key")="yes" Then
		rs("clkj_config_key_open")=1
	Else
		rs("clkj_config_key_open")=0
	End IF
	Rs("clkj_meta")=request.Form("clkj_meta")
	Rs("clkj_addthis")=request.Form("clkj_addthis")
	Rs.update
	Response.Redirect "Nimda_cansu.asp?Lei=成功参数设置"
	Rs.close


	
End If

'============================== 友情链接 ===============================================================================
IF Request.Querystring("Class")="links" Then
	Set Rs=server.createobject("ADODB.Recordset")
	Sql="select * from clkj_Links where clkj_title='"&trim(request.Form("clkj_name"))&"'"
	Rs.open Sql,conn,1,3
  If Not (Rs.Eof and Rs.Bof) or request.Form("clkj_name")="" Then
response.write "<script language='javascript'>alert('重复添加或链接名为空，请返回查看');history.go(-1);</script>"

	Else
	Rs.Addnew
	Rs("clkj_title") = trim(Request.Form("clkj_name"))
	Rs("clkj_LinkURL") = Request.Form("clkj_links")
	Rs("linkspic")= replace(Request.Form("linkspic"),"../","")
	Rs.update
	Response.Redirect "Nimda_links.asp?Lei=增加改成功"
	Rs.close
  End If
Else IF Request.Querystring("Class")="links_Edit" Then
	Set Rs=server.createobject("ADODB.Recordset")
	Sql="select * from clkj_Links where clkj_yqId="&request("clkj_yqId")
	Rs.open Sql,conn,1,3
    Rs("clkj_title") = trim(Request.Form("clkj_name"))
	Rs("clkj_LinkURL") = Request.Form("clkj_links")
	Rs("linkspic")= replace(Request.Form("linkspic"),"../","")
	Rs.update
	Response.Redirect "Nimda_links.asp?Lei=修改成功"
	Rs.close
Else If request.QueryString("Class")="links_Del" Then'栏目删除
		Sql="delete * from clkj_Links where clkj_yqId="&request("clkj_yqId")  
		conn.execute Sql
		Response.Redirect "Nimda_links.asp?Lei=删除成功"
End If			
End If
End If

'============================== 密码管理 ===============================================================================
IF Request.Querystring("Class")="pas" Then
	Set Rs=server.createobject("ADODB.Recordset")
	Sql="select * from clkj_admin where clkj_admin='"&trim(request.Form("user_name"))&"'"
	Rs.open Sql,conn,1,3
  If Not (Rs.Eof and Rs.Bof) or request.Form("user_name")="" Then
response.write "<script language='javascript'>alert('重复添加或用户名为空，请返回查看');history.go(-1);</script>"
	Else
	Rs.Addnew
	Rs("clkj_admin") = trim(Request.Form("user_name"))
	Rs("clkj_password") = Md5(Request.Form("user_pas"))
	Rs("clkj_flg") = Request.Form("clkj_flg")
	Rs.update
	Response.Redirect "Nimda_user.asp?Lei=增加成功"
	Rs.close
   End if
Else IF Request.Querystring("Class")="pas_Edit" Then
	Set Rs=server.createobject("ADODB.Recordset")
	Sql="select * from clkj_admin where clkj_id="&request("clkj_id")
	Rs.open Sql,conn,1,3
	Rs("clkj_admin") = trim(Request.Form("user_name"))
		IF trim(request.Form("user_pas"))="" or len(request.form("user_pas"))=0 Then
		Rs("clkj_password")=Rs("clkj_password")
		Else
		Rs("clkj_password") = Md5(Request.Form("user_pas"))
		End IF
	Rs("clkj_flg") = Request.Form("clkj_flg")
	Rs.update
	Response.Redirect "Nimda_user.asp?Lei=修改成功"
	Rs.close
Else If request.QueryString("Class")="pas_Del" Then'栏目删除
		Sql="delete * from clkj_admin where clkj_id="&request("clkj_id")  
		conn.execute Sql
		Response.Redirect "Nimda_user.asp?Lei=删除成功"
End If			
End If
End If

'============================== 用户列表 ===============================================================================

Sub user()		
		Set Rs=server.createobject("ADODB.Recordset")
		Sql="select * from clkj_admin"
		Rs.open Sql,conn,1,1	
	If Rs.Eof and Rs.Bof Then
		response.write "<tr><td style='padding:5px; height:30px;'>暂无用户</td></tr>"
	Else
		Do while not Rs.Eof
		n=n+1
		IF Rs("clkj_flg")=1 then
		flg="管理员"
		Else
		flg="普通员"
		End if
		response.write "<td height='30' style='padding:4px;' style='border-bottom:1px dashed #d4d4d4;color:#FF0000'>"&flg&"&nbsp;"&Rs("clkj_admin")&"&nbsp;[<a href='Nimda_user.asp?clkj_id="&Rs("clkj_id")&"&Edit=B_E#xy'>修改</a>&nbsp;<a href='Nimda_user.asp?clkj_id="&Rs("clkj_id")&"&Class=pas_Del' onclick="&chr(34)&"return confirm('是否将此删除?');"&chr(34)&">删除</a>]&nbsp;</td>"
		if n mod 5=0 then
		j=j+1
		response.write("<tr></tr>")
		if j=100 then exit do
		end if
		Rs.Movenext
		Loop
		Rs.close		
	End If				
End Sub


Sub userp()		
		Set Rs=server.createobject("ADODB.Recordset")
		Sql="select * from clkj_admin where clkj_id="&session("clkj_id")
		Rs.open Sql,conn,1,1	
	If Rs.Eof and Rs.Bof Then
		response.write "<tr><td style='padding:5px; height:30px;'>暂无用户</td></tr>"
	Else
		Do while not Rs.Eof
		n=n+1
		IF Rs("clkj_flg")=1 then
		flg="管理员"
		Else
		flg="普通员"
		End if
		response.write "<td height='30' style='padding:4px;' style='border-bottom:1px dashed #d4d4d4;color:#FF0000'>"&flg&"&nbsp;"&Rs("clkj_admin")&"&nbsp;[<a href='Nimda_user.asp?clkj_id="&Rs("clkj_id")&"&Edit=B_E#xy'>修改</a>]</td>"
		if n mod 5=0 then
		j=j+1
		response.write("<tr></tr>")
		if j=100 then exit do
		end if
		Rs.Movenext
		Loop
		Rs.close		
	End If				
End Sub


'============================== 新闻信息管理 ===============================================================================

IF Request.Querystring("Class")="content" Then
		Set Rs=server.createobject("ADODB.Recordset")
		Sql="select * from clkj_News"
		Rs.open Sql,conn,1,3
		If request.Form("Lei")="" Then
response.write "<script language='javascript'>alert('新闻类别不能为空，请选择');history.go(-1);</script>"
	    Else
		Rs.Addnew
		Rs("clkj_new_lei") = Request.Form("Lei")
		Rs("clkj_news_Title") = Request.Form("c_title")
		Rs("clkj_news_content") = Request.Form("content")
		Rs("clkj_news_user") = Request.Form("c_ru")
		Rs("clkj_news_db") = Request.Form("clkj_news_db")
		Rs("clkj_news_key") = Request.Form("clkj_news_key")
		Rs("clkj_news_time") = Request.Form("time") 
		Rs.update
		Response.Redirect "Nimda_news.asp?Lei=添加成功"
		Rs.close
		End IF
   Else IF Request.Querystring("Class")="content_Edit" Then
        Set Rs=server.createobject("ADODB.Recordset")
		Sql="select * from clkj_News where clkj_newsid="&request("clkj_newsid")
		Rs.open Sql,conn,1,3
		Rs("clkj_new_lei") = Request.Form("Lei")
		Rs("clkj_news_Title") = Request.Form("c_title")
		Rs("clkj_news_content") = Request.Form("content")
		Rs("clkj_news_db") = Request.Form("clkj_news_db")
		Rs("clkj_news_user") = Request.Form("c_ru")
		Rs("clkj_news_key") = Request.Form("clkj_news_key")
		Rs("clkj_news_time") = Request.Form("time") 
		Rs.update
		Response.Redirect "Nimda_news_mange.asp?Lei=修改成功"
		Rs.close
	Else IF Request.Querystring("Class")="news_Del" Then
		Sql="delete * from clkj_News where clkj_newsid="&request("clkj_newsid")  
		conn.execute Sql
		Response.Redirect "Nimda_news_mange.asp?Lei=删除成功"
  End IF		
  End IF	
End IF

'============================== 产品管理 ===============================================================================
IF Request.Querystring("Class")="Product" Then
		Set Rs=server.createobject("ADODB.Recordset")
		Sql="select * from clkj_Products"
		Rs.open Sql,conn,1,3
	IF request.Form("big_Lei")="" Then
response.write "<script language='javascript'>alert('注意：产品一级类别,产品文件名,不能为空');history.go(-1);</script>"
	Else
		Rs.Addnew
		Rs("clkj_BigClassID") = Request.Form("big_Lei")
				'读出大类别名称，写入表中开始
				Set Rss=server.createobject("ADODB.Recordset")
				Sqls="select * from clkj_BigClass where clkj_BigClassID ="&trim(request.Form("big_Lei"))
				Rss.open Sqls,conn,1,1
				if not Rss.eof then
					bigg_name=Rss("clkj_BigClassName")
					bigll_ml=Rss("clkj_BigClassurl")
				End if
				Rss.close
		      '读出大类别名称，写入表中结束
			  
			   if  isNumeric(Request.Form("small_Lei"))=true then			
		        Rs("clkj_SmallClassID") = Request.Form("small_Lei")
				'读出小类别名称，写入表中开始
				Set Rsss=server.createobject("ADODB.Recordset")
				Sqlss="select * from clkj_SmallClass where clkj_SmallClassID ="&trim(request.Form("small_Lei"))
				Rsss.open Sqlss,conn,1,1
				smalll_name=Rsss("clkj_SmallClassName")
				smalll_ml=Rsss("clkj_SmallClassurl")'小类换大类链接地址
				Rsss.close
				end if
		      '读出小类别名称，写入表中结束			
		Rs("clkj_SmallClassName") = smalll_name
		Rs("clkj_BigClassName") = bigg_name
		Rs("clkj_ml_cn")=bigll_ml
		Rs("clkj_ml_cn")=smalll_ml
		Rs("clkj_paixu") = Request.Form("clkj_paixu")'排序
		Rs("clkj_db")=Request.Form("clkj_db")
		Rs("clkj_prtitle") = Request.Form("clkj_prtitle")
		Rs("clkj_prcontent") = Request.Form("clkj_prcontent")
		Rs("clkj_prkey") = Request.Form("clkj_prkey")
		Rs("clkj_prprdes") = Request.Form("clkj_prprdes")
		IF Request.Form("clkj_hot")="yes" then
		Rs("clkj_hot") = 1
		Else
		Rs("clkj_hot") = 0
		End IF
		
		'产品图片上传
		ProductsPic=Request.Form("clkj_prpic")
		Rs("clkj_prpic") =Replace(Request.Form("clkj_prpic"),"../","")
		'--------------------------------------------------------------
		Rs("clkj_pr_url") = trim(Request.Form("clkj_pr_url"))
		Rs("clkj_ru") = Request.Form("c_ru")
		Rs("clkj_time") = Request.Form("clkj_time")
		'TT=request.Form("clkj_time")		
		Rs.update
		
		'判断是否支持图片组件,生成小图及打水印
		
		IF isobjinstalled("Persits.Jpeg") Then
			ProductsPic=Split(ProductsPic,",")'分割图片		
		For Each PStrss In ProductsPic
		'response.write PStrss
		    IF PStrss<>" " and PStrss<>""  Then 
			Call TradeCmsJpeg(Imgop,PStrss,ImgPic,ImgText,ImgWidth,Imgyanse,Imgjiaodu)	
					
			End IF
		next
			
	   End iF
	   
		'call shengchen(tt)增加生成静态
		Response.Redirect "Nimda_products.asp"
		'Response.Redirect "Nimda_product.asp?Lei=添加成功"
		Rs.close
	End IF	
	
Else IF Request.Querystring("Class")="Pro_Edit" Then	
	Set Rs=server.createobject("ADODB.Recordset")
		if request.Form("lscp")="yes" then'发布类似产品
		Sql="select * from clkj_Products"
		else
		Sql="select * from clkj_Products where clkj_prid="&request("clkj_prid")
		end if
		Rs.open Sql,conn,1,3
	IF request.Form("big_Lei")="" Then
response.write "<script language='javascript'>alert('注意：产品一级类别,产品文件名,不能为空');history.go(-1);</script>"
	Else
	
	if request.Form("lscp")="yes" then'发布类似产品
	Rs.Addnew
	end if
		Rs("clkj_BigClassID") = request("clkj_BigClassID")
				'读出大类别名称，写入表中开始
				Set Rss=server.createobject("ADODB.Recordset")
				Sqls="select * from clkj_BigClass where clkj_BigClassID ="&trim(request("clkj_BigClassID"))
				Rss.open Sqls,conn,1,1
				if not Rss.eof then
					bigg_name=Rss("clkj_BigClassName")
					bigll_ml=Rss("clkj_BigClassurl")
				End if
				Rss.close
		      '读出大类别名称，写入表中结束
			  if  isNumeric(Request.Form("small_Lei"))=true then
			  
		        Rs("clkj_SmallClassID") = Request.Form("small_Lei")
				'读出小类别名称，写入表中开始
				Set Rsss=server.createobject("ADODB.Recordset")
				Sqlss="select * from clkj_SmallClass where clkj_SmallClassID ="&trim(request.Form("small_Lei"))
				Rsss.open Sqlss,conn,1,1
				smalll_name=Rsss("clkj_SmallClassName")
				smalll_ml=Rsss("clkj_SmallClassurl")
				Rsss.close
				
				end if
		      '读出小类别名称，写入表中结束			
		Rs("clkj_SmallClassName") = smalll_name
		Rs("clkj_BigClassName") = bigg_name
		Rs("clkj_ml_cn")=bigll_ml
		Rs("clkj_bl_cn")=smalll_ml
		Rs("clkj_paixu") = Request.Form("clkj_paixu")'排序
		Rs("clkj_db")=Request.Form("clkj_db")
		Rs("clkj_prtitle") = Request.Form("clkj_prtitle")
		Rs("clkj_prcontent") = Request.Form("clkj_prcontent")
		Rs("clkj_prkey") = Request.Form("clkj_prkey")
		Rs("clkj_prprdes") = Request.Form("clkj_prprdes")
		IF Request.Form("clkj_hot")="yes" then
		Rs("clkj_hot") = 1
		Else
		Rs("clkj_hot") = 0
		End IF
		TT=request.Form("clkj_time")
		ProductsPicx=Request.Form("clkj_prpic")
		Rs("clkj_prpic") = Replace(Request.Form("clkj_prpic"),"../","")
		Rs("clkj_pr_url") = trim(Request.Form("clkj_pr_url"))
		Rs("clkj_ru") = Request.Form("c_ru")
		Rs("clkj_time") = Request.Form("clkj_time")
		Rs.update
		
		'判断是否支持图片组件	
		IF isobjinstalled("Persits.Jpeg") Then
		
			ProductsPicx=Split(ProductsPicx,",")'分割图片	
				
			For Each XStrss In ProductsPicx
			 IF trim(XStrss)<>" " and trim(XStrss)<>"" Then'图片不为空执行
				if instr(xStrss,"../") <> 0 then

					Call TradeCmsJpeg(Imgop,trim(xStrss),ImgPic,ImgText,ImgWidth,Imgyanse,Imgjiaodu)'图片打水印

				End IF
			End if
				
			next
		End IF
		
		'链接跳回

		if request.form("sf")="s" then
			Response.Redirect "nimda_products.asp?clkj_BigClassID="&request("clkj_BigClassID")&"&clkj_BigClassName="&bigg_name&"&clkj_SmallClassID="&Request.form("small_Lei")&"&clkj_SmallClassName="&smalll_name&"&ToPage="&Request.form("ToPage")&"&sf=s"
			
		else if request.form("sf")="b" then
			Response.Redirect "nimda_products.asp?clkj_BigClassID="&request("clkj_BigClassID")&"&clkj_BigClassName="&bigg_name&"&ToPage="&Request.form("ToPage")&"&sf=b"
		else
		Response.Redirect "nimda_products.asp?ToPage="&Request.form("ToPage")
		end if
		end if	
	Rs.close	
	End IF
	
Else IF Request.Querystring("Class")="P_Del" Then
		delid=Split(request("delid"),",")'图片删除
		For Each Strss In delid
			Call DelImage("P_Del",cint(Strss))
		next
		Sql="delete * from clkj_Products where clkj_prid in ("&request("delid")&")"
		conn.execute Sql
		'Response.Redirect "Nimda_products.asp?Lei=删除成功"
		if request("clkj_BigClassID")<>"" and  request("clkj_SmallClassID")=""then
		Response.Redirect "Nimda_products.asp?clkj_BigClassID="&request("clkj_BigClassID")&"&clkj_BigClassName="&request.QueryString("clkj_BigClassName")&"&sf="&request.QueryString("sf")&"&ToPage="&request("ToPage")&"&Lei=删除成功"
		else if request("clkj_SmallClassID")<>"" then
		
		Response.Redirect "Nimda_products.asp?clkj_BigClassID="&request("clkj_BigClassID")&"&clkj_BigClassName="&request.QueryString("clkj_BigClassName")&"&clkj_SmallClassID="&request("clkj_SmallClassID")&"&clkj_SmallClassName="&request.QueryString("clkj_SmallClassName")&"&sf="&request.QueryString("sf")&"&ToPage="&request("ToPage")&"&Lei=删除成功"
		else		
		Response.Redirect "Nimda_products.asp?ToPage="&request("ToPage")&"&Lei=删除成功"
		
		end if
		end if
		
  End IF
	
End IF	
End IF


'栏目
IF Request.Querystring("Class")="Moban" Then

		Set Rs=server.createobject("ADODB.Recordset")
		Sql="select * from clkj_config where clkj_config_id=1"
		Rs.open Sql,conn,1,3
		Rs("clkj_daohang")=request.Form("daohang")
		Rs.update
		Response.Redirect "Nimda_moban.asp?Lei=生成栏目成功"
End IF


'后台进入
IF Request.Querystring("Class")="userpas" Then
	Dim username,password
	username=replace(trim(request.Form("username")),"'","")
	password=md5(replace(trim(request.Form("userpas")),"'",""))
		Set Rs=server.createobject("adodb.recordset")
		Sql="select * from  clkj_admin  where clkj_admin='"&username&"' and clkj_password='"&password&"'"
		Rs.open sql,conn,1,1
			IF Rs.eof or Rs.bof then
				  response.write "<script language='javascript'>alert('用户名域密码错误!');top.location.href='index.html';</script>"
				  rs.close
				  set rs=nothing
				  response.end()
			Else
				session("username")=rs("clkj_admin")
				session("userpas")=rs("clkj_password")
				session("clkj_id")=rs("clkj_id")
				session("clkj_flg")=rs("clkj_flg")
				response.cookies("username")("uname")=rs("clkj_admin")
                response.cookies("userpas")("upas")=rs("clkj_password")
				response.cookies("username").expires=now()+1
				'session.Timeout=60
			End IF
		Response.Redirect("nimda_admin.asp")
		Rs.close
		Set Rs=nothing
End IF

'安全退出
IF request.querystring("class")="Out" Then
	session("username") = ""
	session("userpas") = ""
	response.cookies("username")("uname")=""
	response.cookies("userpas")("upas")=""	
	response.write "<script language='javascript'>alert('安全退出，返回登陆页面');top.location.href='!Index.html';</script>"
End IF


'''''产品分类管理
Sub FenLei()
 set rs=server.createobject("adodb.recordset")
		exec="select * from clkj_BigClass order by big_paixu,clkj_BigClassID asc"		
		rs.open exec,conn,1,1
		do while not rs.eof
		big_name=rs("clkj_BigClassName")
		big_id = rs("clkj_BigClassID")
		response.write "<tr bgcolor="&chr(34)&"#ECF6FC"&chr(34)&"><td width="&chr(34)&"5%"&chr(34)&" height="&chr(34)&"25"&chr(34)&" align="&chr(34)&"center"&chr(34)&" valign="&chr(34)&"middle"&chr(34)&"><font color='red'>"&rs("big_paixu")&"</font></td><td width="&chr(34)&"30%"&chr(34)&" height="&chr(34)&"30"&chr(34)&" align="&chr(34)&"left"&chr(34)&" valign="&chr(34)&"middle"&chr(34)&" style="&chr(34)&"padding:2px;"&chr(34)&"><b>"&big_name&"</b></td><td height="&chr(34)&"30"&chr(34)&" align="&chr(34)&"left"&chr(34)&" valign="&chr(34)&"middle"&chr(34)&" style="&chr(34)&"padding:2px;"&chr(34)&"><a href=Nimda_class.asp?clkj_BigClassID="&Rs("clkj_BigClassID")&"&Edit=S_Z#zr><font color='#FF9900'>添加二级分类</font></a> | <a href='Nimda_class.asp?clkj_BigClassID="&Rs("clkj_BigClassID")&"&Edit=B_E#xy'>修改分类</a> | <a href='Nimda_class.asp?clkj_BigClassID="&Rs("clkj_BigClassID")&"&Class=big_Del' onclick="&chr(34)&"return confirm('删除此栏，将会删除此栏目下的所有产品!\n\n是否将此一级栏目删除?');"&chr(34)&"><font color='#FF9900'>分类删除</font></a></td></tr>"
			set rs_1=server.createobject("adodb.recordset")
			exec_1="select * from clkj_SmallClass where clkj_BigClassID="&big_id&" order by small_paixu,clkj_SmallClassID asc"		
			rs_1.open exec_1,conn,1,1
			do while not rs_1.eof
			small_name = rs_1("clkj_SmallClassName")
			response.write "<tr onMouseOver="&chr(34)&"this.style.backgroundColor='#ECF6FC';"&chr(34)&" onmouseout="&chr(34)&"this.style.backgroundColor='#ffffff';"&chr(34)&"><td width="&chr(34)&"5%"&chr(34)&" height="&chr(34)&"25"&chr(34)&" align="&chr(34)&"center"&chr(34)&" valign="&chr(34)&"middle"&chr(34)&">"&rs_1("small_paixu")&"</td><td width="&chr(34)&"30%"&chr(34)&" height="&chr(34)&"30"&chr(34)&" align="&chr(34)&"left"&chr(34)&" valign="&chr(34)&"middle"&chr(34)&" style="&chr(34)&"padding:2px;text-indent:2em;"&chr(34)&">-| "&small_name&"</td><td height="&chr(34)&"30"&chr(34)&" align="&chr(34)&"left"&chr(34)&" valign="&chr(34)&"middle"&chr(34)&" style="&chr(34)&"padding:2px;"&chr(34)&"><a href='nimda_product.asp'><font color='#0000FF'>添加内容</font></a> | <a href='Nimda_class.asp?clkj_SmallClassID="&Rs_1("clkj_SmallClassID")&"&Edit=S_E#xr'>修改二级分类</a> | <a href='Nimda_class.asp?clkj_SmallClassID="&Rs_1("clkj_SmallClassID")&"&Class=small_Del' onclick="&chr(34)&"return confirm('删除此栏，将会删除此栏目下的所有产品!\n\n是否将此二级栏目删除?');"&chr(34)&"><font color='#0000FF'>分类删除</font></a></td></tr>"
			rs_1.movenext			
			loop
			rs_1.close	
		rs.movenext
		loop
		rs.close
End Sub

'''关键词入录''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


IF Request.Querystring("Class")="keyy" Then
	Set Rs=server.createobject("ADODB.Recordset")
	Sql="select * from key where keyname='"&trim(request.Form("key_name"))&"'"
	Rs.open Sql,conn,1,3
  If Not (Rs.Eof and Rs.Bof) or request.Form("key_name")="" Then
response.write "<script language='javascript'>alert('重复添加或链接名为空，请返回查看');history.go(-1);</script>"
	Else
	Rs.Addnew
	Rs("keyname") = trim(Request.Form("key_name"))
	Rs("keytitle") = Request.Form("key_title")
	Rs("keylink") = Request.Form("key_links")
   
	Rs.update
	Response.Redirect "Nimda_key.asp?Lei=增加改成功"
	Rs.close
  End If
Else IF Request.Querystring("Class")="key_Edit" Then
	Set Rs=server.createobject("ADODB.Recordset")
	Sql="select * from key where keyid="&request("keyid")
	Rs.open Sql,conn,1,3
	Rs("keyname") = trim(Request.Form("key_name"))
	Rs("keytitle") = Request.Form("key_title")
	Rs("keylink") = Request.Form("key_links")
	Rs.update
	Response.Redirect "Nimda_key.asp?Lei=修改成功"
	Rs.close
Else If request.QueryString("Class")="key_Del" Then'栏目删除
		Sql="delete * from key where keyid="&request("keyid")  
		conn.execute Sql
		Response.Redirect "Nimda_key.asp?Lei=删除成功"
End If	
		
End If
End If



%>
