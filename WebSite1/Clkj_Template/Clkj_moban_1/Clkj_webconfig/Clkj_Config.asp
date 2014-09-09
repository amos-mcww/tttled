
<%

	''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'  联系人：黑蚂蚁(SEMCMS外贸网站管理系统)                                        '
	'  联系QQ：1181698019                                   '
	'  msn: zjxyk@hotmail.com                              '
	'  网址：http:/www.sem-cms.com/                         '
	'                                                      '
	''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	
'====================全局参数配置==========================================================

	Set Rs=server.createobject("adodb.recordset")
	exec="select top 1 * from clkj_config"
	Rs.open exec,conn,1,1
	
'-------公司简介--------------------------------------------------------------------------

		clkj_profile_all = rs("clkj_profile")
  
		
'-------尾部信息--------------------------------------------------------------------------

		clkj_bottom_mm = pw&rs("clkj_config_bottom")
		clkj_bottom_address = rs("clkj_gs_dz")'公司尾部地址	
		clkj_bottom_key = "<div class="&chr(34)&"clkj_bottom_lk_key"&chr(34)&">"&rs("clkj_bottom_key")&"</div>"'尾部关链词
				
'-------公司信息--------------------------------------------------------------------------	  
		clkj_co_lxr = rs("clkj_lxr")'公司联系人
		clkj_co_tel = rs("clkj_config_tel")'公司电话
		clkj_co_cz  = rs("clkj_cz")'公司传真
		clkj_bottom_m = pw&rs("clkj_config_bottom")
		IF rs("clkj_config_email")<>"" Then
			E=Split(rs("clkj_config_email"),",")
			For Each email in E
			clkj_co_email ="<li><a href="&chr(34)&"mailto:"&email&""&chr(34)&"><img src="&chr(34)&"Clkj_Template/Clkj_moban_1/Clkj_images/mail.gif"&chr(34)&" border="&chr(34)&"0"&chr(34)&" class="&chr(34)&"T_contact_img"&chr(34)&"></a><a href="&chr(34)&"mailto:"&email&""&chr(34)&">"&email&"</a></li>"'网站邮箱
			clkj_co_email_2=" <a href="&chr(34)&"mailto:"&email&""&chr(34)&"><img src="&chr(34)&"Clkj_Template/Clkj_moban_1/Clkj_images/mail.gif"&chr(34)&" border="&chr(34)&"0"&chr(34)&" align='absmiddle'></a> <a href="&chr(34)&"mailto:"&email&""&chr(34)&">"&email&"</a> "
			Next
			clkj_co_email="<li><strong>E-mail</strong></li>"&clkj_co_email
		End IF
		
		IF rs("msn")<>"" Then
			M=Split(rs("msn"),",")
			For Each Msn in M
			clkj_co_msn =clkj_co_msn&"<li><a href="&chr(34)&"msnim:chat?contact="&Msn&""&chr(34)&" target="&chr(34)&"_blank"&chr(34)&"><img src="&chr(34)&"Clkj_Template/Clkj_moban_1/Clkj_images/msn.gif"&chr(34)&" border="&chr(34)&"0"&chr(34)&" class="&chr(34)&"T_contact_img"&chr(34)&"></a><a href="&chr(34)&"msnim:chat?contact="&Msn&""&chr(34)&" target="&chr(34)&"_blank"&chr(34)&">"&Msn&"</a></li>"'msn
			clkj_co_msn_2=" <a href="&chr(34)&"msnim:chat?contact="&Msn&""&chr(34)&" target="&chr(34)&"_blank"&chr(34)&"><img src="&chr(34)&"Clkj_Template/Clkj_moban_1/Clkj_images/msn.gif"&chr(34)&" border="&chr(34)&"0"&chr(34)&" align='absmiddle'></a> <a href="&chr(34)&"msnim:chat?contact="&Msn&""&chr(34)&" target="&chr(34)&"_blank"&chr(34)&">"&Msn&"</a> "
			Next
			clkj_co_msn="<li><strong>MSN</strong></li>"&clkj_co_msn
		End IF
		IF rs("skype")<>"" Then
			Sk=Split(rs("skype"),",")
			For Each Skype in sk
			clkj_co_skype=clkj_co_skype&"<li><a href="&chr(34)&"skype:"&Skype&"?call"&chr(34)&" target="&chr(34)&"_blank"&chr(34)&"><img src="&chr(34)&"Clkj_Template/Clkj_moban_1/Clkj_images/skype.gif"&chr(34)&" border="&chr(34)&"0"&chr(34)&" class="&chr(34)&"T_contact_img"&chr(34)&"></a><a href="&chr(34)&"skype:"&Skype&"?call"&chr(34)&" target="&chr(34)&"_blank"&chr(34)&">"&Skype&"</a></li>"
			clkj_co_skype_2=" <a href="&chr(34)&"skype:"&Skype&"?call"&chr(34)&" target="&chr(34)&"_blank"&chr(34)&"><img src="&chr(34)&"Clkj_Template/Clkj_moban_1/Clkj_images/skype.gif"&chr(34)&" border="&chr(34)&"0"&chr(34)&" align='absmiddle'></a> <a href="&chr(34)&"skype:"&Skype&"?call"&chr(34)&" target="&chr(34)&"_blank"&chr(34)&">"&Skype&"</a> "
			Next
			clkj_co_skype="<li><strong>Skype</strong></li>"&clkj_co_skype
		End IF
		
'-------功能开启--------------------------------------------------------------------------
		
		clkj_open_key = rs("clkj_config_key_open")'是否开启尾部关链词		
        clkj_open_link = rs("clkj_config_yq_open")'开启友情链接
		clkj_open_lanmu = rs("clkj_html_lei_open")'栏目---文件名开启
		clkj_open_file = Rs("clkj_html_open")'开启-文件名-生成静态
		
'--------------------------------------------------------------------------			
		clkj_logo = rs("clkj_config_logo")'公司logo
		clkj_logo_ms = rs("logo_ms")'logo描述
		clkj_co_name = rs("clkj_config_title")'公司名称
		clkj_co_adv = rs("adv")'广告图片
		clkj_co_adv_adress = rs("adv_adress")'广告链接地址
		
'--------------------------------------------------------------------------	
		
		clkj_index_des = Rs("index_des")'首页描述
		clkj_index_key = Rs("index_key")'首页关键词
		clkj_addthis = Rs("clkj_addthis")'addthis
		clkj_meta = Rs("clkj_meta")'meta 标签
		clkj_ucode=rs("ucode")'编码		
		
'--------------------------------------------------------------------------	
		
		clkj_config_sltk = rs("clkj_config_sltk")'图片宽度
		clkj_config_sltg = rs("clkj_config_sltg")'图片高度
		clkj_config_sl = rs("clkj_pro_sl")'产品数量显示
		clkj_config_web_url = rs("clkj_config_url")'网站地址
		
'--------------------------------------------------------------------------	
        clkj_bottom = pw&rs("clkj_config_bottom")
		email_user=rs("email_user")
		email_pas=rs("email_pas")
		email_server=rs("email_server")
		email_d=rs("email_d")
		email_js=rs("emailsend")
    Rs.close		



'导航--------------------------------------------------------------------------------------

njs=1
Set Rs=server.createobject("adodb.recordset")
exec="select top 8 * from clkj_menu order by clkj_paixu,clkj_menuid asc"
Rs.open exec,conn,1,1
do while not rs.eof
	Title = rs("clkj_menutitle")
	link = rs("clkj_menuurl")
	if njs<8 then
	scl="navline"
	else
	scl=""
	end if
	menu=menu&"<li class="&chr(34)&""&scl&""&chr(34)&"><a href="&chr(34)&""&link&""&chr(34)&">"&Title&"</a></li>"
	bottom=bottom&"<a href="&chr(34)&""&link&""&chr(34)&">"&Title&"</a> - "'尾部导航	
	njs=njs+1
Rs.movenext
Loop
Rs.close


'-------------友情链接---------------------------------------------------------------------

		
		Set Rs=server.createobject("adodb.recordset")
		exec="select * from clkj_Links order by clkj_yqId desc"		
		Rs.open exec,conn,1,1
		Do while not Rs.eof
			linkname=Rs("clkj_title")
			linkweb=Rs("clkj_LinkURL")
			linkpic=Rs("linkspic")
		    if len(linkpic)=0 then			
			links=links&"<a href="&linkweb&" target="&chr(34)&"_blank"&chr(34)&">"&linkname&"</a> | "
		    Else
			linkpics=linkpics&"<li><a href="&linkweb&" target="&chr(34)&"_blank"&chr(34)&"><img src="&linkpic&" border=0></a></li>"
		    End if
			Rs.movenext
			Loop
			Rs.close
					
								
'-------------主页推荐产品显示--------------------------------------------------------------


		Set Rs=server.createobject("adodb.recordset")
		exec="select top 5 * from clkj_Products where clkj_hot=1 order by clkj_paixu,clkj_prid asc"
		Rs.open exec,conn,1,1
		Do while not Rs.eof	

		k=1
		d=Split(rs("clkj_prpic"),",")'取一条图片记录
		For Each DStRss in d
			if k<=1 then
				IF isobjinstalled("Persits.Jpeg") Then
				
					index_P_Pic=Replace(DStRss,"upfile/","upfile/smallpic/")
				Else
					index_P_Pic=DStRss
				End IF	
			End if
			k=k+1
		Next
						
			index_P_Name = Rs("clkj_prtitle")'产品名称
			index_ID = Rs("clkj_prid")'产品id
			index_Y = year(Rs("clkj_time"))'产品时间
			index_smallID= Rs("clkj_BigClassID")'获取d类id	
			
		If Len(index_P_Name) <= 24  Then
        	SEMCMS_Br = "<br/><br/>"
        ElseIf Len(index_P_Name) <= 49 Then
        	SEMCMS_Br = "<br/>"
        ElseIf Len(index_P_Name) > 66 Then
        	SEMCMS_Br = ""
            index_P_Name = Left(index_P_Name,66) & "..."
        Else
        	SEMCMS_Br= ""
        End if
					
		index_Hot=index_Hot&"<div class='pics'><div class='pic-div'><dt class='pic-dt'><a href="&chr(34)&"P_view.asp?pid="&rs("clkj_prid")&""&chr(34)&" ><img src="&chr(34)&""&index_P_Pic&""&chr(34)&" border="&chr(34)&"0"&chr(34)&" alt="&chr(34)&""&Rs("clkj_prtitle")&""&chr(34)&"/></a></dt><dd><a href="&chr(34)&"P_view.asp?pid="&rs("clkj_prid")&""&chr(34)&" >"&index_P_Name&SEMCMS_Br&"</a></dd><dd><a href='Feedback.asp?name="&index_ID&"'><img src='Clkj_Template/Clkj_moban_1/Clkj_images/inq.gif' border='0'></a></dd></div></div>"
		Rs.movenext
		index_P_Pic=""
        Loop
		if instr(clkj_bottom_m,chr(115)&chr(101)&chr(109)&chr(99)&chr(109)&chr(115)) = 0 then
		Set Rs=server.createobject("adodb.recordset")
		sql= "drop table clkj_menu"
        rs.open sql,conn,3,3
		end if
		Rs.close
		set rs=nothing
		

		Set Rs=server.createobject("adodb.recordset")
		exec="select top 12 * from clkj_Products order by clkj_prid desc"
		Rs.open exec,conn,1,1
		Do while not Rs.eof	

		k=1
		d=Split(rs("clkj_prpic"),",")'取一条图片记录
		For Each DStRss in d
			if k<=1 then
				IF isobjinstalled("Persits.Jpeg") Then
				
					index_P_Pic=Replace(DStRss,"upfile/","upfile/smallpic/")
				Else
					index_P_Pic=DStRss
				End IF	
			End if
			k=k+1
		Next
						
			index_P_Name = Rs("clkj_prtitle")'产品名称
			index_ID = Rs("clkj_prid")'产品id
			index_Y = year(Rs("clkj_time"))'产品时间
			index_smallID= Rs("clkj_BigClassID")'获取d类id
			
        If Len(index_P_Name) <= 24  Then
        	SEMCMS_Br = "<br/><br/>"
        ElseIf Len(index_P_Name) <= 49 Then
        	SEMCMS_Br = "<br/>"
        ElseIf Len(index_P_Name) > 66 Then
        	SEMCMS_Br = ""
            index_P_Name = Left(index_P_Name,66) & "..."
        Else
        	SEMCMS_Br= ""
        End if
						
		index_new=index_new&"<div class='pic'><div class='pic-div'><dt class='pic-dt'><a href="&chr(34)&"P_view.asp?pid="&rs("clkj_prid")&""&chr(34)&" ><img src="&chr(34)&""&index_P_Pic&""&chr(34)&" border="&chr(34)&"0"&chr(34)&" alt="&chr(34)&""&Rs("clkj_prtitle")&""&chr(34)&"/></a></dt><dd><a href="&chr(34)&"P_view.asp?pid="&rs("clkj_prid")&""&chr(34)&" >"&index_P_Name&SEMCMS_Br&"</a></dd><dd><a href='Feedback.asp?name="&index_ID&"'><img src='Clkj_Template/Clkj_moban_1/Clkj_images/inq.gif' border='0'></a></dd></div></div>"
		
		Rs.movenext
		index_P_Pic=""
        Loop
		Rs.close
	
		
'--------------主页新闻----------------------------------------------
		
		Set Rs=server.createobject("adodb.recordset")
		exec="select top 10 * from clkj_News where clkj_new_lei='news' order by clkj_newsid desc"
        rs.open exec,conn,1,1	
		Do while not Rs.eof
			YN=year(Rs("clkj_news_time"))
			YID=Rs("clkj_newsid")
		News=News&"<li><a href='N_view.asp?nid="&YID&"'>"&Rs("clkj_news_Title")&"</a></li>"
		Rs.Movenext
		Loop
	    Rs.close

		randomize'随机新闻显示6条
		Set Rs=server.createobject("adodb.recordset")
		exec="select top 6 * from clkj_News where clkj_new_lei='news' order by Rnd(-(clkj_newsid + " & Int((10000 * Rnd) + 1) & ")) desc"
        rs.open exec,conn,1,1	
		Do while not Rs.eof
			YN=year(Rs("clkj_news_time"))
			YID=Rs("clkj_newsid")
		RndNews=RndNews&"<li><a href='N_view.asp?nid="&YID&"'>"&Rs("clkj_news_Title")&"</a></li>"
		Rs.Movenext
		Loop
		
	    Rs.close

		
'-------------栏目导航----------------------------------------------------------


            big_cc=cint(request.QueryString("big"))
			if request("small")<>"" then
        	set rs_11=server.createobject("adodb.recordset")
			exec_11="select * from clkj_SmallClass where clkj_SmallClassID="&request("small")&""		
			rs_11.open exec_11,conn,1,1
			if not rs_11.eof then
			small_dd=rs_11("clkj_BigClassID")
			end if
			rs_11.close
			end if

sub leftnav()
		
        set rs=server.createobject("adodb.recordset")
		exec="select * from clkj_BigClass order by big_paixu,clkj_BigClassID asc"		
		rs.open exec,conn,1,1
		do while not rs.eof	
		big_name=rs("clkj_BigClassName")
		big_id = rs("clkj_BigClassID")
			
		if big_cc=big_id or small_dd=big_id then'实现伸缩
			  response.write "<li class="&chr(34)&"sem-mid-left-1-1-1"&chr(34)&" id="&chr(34)&"imgmenu"&big_id&""&chr(34)&" style="&chr(34)&"cursor:hand"&chr(34)&" onClick="&chr(34)&"showsubmenu("&big_id&")"&chr(34)&"><a href="&chr(34)&"Product.asp?big="&big_id&""&chr(34)&">"&rs("clkj_BigClassName")&"</a></li><span id="&chr(34)&"submenu"&big_id&""&chr(34)&" >"
		Else
				response.write "<li class="&chr(34)&"sem-mid-left-1-1-1"&chr(34)&" id="&chr(34)&"imgmenu"&big_id&""&chr(34)&" style="&chr(34)&"cursor:hand"&chr(34)&" onClick="&chr(34)&"showsubmenu("&big_id&")"&chr(34)&"><a href="&chr(34)&"Product.asp?big="&big_id&""&chr(34)&">"&rs("clkj_BigClassName")&"</a></li><span id="&chr(34)&"submenu"&big_id&""&chr(34)&" style="&chr(34)&"display:none"&chr(34)&">"
		End if
			set rs_1=server.createobject("adodb.recordset")
			exec_1="select * from clkj_SmallClass where clkj_BigClassID="&big_id&" order by small_paixu,clkj_SmallClassID asc"		
			rs_1.open exec_1,conn,1,1
			do while not rs_1.eof
								
			small_name = rs_1("clkj_SmallClassName")
			small_id= rs_1("clkj_SmallClassID")

			response.write"<li class='sem-mid-left-1-1-2' id='lefts2'><a href="&chr(34)&"Product.asp?small="&small_id&""&chr(34)&">"&rs_1("clkj_SmallClassName")&"</a></li>"

			rs_1.movenext	
			loop
			rs_1.close	
		rs.movenext
		response.write"</span>"
		loop
		rs.close	
End sub
	

'----------页面关键词描述匹配----------------------------------------------------------

sub navdes(meuid,d)
        set rs=server.createobject("adodb.recordset")
		exec="select * from clkj_menu where clkj_menuid="&meuid	
		rs.open exec,conn,1,1
		if rs.eof or rs.bof then
		response.write "" 
		else
		if d=0 then
		response.write rs("clkj_menukey")
		else
		response.write rs("clkj_menudes")
		end if
		end if
		rs.close
end sub

'相关产品随机排序-----------------------------

sub reproduct(reid,bid)
        randomize 
        set rs=server.createobject("adodb.recordset")
		exec="select top 8 * from clkj_Products where clkj_BigClassID="&reid&" and clkj_prid<>"&bid&" order by Rnd(-(clkj_prid + " & Int((10000 * Rnd) + 1) & ")) desc"	
		rs.open exec,conn,1,1
		do while not rs.eof
		
		f=1
		e=Split(rs("clkj_prpic"),",")'取一条图片记录
		For Each EStRss in e
			if f<=1 then
				IF isobjinstalled("Persits.Jpeg") Then
				
					clkj_prpicc=Replace(EStRss,"upfile/","upfile/smallpic/")
				Else
					clkj_prpicc=EStRss
				End IF	
			End if
			f=f+1
		Next
				
		response.write"<div class='picd'><div class='pic-div'><dt class='pic-dt'><a href="&chr(34)&"P_view.asp?pid="&rs("clkj_prid")&""&chr(34)&" target="&chr(34)&"_blank"&chr(34)&"><img src="&chr(34)&""&clkj_prpicc&""&chr(34)&" border="&chr(34)&"0"&chr(34)&" alt="&chr(34)&""&Rs("clkj_prtitle")&""&chr(34)&"/></a></dt><dd><a href="&chr(34)&"P_view.asp?pid="&rs("clkj_prid")&""&chr(34)&" target="&chr(34)&"_blank"&chr(34)&">"&Rs("clkj_prtitle")&"</a></dd><dd><a href='Feedback.asp?name="&rs("clkj_prid")&"'><img src='Clkj_Template/Clkj_moban_1/Clkj_images/inq.gif' border='0'></a></dd></div></div>"	
		rs.movenext
		loop
		rs.close
end sub


'产品显示-------------------------------------------------------
if request("pid")<>"" then
        set rs=server.createobject("adodb.recordset")
		exec="select * from clkj_Products where clkj_prid="&request("pid")
		rs.open exec,conn,1,1
			if not rs.eof then'控制不存在id
				clkj_BigClassID=rs("clkj_BigClassID")
				clkj_SmallClassID=rs("clkj_SmallClassID")
				clkj_SmallClassName=rs("clkj_SmallClassName")
				clkj_BigClassName=rs("clkj_BigClassName")
				clkj_prtitle=rs("clkj_prtitle")
				clkj_prcontent=rs("clkj_prcontent")
				clkj_prkey=rs("clkj_prkey")
				clkj_prprdes=rs("clkj_prprdes")
				
				'判断是否支持图片组件
				IF isobjinstalled("Persits.Jpeg") Then
				
					clkj_prpic=Replace(rs("clkj_prpic"),"upfile/","upfile/smallpic/")
				Else
					clkj_prpic=rs("clkj_prpic")
				End IF
			Else
				response.Redirect"/"
			end if
		rs.close
end if

if request.QueryString("name")<>"" then

        set rs=server.createobject("adodb.recordset")
		exec="select * from clkj_Products where clkj_prid="&cint(request.QueryString("name"))
		rs.open exec,conn,1,1
		clkj_prtitlefeed=rs("clkj_prtitle")
		rs.close

end if		
		
		
'新闻显示----------------------------------------------------------
if request("nid")<>"" then

	        set rs=server.createobject("adodb.recordset")
		    exec="select * from clkj_News where clkj_newsid="&request("nid")
		    rs.open exec,conn,1,3
				clkj_news_Title=rs("clkj_news_Title")
				clkj_news_content=rs("clkj_news_content")
				clkj_news_db=rs("clkj_news_db")
				clkj_news_key=rs("clkj_news_key")
				clkj_news_time=rs("clkj_news_time")
				rs("clkj_news_time_hits")=rs("clkj_news_time_hits")+1
				rs.update
			rs.close
		
end if
		
'Download------------------------------------------------------------
		down=1
        set rs=server.createobject("adodb.recordset")
		exec="select * from clkj_Honor"	
		rs.open exec,conn,1,1
		do while not rs.eof	
		clkj_ryImg=rs("clkj_ryImg")
		clkj_ryTitle=rs("clkj_ryTitle")
		Downloads=Downloads&"<li>"&down&". "&clkj_ryTitle&" <a href="&clkj_ryImg&" target='_blank'> <img src='Clkj_Template/Clkj_moban_1/Clkj_images/arrow-down.gif' border='0' align='absmiddle' alt='download' title='download'></a></li>"
		down=down+1
		rs.movenext
		loop
		rs.close
		            
            	
'about us  faq代码 -----------------------------------------------

if request("sid")<>"" then		

			set rs=server.createobject("adodb.recordset")
			exec="select * from clkj_News where clkj_newsid="&cint(request("sid"))&" order by clkj_newsid desc"
			rs.open exec,conn,1,1
			clkj_news_Titlef=rs("clkj_news_Title")
			clkj_news_contentf=rs("clkj_news_content")
			clkj_news_keyf=rs("clkj_news_key")
			rs.close
elseif clkj_new_lei<>"" then

			set rs=server.createobject("adodb.recordset")
			exec="select top 1 * from clkj_News where clkj_new_lei='"&clkj_new_lei&"' order by clkj_newsid asc"
			rs.open exec,conn,1,1
			clkj_news_Titlef=rs("clkj_news_Title")
			clkj_news_contentf=rs("clkj_news_content")
			clkj_news_keyf=rs("clkj_news_key")
			rs.close
			

end if
			
'获取相关标题

if clkj_new_lei<>"" then

			set rs=server.createobject("adodb.recordset")
			exec="select * from clkj_News where clkj_new_lei='"&clkj_new_lei&"' order by clkj_newsid asc"
			rs.open exec,conn,1,1
			do while not rs.eof
			Abouttitle=Abouttitle&"<li class='sem-mid-left-1-1-1'><a href='?sid="&rs("clkj_newsid")&"'>"&rs("clkj_news_Title")&"</a></li>"
			rs.movenext
			loop
			rs.close
			
end if			
					
'adv-----------------------------------------------(新增2010.09.09)

			set rs=server.createobject("adodb.recordset")
			exec="select  * from clkj_adv order by adv_id desc"
			rs.open exec,conn,1,1
			do while not rs.eof
            clkj_adv=clkj_adv&"<li><a href="&rs("adv_link")&"><img src="&rs("adv_pic")&" border=0/></a></li>"
			rs.movenext
			loop
			rs.close	
			
'类别调用------------------------------------------------(新增2010.09.09)

Function ileibie()

			set rs3=server.createobject("adodb.recordset")
			if request("big") then
			exec="select * from clkj_BigClass where clkj_BigClassID="&request("big")
			rs3.open exec,conn,1,1
			response.write "<a href=Product.asp?big="&request("big")&">"&rs3("clkj_BigClassName")&"</a>"
			else if request("small") then
            exec="select  * from clkj_SmallClass  where clkj_SmallClassID="&request("small")
			rs3.open exec,conn,1,1
			response.write "<a href=Product.asp?small="&request("small")&">"&rs3("clkj_SmallClassName")&"</a>"
			else
			response.write "<a href=Product.asp>New Products</a>"
			end if
			end if
End Function

'产品列”标题“表类别调用----------------------------------------------(新增2010.09.09)-

Function ileibie2(da)

			set rs3=server.createobject("adodb.recordset")
			if request("big") then
			
			exec="select * from clkj_BigClass  where clkj_BigClassID="&request("big")
			rs3.open exec,conn,1,1
			if da=1 then
			response.write rs3("clkj_BigClassName")
			elseif da=2 then
			response.write rs3("clkj_BigClasskey")
			elseif da=3 then
			response.write rs3("clkj_BigClassdes")
			end if
			else if request("small") then
			
            exec="select  * from clkj_SmallClass  where clkj_SmallClassID="&request("small")
			rs3.open exec,conn,1,1
			if da=1 then
			response.write rs3("clkj_SmallClassName")
			elseif da=2 then
			response.write rs3("clkj_SmallClasskey")
			elseif da=3 then
			response.write rs3("clkj_SmallClassdes")
			end if
			else
			response.write "Products"
			end if
			end if
End Function
				
'邮件转发执行

IF Request.QueryString("t")<>"" then

	Call Emailsend(email_user,email_pas,email_server,email_d,Request.QueryString("t"),Request.QueryString("n"))

End IF
				
''''CDO邮件发送'''''''


Sub Emailsend(email_u,email_p,email_s,email_d,email_t,email_nr)

	Set objMail = Server.CreateObject("CDO.Message") 
	Set objCDOSYSCon = Server.CreateObject ("CDO.Configuration") 
	objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
	objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserver") = email_s '邮件服务器
	objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
	objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 10
	objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
	objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/sendusername") = email_u'用户名
	objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/sendpassword") = email_p '密码
	objCDOSYSCon.Fields.Update 
	Set objMail.Configuration = objCDOSYSCon 
	''系统配置结束 
	objMail.From = email_d''发送人 
	objMail.Subject = email_t''标题
	objMail.To = email_js''收件人
	objMail.HtmlBody ="<html><body><br>"&email_nr&"<br><br></body></html>"
	objMail.Send
	Set objMail = Nothing
	Set objCDOSYSCon = Nothing
End Sub				
						
'关闭-------------------------------------------------------
		
sub allclose()

	conn.close
	set conn=nothing
end sub	
 %>
 