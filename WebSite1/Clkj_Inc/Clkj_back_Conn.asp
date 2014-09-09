<%

'----------------Sqlע-----------------------------------------------------------------------------------------
Dim clkj_js,clkj_dui,clkj_i
	clkj_js=lcase(request.servervariables("query_string"))
Dim deStr(17)
	deStr(0)="net user"
	deStr(1)="xp_cmdshell"
	deStr(2)="/add"
	deStr(3)="exec%20master.dbo.xp_cmdshell"
	deStr(4)="net localgroup administrators"
	deStr(5)="select"
	deStr(6)="count"
	deStr(7)="asc"
	deStr(8)="char"
	deStr(9)="mid"
	deStr(10)="'"
	deStr(11)=":"
	deStr(12)=""""
	deStr(13)="insert"
	deStr(14)="delete"
	deStr(15)="drop"
	deStr(16)="truncate"
	deStr(17)="from"
	clkj_dui=false
For clkj_i= 0 to ubound(deStr)
	IF instr(clkj_js,deStr(clkj_i))<>0 then
	clkj_dui=true
	end IF
Next

IF clkj_dui Then
	Response.Write("Excuse me:take a break!")
	response.end
end if



URl=Request.ServerVariables("Server_name")

'数据库链接--------------------------------------------------------------------------------------------------------
Dim conn,connstr,Clkj_mdb
	Clkj_mdb="../"&Datasmdb'数据库链接路径
on error resume next
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "driver={Microsoft Access Driver (*.mdb)};dbq="&Server.MapPath(""&Clkj_mdb&"")
		if err then
			err.clear
			set conn=nothing
			response.write "数据库链接出错,请检查数据路径"
			response.End		 
		End IF



'读取水印图片参数-----------------------------------------------

	Set Rs=server.createobject("adodb.recordset")
	exec="select top 1 * from clkj_config"
	Rs.open exec,conn,1,1
		ImgText=rs("pictext")'水印文字
		ImgPic=rs("picimg")'水印图片地址
		Imgop=cint(rs("picof"))'水印图片是否开启
		ImgWidth=cint(rs("clkj_config_sltk"))'缩略图宽度
		Imgyanse="&H"&rs("yanse")
		Imgjiaodu=rs("jiaodu")
		WebUrl=rs("clkj_config_url")'网站网址
	Rs.close



'自动生成(大图片与小图片文件夹)-----------------------------------------------------------------

    Set fso=Server.CreateObject("Scripting.FileSystemObject")
	
		IF Not fso.FolderExists(Server.MapPath("../Clkj_Images/upfile/Bigpic/")) Then
			   fso.CreateFolder(Server.MapPath("../Clkj_Images/upfile/Bigpic/"))
		Else IF Not fso.FolderExists(Server.MapPath("../Clkj_Images/upfile/Smallpic/")) Then
				fso.CreateFolder(Server.MapPath("../Clkj_Images/upfile/Smallpic/"))				
		End IF
		End IF
		

'删除图片调用------------------------------------------------(新增2010.10.22)-

Function DelImage(p,id)
          IF p="hone_Del" Then
				set rs=server.createobject("adodb.recordset")
				exec="select * from clkj_Honor where clkj_ryid="&id
				rs.open exec,conn,1,1
				Dimg="../"&rs("clkj_ryImg")
				rs.close

		   Else IF p="P_Del" Then
		 		set rs=server.createobject("adodb.recordset")
				exec="select * from clkj_Products where clkj_prid="&id
				rs.open exec,conn,1,1
						           
				Productsdel=Split(rs("clkj_prpic"),",")'逐一删除图片
				For Each PicStrssdel In Productsdel
					IF PicStrssdel<>" " and PicStrssdel<>""  Then				
							IF isobjinstalled("Persits.Jpeg") Then
								Dimg="../"&Trim(Replace(PicStrssdel,"upfile/","upfile/Bigpic/"))
								simg="../"&Trim(replace(PicStrssdel,"upfile/","upfile/Smallpic/"))		
								Set fso1 = CreateObject("Scripting.FileSystemObject")
								Set fso3 = fso.getfile(server.mappath(""&simg&""))
								Set fso4 = fso.getfile(server.mappath(""&Dimg&""))
								fso3.delete
								fso4.delete								
							Else
								Dimg="../"&Trim(PicStrssdel)
								Set fso5 = CreateObject("Scripting.FileSystemObject")
								Set fso6 = fso.getfile(server.mappath(""&Dimg&""))
								fso6.delete
								rs.close
							End IF
					End IF
				Next
			Else
				set rs=server.createobject("adodb.recordset")
				exec="select * from Clkj_adv where adv_id="&id
				rs.open exec,conn,1,1
				Dimg="../"&rs("adv_pic")
				rs.close
				
			Set fso = CreateObject("Scripting.FileSystemObject")
			Set fso2 = fso.getfile(server.mappath(""&Dimg&""))
			fso2.delete
			fso.close
			fso2.close
			set fso=nothing
			set fso2=nothing
			End IF
			End IF

End Function


'图片打水印大图小图---------------------------------------------(2010.11.03)------------------------------------

 
	Function TradeCmsJpeg(Jpegclass,ImageUrl,LogoUrl,Textcontent,ImageW,Imagyanse,Imagejiaodu)
	
		Dim Jpeg,Path,Logo,LogoPath,SmallJpeg,SmallPath,BigImg,SmallImg,picpath
		Set fso = CreateObject("Scripting.FileSystemObject")
		picpath=fso.FileExists(Server.MapPath(""&ImageUrl&""))
		
		IF picpath<>""then '控制空路径
				BigImg = Replace(ImageUrl,"upfile/","upfile/Bigpic/")
				SmallImg = Replace(ImageUrl,"upfile/","upfile/Smallpic/")
			
			'图片生成原图等比例缩放(小图生成)-------------------
						
			Set SmallJpeg = Server.CreateObject("Persits.Jpeg")
				SmallPath = Server.MapPath(""&ImageUrl&"")
				SmallJpeg.Open SmallPath
					
					SmallJpeg.Width = ImageW
					SmallJpeg.Height = SmallJpeg.Width*SmallJpeg.OriginalHeight/SmallJpeg.OriginalWidth-1
					SmallJpeg.Quality=100
					SmallJpeg.Save Server.MapPath(""&SmallImg&"")'小图保存路径
			Set SmallJpeg = Nothing
			
			
			'图片水印-----------------------------------------------------------------------------------
			
			Set Jpeg = Server.CreateObject("Persits.Jpeg")
				Path = Server.MapPath(""&ImageUrl&"")
				Jpeg.Open Path
					
					IF Jpegclass=1 Then
					
					   '打上图片水印---------------------------------------------------------------------
					   
						Set Logo  = Server.CreateObject("Persits.Jpeg") 
							LogoPath  =  Server.MapPath( ""&LogoUrl&"") '在这里修改水印图片所在的路径 
							'Logo.RegKey = "48958-77556-02411" 
							Logo.Open LogoPath '重新定义水印图片的大小
							    
								Logo.Width = 150 '更改水印图片的宽度
								if  Jpeg.width> Logo.Width then
									Logo.Height = Logo.Width * Logo.OriginalHeight/Logo.OriginalWidth '按照原先的长宽比计算新的水印高度
									
									'--位置----
									 IF Imagejiaodu=1 then				
									 	Jpeg.DrawImage 10,10,Logo,0.8 '将水印放置左上角
									 Else IF Imagejiaodu=2 then
									 	Jpeg.DrawImage Jpeg.width-Logo.Width-10,10,Logo,0.8 '将水印放置右上角
									 Else IF Imagejiaodu=3 then
									 	Jpeg.DrawImage Jpeg.width-Logo.Width-10,Jpeg.height-Logo.Height-10,Logo,0.8 '将水印放置右下角
									 Else IF Imagejiaodu=4 then
									 	Jpeg.DrawImage 10,Jpeg.height-Logo.Height-10,Logo,0.8 '将水印放置左下角
									 Else
									 	Jpeg.DrawImage Jpeg.width/2-Logo.Width/2,Jpeg.height/2-Logo.Height/2,Logo,0.8 '将水印放置置于中间
									 End IF
									 End IF
									 End IF
									 End IF
									
									'----位置--
								end if
								Jpeg.Save Server.MapPath(""&BigImg&"")'大图片水印保存路径
						Set Logo = Nothing
					
					Else
			
					   '添加文字水印--------------------------------------------------------------------
					    Jpeg.Canvas.Font.Rotation=0'角度
						Jpeg.Canvas.Font.Color =Imagyanse'颜色
						Jpeg.Canvas.Font.Family = "arial" '字体
						Jpeg.Canvas.Font.Bold = True '是否加粗 
						Jpeg.Canvas.Font.Size = 30'字体大小
						jpeg.canvas.font.quality = 4 
						
						TitleWidth=Jpeg.Canvas.GetTextExtent(""&Textcontent&"")
						if Jpeg.Width>TitleWidth then
						
							IF Imagejiaodu=1 then	
								Jpeg.Canvas.Print 10, 10, ""&Textcontent&"" '左上角
							Else IF Imagejiaodu=2 then
								Jpeg.Canvas.Print Jpeg.Width-TitleWidth-10,10,""&Textcontent&""'文字打右上角
							Else IF Imagejiaodu=3 then
								Jpeg.Canvas.Print Jpeg.Width-TitleWidth-10,Jpeg.Height-40,""&Textcontent&""'文字打右下角
							Else IF Imagejiaodu=4 then	
								Jpeg.Canvas.Print 10,Jpeg.Height-40,""&Textcontent&""'文字打左下角
							Else 
								Jpeg.Canvas.Print Jpeg.Width/2-TitleWidth/2,Jpeg.Height/2,""&Textcontent&""'文字打中间
							End IF
							End IF
							End IF
							End IF
							
						end if
						jpeg.Quality=100
						Jpeg.Save Server.MapPath(""&BigImg&"")'大图片水印保存路径
						
					End IF
						
				Set Jpeg = Nothing
				
				'删除临时图片
				Set fso = CreateObject("Scripting.FileSystemObject")
				Set fso2 = fso.getfile(server.mappath(""&ImageUrl&""))
					fso2.delete	
		End IF
	End Function


	
'--模板转换---------------------------------------------------------
	
	Function ReadFromUTF(TempString,CharSet)'读取模板内容
	Dim str
	Set stm=server.CreateObject("adodb.stream")
		stm.Type=2
		stm.Mode=3
		stm.Charset=CharSet
		stm.Open
		stm.loadfromfile Server.MapPath(TempString)
		str=stm.readtext
		stm.Close
		Set stm=Nothing
		ReadFromUTF=str
End Function

Function WriteToUTF(content,Filen)'替换模板内容
Set objStream=Server.CreateObject("ADODB.Stream")
    With objStream
    .Open
    .Charset="utf-8"
    .Position=objStream.Size
    .WriteText=content
    .SaveToFile server.mappath(Filen),2
    .Close
    End With
Set objStream=Nothing
End Function

'修改图片时删除-----------------------------

Function Delpic(id,picurl,rurl)

	Set Rs=server.createobject("ADODB.Recordset")
		Sql="select * from clkj_Products where clkj_prid="&id
		Rs.open Sql,conn,1,3
			Rs("clkj_prpic") = Replace(Rs("clkj_prpic"),picurl,"")'更新图片字段
		Rs.update
			spicurl=Replace(picurl,"upfile/","upfile/smallpic/")
			bpicurl=Replace(picurl,"upfile/","upfile/Bigpic/")
		Set fsoo = CreateObject("Scripting.FileSystemObject")
		Set fsoo2 = fsoo.getfile(server.mappath("../"&spicurl&""))'删除小图
			fsoo2.delete
		set fsoo3= fsoo.getfile(server.mappath("../"&bpicurl&""))'删除大图
			fsoo3.delete
        	Response.Redirect rurl
End Function


'删除分类时调用图片删除共通
Function DelFenLeiImages(Delimg)
			
				Productsdel=Split(Delimg,",")'逐一删除图片
				For Each PicStrssdel In Productsdel
					IF PicStrssdel<>" " and PicStrssdel<>""  Then				
							IF isobjinstalled("Persits.Jpeg") Then
								Dimg="../"&Trim(Replace(PicStrssdel,"upfile/","upfile/Bigpic/"))
								simg="../"&Trim(replace(PicStrssdel,"upfile/","upfile/Smallpic/"))		
								Set fso1 = CreateObject("Scripting.FileSystemObject")
								Set fso3 = fso.getfile(server.mappath(""&simg&""))
								Set fso4 = fso.getfile(server.mappath(""&Dimg&""))
								fso3.delete
								fso4.delete								
							Else
								Dimg="../"&Trim(PicStrssdel)
								Set fso5 = CreateObject("Scripting.FileSystemObject")
								Set fso6 = fso5.getfile(server.mappath(""&Dimg&""))
								fso6.delete
								rs.close
							End IF
					End IF
				Next					
						
End Function

'获取ip

Private Function getIP() 
	Dim strIPAddr 
	If Request.ServerVariables("HTTP_X_FORWARDED_FOR") = "" OR InStr(Request.ServerVariables("HTTP_X_FORWARDED_FOR"), "unknown") > 0 Then 
	strIPAddr = Request.ServerVariables("REMOTE_ADDR") 
	ElseIf InStr(Request.ServerVariables("HTTP_X_FORWARDED_FOR"), ",") > 0 Then 
	strIPAddr = Mid(Request.ServerVariables("HTTP_X_FORWARDED_FOR"), 1, InStr(Request.ServerVariables("HTTP_X_FORWARDED_FOR"), ",")-1) 
	ElseIf InStr(Request.ServerVariables("HTTP_X_FORWARDED_FOR"), ";") > 0 Then 
	strIPAddr = Mid(Request.ServerVariables("HTTP_X_FORWARDED_FOR"), 1, InStr(Request.ServerVariables("HTTP_X_FORWARDED_FOR"), ";")-1) 
	Else 
	strIPAddr = Request.ServerVariables("HTTP_X_FORWARDED_FOR") 
	End If 
getIP = Trim(Mid(strIPAddr, 1, 30)) 
End Function

'生成google sitemap
IF request.QueryString("xml")="xml" Then

	Dim sSQL,rs,sCrLf,sXmlClear,sRssHead,sRssEnd

	sCrLf = chr(13)&chr(10)
	sXmlClear = "<?xml version="&chr(34)&"1.0"&chr(34)&" encoding="&chr(34)&"utf-8"&chr(34)&"?>"&sCrLf 
	sRssHead = "<urlset xmlns="&chr(34)&"http://www.sitemaps.org/schemas/sitemap/0.9"&chr(34)&">"&sCrLf 
	sRssEnd = "</urlset>"
	
			Set rs = Server.CreateObject("ADODB.Recordset") 
			sSQL="select * from clkj_Products order by clkj_prid desc" 
			rs.Open sSQL,Conn,1,1
					do while not rs.eof													
						listinfo=listinfo &"<url>"&sCrLf 
						listinfo=listinfo &"<loc>"&WebUrl&"P_view.asp?pid="&rs("clkj_prid")&"</loc>"&sCrLf 
						'listinfo=listinfo &"<lastmod>"&rs("clkj_time")&"</lastmod>"&sCrLf
						listinfo=listinfo &"<changefreq>always</changefreq>"&sCrLf
						listinfo=listinfo &"<priority>1.0</priority>"&sCrLf
						listinfo=listinfo &"</url>"&sCrLf				
					rs.movenext 
					loop 		
			rs.close 
			set rs=nothing
	
	body=sXmlClear&sRssHead&listinfo&sRssEnd 
	
	writexml "sitemap.xml",body
	
	response.write"<script language='javascript'>alert('SiteMap成功生成!');history.go(-1);</script>"

End IF

Function writexml(filename,bodytext) 
	Set fso = Server.CreateObject("Scripting.FileSystemObject") 
	Set fout = fso.CreateTextFile(Server.MapPath("../"&filename),true)
	fout.WriteLine bodytext 
	fout.close 
End Function

%>