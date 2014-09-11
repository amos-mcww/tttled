<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
Server.ScriptTimeOut=20000000

%>
<%Response.Charset= "utf-8"%>
<!--#include file="upload.inc"-->

<%
	Set Upload = New UpFile_Class						
	Upload.InceptFileType = "gif,jpg,bmp,jpeg,png,rar"		
	Upload.MaxSize = 102400000
	Upload.GetDate()
	If Upload.Err > 0 Then
		Select Case Upload.Err
			Case 1 : Response.Write "请先选择你要上传的文件　[ <a href=# onclick=history.go(-1)>重新上传</a> ]"
			Case 2 : Response.Write "图片大小超过了限制 "&Dvbbs.Forum_Setting(56)&"K　[ <a href=# onclick=history.go(-1)>重新上传</a> ]"
			Case 3 : Response.Write "所上传类型不正确　[ <a href=# onclick=history.go(-1)>重新上传</a> ]"
		End Select
	Else
		 FormPath=Upload.Form("filepath")
		 For Each FormName in Upload.file		
			 Set File = Upload.File(FormName)	
			 If File.Filesize<10 Then
		 		Response.Write "请先选择你要上传的图片　[ <a href=# onclick=history.go(-1)>重新上传</a> ]"
	 		End If
			FileExt	= FixName(File.FileExt)
 			If Not ( CheckFileExt(FileExt) and CheckFileType(File.FileType) ) Then
 				Response.Write ""
			End If
			
 			FileName=FormPath&UserFaceName(FileExt)
			
 			If File.FileSize>0 Then   
				File.SaveToFile Server.mappath(FileName)   
				response.write "<script>window.opener.document."&upload.form("FormName")&"."&upload.form("EditName")&".value='"&FileName&"'</script>"
				Response.Write "<script language=""javascript"">window.close();</script>"
 			End If
 			Set File=Nothing
		Next
	End If
	Set Upload=Nothing
	
Private Function CheckFileExt(FileExt)
	Dim ForumUpload,i
	ForumUpload="gif,jpg,bmp,jpeg,png,rar"
	ForumUpload=Split(ForumUpload,",")
	CheckFileExt=False
	For i=0 to UBound(ForumUpload)
		If LCase(FileExt)=Lcase(Trim(ForumUpload(i))) Then
			CheckFileExt=True
			Exit Function
		End If
	Next
End Function

Function FixName(UpFileExt)
	If IsEmpty(UpFileExt) Then Exit Function
	FixName = Lcase(UpFileExt)
	FixName = Replace(FixName,Chr(0),"")
	FixName = Replace(FixName,".","")
	FixName = Replace(FixName,"asp","")
	FixName = Replace(FixName,"asa","")
	FixName = Replace(FixName,"aspx","")
	FixName = Replace(FixName,"cer","")
	FixName = Replace(FixName,"cdx","")
	FixName = Replace(FixName,"htr","")
End Function

Private Function UserFaceName(FileExt)

     Img=trim(upload.form("imgname"))
	 
	 if img="" or len(img)=0 then
		Randomize
		RanNum = Int(90000*rnd)+10000
 		UserFaceName = UserID&Year(now)&Month(now)&Day(now)&Hour(now)&Minute(now)&Second(now)&RanNum&"."&FileExt
	else
		UserFaceName = Img&"."&FileExt
	end if
	
End Function

Private Function CheckFileType(FileType)
	CheckFileType = False
	If Left(Cstr(Lcase(Trim(FileType))),6)="image/" Then CheckFileType = True
End Function
%>
