<!--#include file="../Clkj_Inc/inc.asp"-->

<%	
Dim TemplateUrl,Template  
TemplateUrl=Request.QueryString("MoBanurl")
	
Template="../Clkj_Template/"&TemplateUrl&"/Clkj_Include/Trade_index.asp,../Clkj_Template/"&TemplateUrl&"/Clkj_Include/Trade_Product.asp,../Clkj_Template/"&TemplateUrl&"/Clkj_Include/Trade_P_view.asp,../Clkj_Template/"&TemplateUrl&"/Clkj_Include/Trade_About_Us.asp,../Clkj_Template/"&TemplateUrl&"/Clkj_Include/Trade_News.asp,../Clkj_Template/"&TemplateUrl&"/Clkj_Include/Trade_N_View.asp,../Clkj_Template/"&TemplateUrl&"/Clkj_Include/Trade_Download.asp,../Clkj_Template/"&TemplateUrl&"/Clkj_Include/Trade_Contact.asp,../Clkj_Template/"&TemplateUrl&"/Clkj_Include/Trade_Feedback.asp,../Clkj_Template/"&TemplateUrl&"/Clkj_Include/Trade_Faq.asp,../Clkj_Template/"&TemplateUrl&"/Clkj_Include/Trade_Inc.asp,../Clkj_Template/"&TemplateUrl&"/Clkj_Include/Trade_message.asp,../Clkj_Template/"&TemplateUrl&"/Clkj_Include/Trade_Search.asp"
  

       Template=Split(Template,",")
		For Each Strss In Template
		
			str=ReadFromUTF(Strss,"utf-8")					
	
			content=Replace(str,"$TemplateUrl$",TemplateUrl)
				adress=replace(Strss,"Clkj_Template/"&TemplateUrl&"/Clkj_Include/Trade_","")'获取文件名
				
				IF adress="../Inc.asp" Then
					Filen="../Clkj_Inc/Clkj_Inc.asp"			
				Else IF adress="../message.asp" Then				    
					Filen="../Clkj_Inc/Clkj_message.asp"
				Else
					Filen=adress
				End IF
				End IF
				
			Call WriteToUTF(content,Filen)
		Next
		
		'模板标记写入库
		set rs=server.createobject("adodb.recordset")
		exec="select * from clkj_config where clkj_config_id=1"
		rs.open exec,conn,1,3
		rs("mb")=TemplateUrl
		rs.update
		rs.close			
		response.Redirect "nimda_Templates.asp?lei=ok"
%>

