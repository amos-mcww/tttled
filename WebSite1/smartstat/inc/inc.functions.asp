<%
theURL="http://" & Request.ServerVariables("http_host") & Getfinddir(Request.ServerVariables("url"))

If request("BackUrl") =" " Then
    BackUrl = Trim(Request.ServerVariables("HTTP_REFERER"))
else
    BackUrl = Trim(request("BackUrl"))
End If

'**************************************************
'��������GetCaptcha
'��  �ã������֤��Ϸ���
'��  �ߣ�����ǿ
'����ֵ��True  ----�Ϸ�
'        False ----���Ϸ�
'**************************************************
Function GetCAPTCHA()
	Dim TestObj
	On Error Resume Next
	Set TestObj = Server.CreateObject("Adodb.Stream")
	Set TestObj = Nothing
	If Err Then
		Dim TempNum
		Randomize timer
		TempNum = cint(8999*Rnd+1000)
		Session("GetCAPTCHA") = TempNum
		GetCAPTCHA = session("GetCAPTCHA")		
	Else
		GetCAPTCHA = "<img src=""inc/captcha.asp"" id=""safecode"" border=""0"" onclick=""reloadcode()"" />"
	End If
End Function

'****************************************************
'��������ErrMsg
'��  �ã���ʾ������ʾ��Ϣ
'��  ������
'****************************************************
sub ErrMsg(Message)
	dim strErr
	strErr=strErr & "<html><head><title>��ʾ��Ϣ</title><meta http-equiv='Content-Type' content='text/html; charset=gb2312'>" & vbcrlf
	strErr=strErr & "<style type='text/css'>" & vbcrlf
	strErr=strErr & "<!--" & vbcrlf
	strErr=strErr & "body {font:12.8px Arial, Helvetica, sans-serif;margin-top: 0px; margin-right: 0px; margin-bottom: 0px; margin-left: 0px;COLOR: #6e6e6e;}td {font: 12px/1.6em Tahoma,Verdana;line-height: 18.8px;}A:link {COLOR: #000000; TEXT-DECORATION: none}A:active {COLOR: #000000;TEXT-DECORATION: underline}A:visited {COLOR: #000000; TEXT-DECORATION: none}A:hover {COLOR: #000000;TEXT-DECORATION: underline}" & vbcrlf
	strErr=strErr & "-->" & vbcrlf
	strErr=strErr & "</style>" & vbcrlf
	strErr=strErr & "</head><body><br><br>" & vbcrlf
	strErr=strErr & "<table width=""50%"" border=""0"" cellpadding=""5"" cellspacing=""1"" bgcolor=""cad9ea"" align=""center"">" & vbcrlf
	strErr=strErr & "<tr>" & vbcrlf
	strErr=strErr & "<td bgcolor=""e8f3fd"">��ʾ��Ϣ</td>" & vbcrlf
	strErr=strErr & "</tr>" & vbcrlf
	strErr=strErr & "<tr>" & vbcrlf
	strErr=strErr & "<td bgcolor=""#FFFFFF"">" & Message &"" & vbcrlf
	strErr=strErr & "<br>" & vbcrlf
	strErr=strErr & "<a href=""javascript:history.back(1)"">������ﷵ��</a>" & vbcrlf
	strErr=strErr & "</td>" & vbcrlf
	strErr=strErr & "</tr>" & vbcrlf
	strErr=strErr & "</table>" & vbcrlf
	strErr=strErr &  vbCrlf
	strErr=strErr & "</body></html>" & vbcrlf
	response.write strErr
end sub

'****************************************************
'��������performMsg
'��  �ã���ʾ�ɹ���ʾ��Ϣ
'��  ������
'****************************************************
Sub performMsg(Perform, BackUrl)
    Dim strPerform
	strPerform=strPerform & "<html><head><title>��ʾ��Ϣ</title><meta http-equiv='Content-Type' content='text/html; charset=gb2312'>" & vbcrlf
	strPerform=strPerform & "<style type='text/css'>" & vbcrlf
	strPerform=strPerform & "<!--" & vbcrlf
	strPerform=strPerform & "body {font:12.8px Arial, Helvetica, sans-serif;margin-top: 0px; margin-right: 0px; margin-bottom: 0px; margin-left: 0px;COLOR: #6e6e6e;}td {font: 12px/1.6em Tahoma,Verdana;line-height: 18.8px;}A:link {COLOR: #000000; TEXT-DECORATION: none}A:active {COLOR: #000000;TEXT-DECORATION: underline}A:visited {COLOR: #000000; TEXT-DECORATION: none}A:hover {COLOR: #000000;TEXT-DECORATION: underline}" & vbcrlf
	strPerform=strPerform & "-->" & vbcrlf
	strPerform=strPerform & "</style>" & vbcrlf
	strPerform=strPerform & "</head><body><br><br>" & vbcrlf
	strPerform=strPerform & "<table width=""50%"" border=""0"" cellpadding=""5"" cellspacing=""1"" bgcolor=""cad9ea"" align=""center"">" & vbcrlf
	strPerform=strPerform & "<tr>" & vbcrlf
	strPerform=strPerform & "<td bgcolor=""e8f3fd"">��ʾ��Ϣ</td>" & vbcrlf
	strPerform=strPerform & "</tr>" & vbcrlf
	strPerform=strPerform & "<tr>" & vbcrlf
	strPerform=strPerform & "<td bgcolor=""#FFFFFF"">" & Perform &"" & vbcrlf
	strPerform=strPerform & "<br>" & vbcrlf
    If BackUrl <> "" Then
	strPerform = strPerform & "<a href='" & BackUrl & "'>������ﷵ��</a>"
    Else
	strPerform = strPerform & "<a href='" & Trim(Request.ServerVariables("HTTP_REFERER")) & "'>������ﷵ��</a>"
    End If
	strPerform=strPerform & "</td>" & vbcrlf
	strPerform=strPerform & "</tr>" & vbcrlf
	strPerform=strPerform & "</table>" & vbcrlf
	strPerform=strPerform & "</body></html>" & vbcrlf
	response.write strPerform
end Sub

'**************************************************
'��������GetIsAdmin
'��  �ã�����̨�û��Ƿ��¼
'��  ������
'����ֵ��True ----�Ѿ���¼
'        False ---û�е�¼
'**************************************************
function GetIsAdmin()
	Logined=True
	strUserName=session("strUserName")
	strPassword=session("strPassword")

	if strUserName="" then
		Logined=False
	end if
	if strPassword="" then
		Logined=False
	end if

	if Logined=True then
		strUserName=replace(trim(strUserName),"'","")
		strPassword=replace(trim(strPassword),"'","")

		set txtRs = server.createobject("adodb.recordset")
		sql="select * from tblSite where UserName='" & strUserName & "' and UserPassword='" & strPassword &"'"
		'response.write sql
		'response.End
		txtRs.open sql,conn,1,1
		if txtRs.bof and txtRs.eof then
			Logined=False
		else
			if strUserName<>txtRs("UserName") or strPassword<>txtRs("UserPassword") then
				Logined=False
			end if
		end if
		txtRs.close
		set txtRs=nothing
	end if
	GetIsAdmin=Logined
end function

'****************************************************
'��������GetCity
'��  �ã�ʡ��
'****************************************************
Function GetCity(UserProvince, UserCity)
	Dim tmpstr
	tmpstr = "<select onchange=setcity(); name='UserProvince' style='width:170px'>"
	tmpstr = tmpstr & "<option value=''>-��ѡ��ʡ��-</option>"
	tmpstr = tmpstr & "<option "
	tmpstr = tmpstr & "value=����>����</option> <option value=����>����</option> "
	tmpstr = tmpstr & "<option value=����>����</option> <option "
	tmpstr = tmpstr & "value=����>����</option> <option value=����>����</option> "
	tmpstr = tmpstr & "<option value=�㶫>�㶫</option> <option "
	tmpstr = tmpstr & "value=����>����</option> <option value=����>����</option> "
	tmpstr = tmpstr & "<option value=����>����</option> <option "
	tmpstr = tmpstr & "value=�ӱ�>�ӱ�</option> <option value=������>������</option> "
	tmpstr = tmpstr & "<option value=����>����</option> <option "
	tmpstr = tmpstr & "value=���>���</option> <option value=����>����</option> "
	tmpstr = tmpstr & "<option value=����>����</option> <option "
	tmpstr = tmpstr & "value=����>����</option> <option value=����>����</option> "
	tmpstr = tmpstr & "<option value=����>����</option> <option "
	tmpstr = tmpstr & "value=����>����</option> <option value=����>����</option>"
	tmpstr = tmpstr & "<option value=���ɹ�>���ɹ�</option> <option "
	tmpstr = tmpstr & "value=����>����</option> <option value=�ຣ>�ຣ</option> "
	tmpstr = tmpstr & "<option value=ɽ��>ɽ��</option> <option "
	tmpstr = tmpstr & "value=�Ϻ�>�Ϻ�</option> <option value=ɽ��>ɽ��</option> "
	tmpstr = tmpstr & "<option value=����>����</option> <option "
	tmpstr = tmpstr & "value=�Ĵ�>�Ĵ�</option> <option value=̨��>̨��</option> "
	tmpstr = tmpstr & "<option value=���>���</option> <option "
	tmpstr = tmpstr & "value=�½�>�½�</option> <option value=����>����</option> "
	tmpstr = tmpstr & "<option value=����>����</option> <option "
	tmpstr = tmpstr & "value=�㽭>�㽭</option> <option "
	tmpstr = tmpstr & "value=����>����</option></select>"
	tmpstr = tmpstr & " <select name='UserCity' >"
	tmpstr = tmpstr & "</select>"
	tmpstr = tmpstr & "<SCRIPT src=""inc/inc.area.js""></SCRIPT>"
	tmpstr = tmpstr & "<SCRIPT>initprovcity('" & UserProvince & "','" & UserCity & "');</SCRIPT>"
	GetCity = tmpstr
End Function

'**************************************************
'��������strCLng
'��  �ã����ַ�תΪ������ֵ 
'**************************************************
Function strCLng(ByVal str)
    If IsNumeric(str) Then
        strCLng = Fix(CDbl(str))
    Else
        strCLng = 0
    End If
End Function

'**************************************************
'��������GetNumberRound
'��  �ã�����ƽ����
'**************************************************
Function GetNumberRound(ByVal str,ByVal oStr)
    If IsNumeric(str) Then
        str = fix(cdbl(str))
    Else
        str = 0
    End If
    If IsNumeric(oStr) Then
        oStr = fix(cdbl(oStr))
    Else
        oStr = 0
    End If
	If str<=0 or oStr<=0 then
		GetNumberRound=0
	Else
		GetNumberRound=formatnumber(round(str*100/oStr,1),1,true) & "%"
	End if
End Function

'**************************************************
'��������strleach
'��  �ã����˷Ƿ��ַ����� 
'**************************************************
function strLeach(str)'
	dim tempstr 
	if str="" then exit function 
	tempstr=replace(str,chr(34),"")' " 
	tempstr=replace(tempstr,chr(39),"")' ' 
	tempstr=replace(tempstr,chr(60),"")' < 
	tempstr=replace(tempstr,chr(62),"")' > 
	tempstr=replace(tempstr,chr(37),"")' % 
	tempstr=replace(tempstr,chr(38),"")' & 
	tempstr=replace(tempstr,chr(40),"")' ( 
	tempstr=replace(tempstr,chr(41),"")' ) 
	tempstr=replace(tempstr,chr(59),"")' ; 
	tempstr=replace(tempstr,chr(43),"")' + 
	tempstr=replace(tempstr,chr(45),"")' - 
	tempstr=replace(tempstr,chr(91),"")' [ 
	tempstr=replace(tempstr,chr(93),"")' ] 
	tempstr=replace(tempstr,chr(123),"")' { 
	tempstr=replace(tempstr,chr(125),"")' } 
	strleach=tempstr 
end function 

'**************************************************
'��������Getfinddir
'��  �ã��ҵ��ļ���ַ��ȫ·��
'**************************************************
Function Getfinddir(filepath)
	Getfinddir=left(filepath,instrRev(filepath,"/"))
end Function

'**************************************************
'��������GetNow
'��  �ã�ȡ����
'**************************************************
Function GetNow(ByVal str)
	uNow = dateadd("h",Site_MasterTimeZone-smart_ZoneServer,now())
	strToday=datevalue(uNow) '����
	strTomorrow=dateadd("d",+1,strToday) '����
	strYesterday=dateadd("d",-1,strToday) '����
	strday7=dateadd("d",-6,strToday) '���7��
	strday30=dateadd("d",-29,strToday) '���30��
	select case str
	case 0
		GetNow = strToday
	case 1
		GetNow = strYesterday
	case 2
		GetNow = strTomorrow
	case 7
		GetNow = strday7
	case 30
		GetNow = strday30
	end select
End Function

' ********************************************************
'�� �� �� �� �� �� �� �� ��
' ********************************************************

' ����ͻ�����Ϣ
Sub GetFromClient(SNum,SCon)
	set tmpRS=conn.Execute("select * from tblClient where LogClientType=" & SNum & " and LogClientContent='"&SCon&"' and Site_ID=" & SiteID)
	if tmpRS.eof then
		conn.Execute ("insert into tblClient (Site_ID,LogClientType,LogClientContent,LogClientTotal,LogClientYesterday,LogClientToday,LogClientLastTime) Values("&Siteid&","&SNum&",'"&SCon&"',1,0,1,'"&truenow&"')")
	else
	    conn.Execute ("update tblClient set LogClientTotal=LogClientTotal+1,LogClientToday=LogClientToday+1,LogClientLastTime='"&truenow&"'  where LogClientType=" & SNum & " and LogClientContent='"&SCon&"' and Site_ID=" & SiteID)
	end if
	set tmpRS=nothing
end Sub

' ����������Ϣ
sub GetFromDatabase(SNum,SCon,SCon2)
	SCon=replace(SCon,"'","''")
	SCon2=replace(SCon2,"'","''")
	set tmpRS=conn.Execute("select id from tblOriginPage where LogPageType=" & SNum & " and LogPageContent='"&SCon&"' and Site_ID=" & SiteID)
	if tmpRS.eof then
		conn.Execute ("insert into tblOriginPage (Site_ID,LogPageType,LogPageContent,LogPageLastURL,LogPageTotal,LogPageYesterday,LogPageToday,LogPageLastTime) Values("&Siteid&","&SNum&",'"&SCon&"','"&SCon2&"',1,0,1,now()-"&smart_ZoneServer&"/24)")
	else
	    conn.Execute ("update tblOriginPage set LogPageTotal=LogPageTotal+1,LogPageToday=LogPageToday+1,LogPageLastURL='"&SCon2&"',LogPageLastTime=now()-"&smart_ZoneServer&"/24 where LogPageType=" & SNum & " and LogPageContent='"&SCon&"' and Site_ID=" & SiteID)
	end if
	set tmpRS=nothing
end sub

' �ҵ���ǰURL��Ӧ��վ��
function findhost(furl)
	if furl<> "" then
	ffurl		= split(furl,"/")
	findhost	= ffurl(2)
	if left(findhost,8)="192.168." or left(findhost,3)="10." or findhost="127.0.0.1" or instr(findhost,".")=0 then findhost="LAN"
	else 
	findhost	= ""
	end if
end function

' �ҵ��ļ���ַ��ȫ·��
Function finddir(filepath)
	finddir=left(filepath,instrRev(filepath,"/")-1)
end Function

' �ҵ�IP��ַ��Ӧ�ĵ���
function findArea(vIP)
	on error resume next
	dim inIP,inIPnum,inIPs
	inIP=vip
	inIPs=split(inIP,".")
	inIPnum=16777216*inips(0) + 65536*inips(1) + 256*inips(2) + inips(3)
	set connip=server.createobject("adodb.connection")
	connip.Open "driver={Microsoft Access Driver (*.mdb)};dbq=" & Server.MapPath(DBFolder&"IP.mdb")
	set rsip=connip.Execute("select StartIp,EndIp,Address from address where EndIp>="&inipnum&" and StartIp<=" _
		& inipnum & " order by EndIp-StartIp")
	if rsip.eof then
		findArea=""
	else
		findArea=rsip("Address")
		if instr(findarea,"δ֪") then findarea=""
	end if
end function

'���ݵ�������ʡ��
Function GetProvince(strCountry) '�򵥷����ʡ��
	If Len(trim(strCountry)) = 0 Then
	GetProvince = "δ֪"
	Exit Function
	End If
	On Error Resume Next
	Dim strProvince,i
	strProvince= strProvince & "������|�Ϻ���|�����|������|���|����|�㶫ʡ|�ӱ�ʡ|ɽ��ʡ|���ɹ�|����ʡ|����ʡ|������ʡ|����ʡ"
	strProvince= strProvince & "|�㽭ʡ|����ʡ|����ʡ|����ʡ|ɽ��ʡ|����ʡ|����ʡ|����ʡ|����|����ʡ|�Ĵ�ʡ|����ʡ|����ʡ|����|����ʡ|����ʡ|�ຣʡ|����|�½�|̨��ʡ"
	strProvince=split(strProvince,"|")
	for i=0 to ubound(strProvince)
	if instr(strCountry,strProvince(i))>0 then
	   GetProvince=strProvince(i)
	   exit function
	end if
	next
	GetProvince = "δ֪"
End Function

'ʡתƴ��
function GetChinaMap(strLang)
  select case strLang
  case "������"
    outstr="CN.BJ"
  case "�Ϻ���"
    outstr="CN.SH"
  case "�����"
    outstr="CN.TJ"
  case "������"
    outstr="CN.CQ"
  case "���"
    outstr="CN.HK"
  case "����"
    outstr="CN.MA"
  case "�㶫ʡ"
    outstr="CN.GD"
  case "�ӱ�ʡ"
    outstr="CN.HB"
  case "ɽ��ʡ"
    outstr="CN.SX"
  case "���ɹ�"
    outstr="CN.NM"
  case "����ʡ"
    outstr="CN.LN"
  case "����ʡ"
    outstr="CN.JL"
  case "������ʡ"
    outstr="CN.HL"
  case "�㽭ʡ"
    outstr="CN.ZJ"
  case "����ʡ"
    outstr="CN.AH"
  case "����ʡ"
    outstr="CN.FJ"
  case "����ʡ"
    outstr="CN.JX"
  case "ɽ��ʡ"
    outstr="CN.SD"
  case "����ʡ"
    outstr="CN.HE"
  case "����ʡ"
    outstr="CN.HB"
  case "����ʡ"
    outstr="CN.HN"
  case "����"
    outstr="CN.GX"
  case "����ʡ"
    outstr="CN.HA"
  case "�Ĵ�ʡ"
    outstr="CN.SC"
  case "����ʡ"
    outstr="CN.GZ"
  case "����ʡ"
    outstr="CN.YN"
  case "����"
    outstr="CN.XZ"
  case "����ʡ"
    outstr="CN.SA"
  case "����ʡ"
    outstr="CN.GS"
  case "�ຣʡ"
    outstr="CN.QH"
  case "����"
    outstr="CN.NX"
  case "�½�"
    outstr="CN.XJ"
  case "̨��ʡ"
    outstr="CN.TA"
  case else
    outstr="CN.BJ"
  end select
  GetChinaMap=outstr
end function

' д�������ʵ��û�
Sub GetLastUser()
  on error resume next
	CacheName=smart_CacheName & "_Last_" & Siteid
	if isempty(Application(CacheName)) then
		Application(CacheName)=vIP & "#oSmartstat#" & vAgent & "#oSmartstat#" & vPage & "#oSmartstat#" & vComeHost & "#oSmartstat#" & vcome & "#oSmartstat#" & truenow
	else
		onA=split(Application(CacheName),Vbcrlf)
		onAs=ubound(onA)
		strOut=vIP & " " & vArea & "#oSmartstat#" & vAgent & "#oSmartstat#" & vPage & "#oSmartstat#" & vComeHost & "#oSmartstat#" & vcome & "#oSmartstat#" & truenow
		j=1
		for i=0 to onAs step 1
			strOut=strOut & vbcrlf & onA(i)
			j=j+1
			if j>= Site_SaveNum then exit for
		next
		Application.Lock
		Application(CacheName)=strOut
		Application.UnLock 
	end if
end sub

'����Ҫ�����IP
function vsaveips(inips)
	vsaveips=left(inips,len(inips)-1)
	vsaveips=right(vsaveips,len(vsaveips)-1)
	howip=split(vsaveips,"#")
	if ubound(howip) < Site_DelRefresh then
		vsaveips="#" & vsaveips & "#" & vip & "#"
	else
		vsaveips=replace("#" & vsaveips,"#" & howip(0) & "#","#") & "#" & vip & "#"
	end if
end function

'���Ժ���
function GetLangchs(strLang)
  select case strLang
  case "af"
    outstr="�ϷǺ�����"
  case "sq"
    outstr="������������"
  case "ar-ae"
    outstr="�������� - ����������������"
  case "ar-bh"
    outstr="�������� - ����"
  case "ar-dz"
    outstr="�������� - ����������"
  case "ar-eg"
    outstr="�������� - ����"
  case "ar-iq"
    outstr="�������� - ������"
  case "ar-jo"
    outstr="�������� - Լ��"
  case "ar-kw"
    outstr="�������� - ������"
  case "ar-lb"
    outstr="�������� - �����"
  case "ar-ly"
    outstr="�������� - ������"
  case "ar-ma"
    outstr="�������� - Ħ���"
  case "ar-om"
    outstr="�������� - ����"
  case "ar-qa"
    outstr="�������� - ������"
  case "ar-sa"
    outstr="�������� - ɳ�ذ�����"
  case "ar-sy"
    outstr="�������� - ������"
  case "ar-tn"
    outstr="�������� - ͻ��˹"
  case "ar-ye"
    outstr="�������� - Ҳ��"
  case "hy"
    outstr="����������"
  case "az-az"
    outstr="�������� - ����"
  case "az-az"
    outstr="�������� - �������"
  case "eu"
    outstr="��˹����"
  case "be"
    outstr="�׶���˹��"
  case "bg"
    outstr="����������"
  case "ca"
    outstr="��̩��������"
  case "zh"
    outstr="����"
  case "zh-cn"
    outstr="���� - �л����񹲺͹�"
  case "zh-hk"
    outstr="���� - �л����񹲺͹�����ر�������"
  case "zh-mo"
    outstr="���� - �л����񹲺͹������ر�������"
  case "zh-sg"
    outstr="���� - �¼���"
  case "zh-tw"
    outstr="���� - ̨�����"
  case "hr"
    outstr="���޵�����"
  case "cs"
    outstr="�ݿ���"
  case "da"
    outstr="������"
  case "nl"
    outstr="������"
  case "nl-nl"
    outstr="������"
  case "nl-be"
    outstr="������ - ����ʱ"
  case "en"
    outstr="Ӣ��"
  case "en-au"
    outstr="Ӣ�� - �Ĵ�����"
  case "en-bz"
    outstr="Ӣ�� - ������"
  case "en-ca"
    outstr="Ӣ�� - ���ô�"
  case "en-cb"
    outstr="Ӣ�� - ���ձȵ���"
  case "en-ie"
    outstr="Ӣ�� - ������"
  case "en-jm"
    outstr="Ӣ�� - �����"
  case "en-nz"
    outstr="Ӣ�� - ������"
  case "en-ph"
    outstr="Ӣ�� - ���ɱ�"
  case "en-za"
    outstr="Ӣ�� - �Ϸ�"
  case "en-tt"
    outstr="Ӣ�� - ������ﵺ"
  case "en-gb"
    outstr="Ӣ�� - Ӣ��"
  case "en-us"
    outstr="Ӣ�� - ����"
  case "et"
    outstr="��ɳ������"
  case "fa"
    outstr="��˹��"
  case "fi"
    outstr="������"
  case "fo"
    outstr="������"
  case "fr"
    outstr="����"
  case "fr-fr"
    outstr="���� - ����"
  case "fr-be"
    outstr="���� - ����ʱ"
  case "fr-ca"
    outstr="���� - ���ô�"
  case "fr-lu"
    outstr="���� - ¬ɭ��"
  case "fr-ch"
    outstr="���� - ��ʿ"
  case "gd-ie"
    outstr="�Ƕ��� - ������"
  case "gd"
    outstr="�Ƕ��� - �ո���"
  case "de"
    outstr="����"
  case "de-de"
    outstr="���� - �¹�"
  case "de-at"
    outstr="���� - �µ���"
  case "de-li"
    outstr="���� - ��֧��ʿ��"
  case "de-lu"
    outstr="���� - ¬ɭ��"
  case "de-ch"
    outstr="���� - ��ʿ"
  case "el"
    outstr="ϣ����"
  case "he"
    outstr="ϣ������"
  case "hi"
    outstr="ӡ����"
  case "hu"
    outstr="��������"
  case "is"
    outstr="������"
  case "id"
    outstr="ӡ����������"
  case "it"
    outstr="�������"
  case "it-it"
    outstr="������� - �����"
  case "it-ch"
    outstr="������� - ��ʿ"
  case "ja"
    outstr="����"
  case "ko"
    outstr="������"
  case "lv"
    outstr="����ά����"
  case "lt"
    outstr="��������"
  case "mk"
    outstr="FYRO �������"
  case "ms-my"
    outstr="������ - ��������"
  case "ms-bn"
    outstr="������ - ����"
  case "mt"
    outstr="�������"
  case "mr"
    outstr="��������"
  case "no"
    outstr="Ų����"
  case "no-no"
    outstr="Ų���� - �������"
  case "no-no"
    outstr="Ų���� - ��ŵ˹��"
  case "pl"
    outstr="������"
  case "pt"
    outstr="��������"
  case "pt-pt"
    outstr="�������� - ������"
  case "pt-br"
    outstr="�������� - ����"
  case "rm"
    outstr="����-������"
  case "ro"
    outstr="����������"
  case "ro-mo"
    outstr="���������� - Ħ������"
  case "ru"
    outstr="����"
  case "ru-mo"
    outstr="���� - Ħ������"
  case "sa"
    outstr="����"
  case "sr"
    outstr="����ά����"
  case "sr-sp"
    outstr="����ά���� - �������"
  case "sr-sp"
    outstr="����ά���� - ����"
  case "tn"
    outstr="��������"
  case "sl"
    outstr="˹����������"
  case "sk"
    outstr="˹�工����"
  case "sb"
    outstr="������"
  case "es"
    outstr="��������"
  case "es-es"
    outstr="�������� - ������"
  case "es-ar"
    outstr="�������� - ����͢"
  case "es-bo"
    outstr="�������� - ����ά��"
  case "es-cl"
    outstr="�������� - ����"
  case "es-co"
    outstr="�������� - ���ױ���"
  case "es-cr"
    outstr="�������� - ��˹�����"
  case "es-do"
    outstr="�������� - ������ӹ��͹�"
  case "es-ec"
    outstr="�������� - ��϶��"
  case "es-gt"
    outstr="�������� - Σ������"
  case "es-hn"
    outstr="�������� - �鶼��˹"
  case "es-mx"
    outstr="�������� - ī����"
  case "es-ni"
    outstr="�������� - �������"
  case "es-pa"
    outstr="�������� - ������"
  case "es-pe"
    outstr="�������� - ��³"
  case "es-pr"
    outstr="�������� - �������"
  case "es-py"
    outstr="�������� - ������"
  case "es-sv"
    outstr="�������� - �����߶�"
  case "es-uy"
    outstr="�������� - ������"
  case "es-ve"
    outstr="�������� - ί������"
  case "sx"
    outstr="��ͼ��"
  case "sw"
    outstr="˹��ϣ����"
  case "sv"
    outstr="�����"
  case "sv-se"
    outstr="�����"
  case "sv-fi"
    outstr="����� - ����"
  case "ta"
    outstr="̩�׶���"
  case "tt"
    outstr="������"
  case "th"
    outstr="̩��"
  case "tr"
    outstr="��������"
  case "ts"
    outstr="������"
  case "uk"
    outstr="�ڿ�����"
  case "ur"
    outstr="�ڶ����� - �ͻ�˹̹"
  case "uz-uz"
    outstr="���ȱ���� - �����"
  case "uz-uz"
    outstr="���ȱ���� - ����"
  case "vi"
    outstr="Խ����"
  case "xh"
    outstr="������"
  case "yi"
    outstr="�������"
  case "zu"
    outstr="��³��"
  case else
    outstr="δ֪"
  end select
  GetLangchs=outstr & " (" & strLang & ")"
end function

'��������
dim array_Search(6,1)
array_Search(0,0) = "google."
array_Search(0,1) = "Google"

array_Search(1,0) = "baidu."
array_Search(1,1) = "�ٶ�"

array_Search(2,0) = "youdao."
array_Search(2,1) = "�е�"

array_Search(3,0) = "sogou."
array_Search(3,1) = "�ѹ�"

array_Search(4,0)	= "bing."
array_Search(4,1)	= "��Ӧ"

array_Search(5,0)	= "soso."
array_Search(5,1)	= "SOSO"

array_Search(6,0)	= "yahoo."
array_Search(6,1)	= "�Ż�"

'**************************************************
'��������referrerParsing
'��  �ã�ȡ�����������Ƽ��ؼ���
'**************************************************
sub referrerParsing(referrer, searchEngine, searchKeywords)
	If instr(referrer,"google")<>0 then
		searchEngine		="�ȸ�"   
		pStartingPosition	=instr(referrer,"q=")
		if pStartingPosition>0 then  	
			pStartingPosition=pStartingPosition+2
			pEndingPosition=instr(pStartingPosition+1,referrer,"&")
			if pEndingPosition=0 then
				searchKeywords=decodeURI(mid(referrer,pStartingPosition))	  	  
			Else
				searchKeywords=decodeURI(mid(referrer,pStartingPosition,pEndingPosition-pStartingPosition))
			End if
		End if
	End if
	
	If instr(referrer,"baidu")<>0 then
		searchEngine		="�ٶ�"   
		pStartingPosition	=instr(referrer,"wd=")
		if pStartingPosition>0 then  	
			pStartingPosition=pStartingPosition+3
			pEndingPosition=instr(pStartingPosition+1,referrer,"&")
			if pEndingPosition=0 then
				searchKeywords=decodeURI(mid(referrer,pStartingPosition))  	  	  
			Else
				searchKeywords=decodeURI(mid(referrer,pStartingPosition,pEndingPosition-pStartingPosition))  	  	  
			End if
		End if
	End if
	
	If instr(referrer,"baidu")<>0 then
		searchEngine		="�ٶ�"   
		pStartingPosition	=instr(referrer,"word=")
		if pStartingPosition>0 then  	
			pStartingPosition=pStartingPosition+5
			pEndingPosition=instr(pStartingPosition+1,referrer,"&")
			if pEndingPosition=0 then
				searchKeywords=decodeURI(mid(referrer,pStartingPosition))  	  	  
			Else
				searchKeywords=decodeURI(mid(referrer,pStartingPosition,pEndingPosition-pStartingPosition))  	  	  
			End if
		End if
	End if
	
	If instr(referrer,"youdao")<>0 then
		searchEngine		="�е�"   
		pStartingPosition	=instr(referrer,"q=")
		if pStartingPosition>0 then  	
			pStartingPosition=pStartingPosition+2
			pEndingPosition=instr(pStartingPosition+1,referrer,"&")
			if pEndingPosition=0 then
				searchKeywords=decodeURI(mid(referrer,pStartingPosition))	  	  
			Else
				searchKeywords=decodeURI(mid(referrer,pStartingPosition,pEndingPosition-pStartingPosition))
			End if
		End if
	End if
	
	If instr(referrer,"sogou")<>0 then
		searchEngine		="�ѹ�"   
		pStartingPosition	=instr(referrer,"query=")
		if pStartingPosition>0 then  	
			pStartingPosition=pStartingPosition+6
			pEndingPosition=instr(pStartingPosition+1,referrer,"&")
			if pEndingPosition=0 then
				searchKeywords=decodeURI(mid(referrer,pStartingPosition))	  	  
			Else
				searchKeywords=decodeURI(mid(referrer,pStartingPosition,pEndingPosition-pStartingPosition))
			End if
		End if
	End if
	
	If instr(referrer,"soso")<>0 then
		searchEngine		="����"   
		pStartingPosition	=instr(referrer,"w=")
		if pStartingPosition>0 then  	
			pStartingPosition=pStartingPosition+2
			pEndingPosition=instr(pStartingPosition+1,referrer,"&")
			if pEndingPosition=0 then
				searchKeywords=decodeURI(mid(referrer,pStartingPosition))	  	  
			Else
				searchKeywords=decodeURI(mid(referrer,pStartingPosition,pEndingPosition-pStartingPosition))
			End if
		End if
	End if

	If instr(referrer,"bing")<>0 then
		searchEngine		="��Ӧ"   
		pStartingPosition	=instr(referrer,"q=")
		if pStartingPosition>0 then  	
			pStartingPosition=pStartingPosition+2
			pEndingPosition=instr(pStartingPosition+1,referrer,"&")
			if pEndingPosition=0 then
				searchKeywords=decodeURI(mid(referrer,pStartingPosition))	  	  
			Else
				searchKeywords=decodeURI(mid(referrer,pStartingPosition,pEndingPosition-pStartingPosition))
			End if
		End if
	End if
	
	If instr(referrer,"yahoo")<>0 then
		searchEngine		="�Ż�"   
		pStartingPosition	=instr(referrer,"p=")
		if pStartingPosition>0 then  	
			pStartingPosition=pStartingPosition+2
			pEndingPosition=instr(pStartingPosition+1,referrer,"&")
			if pEndingPosition=0 then
				searchKeywords=decodeURI(mid(referrer,pStartingPosition))	  	  
			Else
				searchKeywords=decodeURI(mid(referrer,pStartingPosition,pEndingPosition-pStartingPosition))
			End if
		End if
	End if
End sub

'����ת��
Function DecodeURI(s) 
	s = UnEscape(s) 
	Dim reg, cs 
	cs = "GBK" 
	Set reg = New RegExp 
	reg.Pattern = "^(?:[\x00-\x7f]|[\xfc-\xff][\x80-\xbf]{5}|[\xf8-\xfb][\x80-\xbf]{4}|[\xf0-\xf7][\x80-\xbf]{3}|[\xe0-\xef][\x80-\xbf]{2}|[\xc0-\xdf][\x80-\xbf])+$" 
	If reg.Test(s) Then cs = "UTF-8" 
	Set reg = Nothing 
	Dim sm 
	Set sm = CreateObject("ADODB.Stream") 
	With sm 
	.Type = 2 
	.Mode = 3 
	.Open 
	.CharSet = "iso-8859-1" 
	.WriteText s 
	.Position = 0 
	.CharSet = cs 
	DecodeURI = .ReadText(-1) 
	.Close 
	End With 
	Set sm = Nothing 
End Function
%>