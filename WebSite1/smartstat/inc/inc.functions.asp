<%
theURL="http://" & Request.ServerVariables("http_host") & Getfinddir(Request.ServerVariables("url"))

If request("BackUrl") =" " Then
    BackUrl = Trim(Request.ServerVariables("HTTP_REFERER"))
else
    BackUrl = Trim(request("BackUrl"))
End If

'**************************************************
'函数名：GetCaptcha
'作  用：检查验证码合法性
'作  者：王国强
'返回值：True  ----合法
'        False ----不合法
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
'过程名：ErrMsg
'作  用：显示错误提示信息
'参  数：无
'****************************************************
sub ErrMsg(Message)
	dim strErr
	strErr=strErr & "<html><head><title>提示信息</title><meta http-equiv='Content-Type' content='text/html; charset=gb2312'>" & vbcrlf
	strErr=strErr & "<style type='text/css'>" & vbcrlf
	strErr=strErr & "<!--" & vbcrlf
	strErr=strErr & "body {font:12.8px Arial, Helvetica, sans-serif;margin-top: 0px; margin-right: 0px; margin-bottom: 0px; margin-left: 0px;COLOR: #6e6e6e;}td {font: 12px/1.6em Tahoma,Verdana;line-height: 18.8px;}A:link {COLOR: #000000; TEXT-DECORATION: none}A:active {COLOR: #000000;TEXT-DECORATION: underline}A:visited {COLOR: #000000; TEXT-DECORATION: none}A:hover {COLOR: #000000;TEXT-DECORATION: underline}" & vbcrlf
	strErr=strErr & "-->" & vbcrlf
	strErr=strErr & "</style>" & vbcrlf
	strErr=strErr & "</head><body><br><br>" & vbcrlf
	strErr=strErr & "<table width=""50%"" border=""0"" cellpadding=""5"" cellspacing=""1"" bgcolor=""cad9ea"" align=""center"">" & vbcrlf
	strErr=strErr & "<tr>" & vbcrlf
	strErr=strErr & "<td bgcolor=""e8f3fd"">提示信息</td>" & vbcrlf
	strErr=strErr & "</tr>" & vbcrlf
	strErr=strErr & "<tr>" & vbcrlf
	strErr=strErr & "<td bgcolor=""#FFFFFF"">" & Message &"" & vbcrlf
	strErr=strErr & "<br>" & vbcrlf
	strErr=strErr & "<a href=""javascript:history.back(1)"">点击这里返回</a>" & vbcrlf
	strErr=strErr & "</td>" & vbcrlf
	strErr=strErr & "</tr>" & vbcrlf
	strErr=strErr & "</table>" & vbcrlf
	strErr=strErr &  vbCrlf
	strErr=strErr & "</body></html>" & vbcrlf
	response.write strErr
end sub

'****************************************************
'过程名：performMsg
'作  用：显示成功提示信息
'参  数：无
'****************************************************
Sub performMsg(Perform, BackUrl)
    Dim strPerform
	strPerform=strPerform & "<html><head><title>提示信息</title><meta http-equiv='Content-Type' content='text/html; charset=gb2312'>" & vbcrlf
	strPerform=strPerform & "<style type='text/css'>" & vbcrlf
	strPerform=strPerform & "<!--" & vbcrlf
	strPerform=strPerform & "body {font:12.8px Arial, Helvetica, sans-serif;margin-top: 0px; margin-right: 0px; margin-bottom: 0px; margin-left: 0px;COLOR: #6e6e6e;}td {font: 12px/1.6em Tahoma,Verdana;line-height: 18.8px;}A:link {COLOR: #000000; TEXT-DECORATION: none}A:active {COLOR: #000000;TEXT-DECORATION: underline}A:visited {COLOR: #000000; TEXT-DECORATION: none}A:hover {COLOR: #000000;TEXT-DECORATION: underline}" & vbcrlf
	strPerform=strPerform & "-->" & vbcrlf
	strPerform=strPerform & "</style>" & vbcrlf
	strPerform=strPerform & "</head><body><br><br>" & vbcrlf
	strPerform=strPerform & "<table width=""50%"" border=""0"" cellpadding=""5"" cellspacing=""1"" bgcolor=""cad9ea"" align=""center"">" & vbcrlf
	strPerform=strPerform & "<tr>" & vbcrlf
	strPerform=strPerform & "<td bgcolor=""e8f3fd"">提示信息</td>" & vbcrlf
	strPerform=strPerform & "</tr>" & vbcrlf
	strPerform=strPerform & "<tr>" & vbcrlf
	strPerform=strPerform & "<td bgcolor=""#FFFFFF"">" & Perform &"" & vbcrlf
	strPerform=strPerform & "<br>" & vbcrlf
    If BackUrl <> "" Then
	strPerform = strPerform & "<a href='" & BackUrl & "'>点击这里返回</a>"
    Else
	strPerform = strPerform & "<a href='" & Trim(Request.ServerVariables("HTTP_REFERER")) & "'>点击这里返回</a>"
    End If
	strPerform=strPerform & "</td>" & vbcrlf
	strPerform=strPerform & "</tr>" & vbcrlf
	strPerform=strPerform & "</table>" & vbcrlf
	strPerform=strPerform & "</body></html>" & vbcrlf
	response.write strPerform
end Sub

'**************************************************
'函数名：GetIsAdmin
'作  用：检查后台用户是否登录
'参  数：无
'返回值：True ----已经登录
'        False ---没有登录
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
'函数名：GetCity
'作  用：省市
'****************************************************
Function GetCity(UserProvince, UserCity)
	Dim tmpstr
	tmpstr = "<select onchange=setcity(); name='UserProvince' style='width:170px'>"
	tmpstr = tmpstr & "<option value=''>-请选择省份-</option>"
	tmpstr = tmpstr & "<option "
	tmpstr = tmpstr & "value=安徽>安徽</option> <option value=北京>北京</option> "
	tmpstr = tmpstr & "<option value=重庆>重庆</option> <option "
	tmpstr = tmpstr & "value=福建>福建</option> <option value=甘肃>甘肃</option> "
	tmpstr = tmpstr & "<option value=广东>广东</option> <option "
	tmpstr = tmpstr & "value=广西>广西</option> <option value=贵州>贵州</option> "
	tmpstr = tmpstr & "<option value=海南>海南</option> <option "
	tmpstr = tmpstr & "value=河北>河北</option> <option value=黑龙江>黑龙江</option> "
	tmpstr = tmpstr & "<option value=河南>河南</option> <option "
	tmpstr = tmpstr & "value=香港>香港</option> <option value=湖北>湖北</option> "
	tmpstr = tmpstr & "<option value=湖南>湖南</option> <option "
	tmpstr = tmpstr & "value=江苏>江苏</option> <option value=江西>江西</option> "
	tmpstr = tmpstr & "<option value=吉林>吉林</option> <option "
	tmpstr = tmpstr & "value=辽宁>辽宁</option> <option value=澳门>澳门</option>"
	tmpstr = tmpstr & "<option value=内蒙古>内蒙古</option> <option "
	tmpstr = tmpstr & "value=宁夏>宁夏</option> <option value=青海>青海</option> "
	tmpstr = tmpstr & "<option value=山东>山东</option> <option "
	tmpstr = tmpstr & "value=上海>上海</option> <option value=山西>山西</option> "
	tmpstr = tmpstr & "<option value=陕西>陕西</option> <option "
	tmpstr = tmpstr & "value=四川>四川</option> <option value=台湾>台湾</option> "
	tmpstr = tmpstr & "<option value=天津>天津</option> <option "
	tmpstr = tmpstr & "value=新疆>新疆</option> <option value=西藏>西藏</option> "
	tmpstr = tmpstr & "<option value=云南>云南</option> <option "
	tmpstr = tmpstr & "value=浙江>浙江</option> <option "
	tmpstr = tmpstr & "value=海外>海外</option></select>"
	tmpstr = tmpstr & " <select name='UserCity' >"
	tmpstr = tmpstr & "</select>"
	tmpstr = tmpstr & "<SCRIPT src=""inc/inc.area.js""></SCRIPT>"
	tmpstr = tmpstr & "<SCRIPT>initprovcity('" & UserProvince & "','" & UserCity & "');</SCRIPT>"
	GetCity = tmpstr
End Function

'**************************************************
'函数名：strCLng
'作  用：将字符转为整型数值 
'**************************************************
Function strCLng(ByVal str)
    If IsNumeric(str) Then
        strCLng = Fix(CDbl(str))
    Else
        strCLng = 0
    End If
End Function

'**************************************************
'函数名：GetNumberRound
'作  用：计算平均数
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
'函数名：strleach
'作  用：过滤非法字符函数 
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
'函数名：Getfinddir
'作  用：找到文件地址的全路径
'**************************************************
Function Getfinddir(filepath)
	Getfinddir=left(filepath,instrRev(filepath,"/"))
end Function

'**************************************************
'函数名：GetNow
'作  用：取日期
'**************************************************
Function GetNow(ByVal str)
	uNow = dateadd("h",Site_MasterTimeZone-smart_ZoneServer,now())
	strToday=datevalue(uNow) '今日
	strTomorrow=dateadd("d",+1,strToday) '明天
	strYesterday=dateadd("d",-1,strToday) '昨日
	strday7=dateadd("d",-6,strToday) '最近7天
	strday30=dateadd("d",-29,strToday) '最近30天
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
'自 定 义 函 数 和 子 程 序
' ********************************************************

' 保存客户端信息
Sub GetFromClient(SNum,SCon)
	set tmpRS=conn.Execute("select * from tblClient where LogClientType=" & SNum & " and LogClientContent='"&SCon&"' and Site_ID=" & SiteID)
	if tmpRS.eof then
		conn.Execute ("insert into tblClient (Site_ID,LogClientType,LogClientContent,LogClientTotal,LogClientYesterday,LogClientToday,LogClientLastTime) Values("&Siteid&","&SNum&",'"&SCon&"',1,0,1,'"&truenow&"')")
	else
	    conn.Execute ("update tblClient set LogClientTotal=LogClientTotal+1,LogClientToday=LogClientToday+1,LogClientLastTime='"&truenow&"'  where LogClientType=" & SNum & " and LogClientContent='"&SCon&"' and Site_ID=" & SiteID)
	end if
	set tmpRS=nothing
end Sub

' 保存内容信息
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

' 找到当前URL对应的站点
function findhost(furl)
	if furl<> "" then
	ffurl		= split(furl,"/")
	findhost	= ffurl(2)
	if left(findhost,8)="192.168." or left(findhost,3)="10." or findhost="127.0.0.1" or instr(findhost,".")=0 then findhost="LAN"
	else 
	findhost	= ""
	end if
end function

' 找到文件地址的全路径
Function finddir(filepath)
	finddir=left(filepath,instrRev(filepath,"/")-1)
end Function

' 找到IP地址对应的地区
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
		if instr(findarea,"未知") then findarea=""
	end if
end function

'根据地区分离省份
Function GetProvince(strCountry) '简单分离出省份
	If Len(trim(strCountry)) = 0 Then
	GetProvince = "未知"
	Exit Function
	End If
	On Error Resume Next
	Dim strProvince,i
	strProvince= strProvince & "北京市|上海市|天津市|重庆市|香港|澳门|广东省|河北省|山西省|内蒙古|辽宁省|吉林省|黑龙江省|江苏省"
	strProvince= strProvince & "|浙江省|安徽省|福建省|江西省|山东省|河南省|湖北省|湖南省|广西|海南省|四川省|贵州省|云南省|西藏|陕西省|甘肃省|青海省|宁夏|新疆|台湾省"
	strProvince=split(strProvince,"|")
	for i=0 to ubound(strProvince)
	if instr(strCountry,strProvince(i))>0 then
	   GetProvince=strProvince(i)
	   exit function
	end if
	next
	GetProvince = "未知"
End Function

'省转拼单
function GetChinaMap(strLang)
  select case strLang
  case "北京市"
    outstr="CN.BJ"
  case "上海市"
    outstr="CN.SH"
  case "天津市"
    outstr="CN.TJ"
  case "重庆市"
    outstr="CN.CQ"
  case "香港"
    outstr="CN.HK"
  case "澳门"
    outstr="CN.MA"
  case "广东省"
    outstr="CN.GD"
  case "河北省"
    outstr="CN.HB"
  case "山西省"
    outstr="CN.SX"
  case "内蒙古"
    outstr="CN.NM"
  case "辽宁省"
    outstr="CN.LN"
  case "吉林省"
    outstr="CN.JL"
  case "黑龙江省"
    outstr="CN.HL"
  case "浙江省"
    outstr="CN.ZJ"
  case "安徽省"
    outstr="CN.AH"
  case "福建省"
    outstr="CN.FJ"
  case "江西省"
    outstr="CN.JX"
  case "山东省"
    outstr="CN.SD"
  case "河南省"
    outstr="CN.HE"
  case "湖北省"
    outstr="CN.HB"
  case "湖南省"
    outstr="CN.HN"
  case "广西"
    outstr="CN.GX"
  case "海南省"
    outstr="CN.HA"
  case "四川省"
    outstr="CN.SC"
  case "贵州省"
    outstr="CN.GZ"
  case "云南省"
    outstr="CN.YN"
  case "西藏"
    outstr="CN.XZ"
  case "陕西省"
    outstr="CN.SA"
  case "甘肃省"
    outstr="CN.GS"
  case "青海省"
    outstr="CN.QH"
  case "宁夏"
    outstr="CN.NX"
  case "新疆"
    outstr="CN.XJ"
  case "台湾省"
    outstr="CN.TA"
  case else
    outstr="CN.BJ"
  end select
  GetChinaMap=outstr
end function

' 写入最后访问的用户
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

'更新要保存的IP
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

'语言函数
function GetLangchs(strLang)
  select case strLang
  case "af"
    outstr="南非荷兰语"
  case "sq"
    outstr="阿尔巴尼亚语"
  case "ar-ae"
    outstr="阿拉伯语 - 阿拉伯联合酋长国"
  case "ar-bh"
    outstr="阿拉伯语 - 巴林"
  case "ar-dz"
    outstr="阿拉伯语 - 阿尔及利亚"
  case "ar-eg"
    outstr="阿拉伯语 - 埃及"
  case "ar-iq"
    outstr="阿拉伯语 - 伊拉克"
  case "ar-jo"
    outstr="阿拉伯语 - 约旦"
  case "ar-kw"
    outstr="阿拉伯语 - 科威特"
  case "ar-lb"
    outstr="阿拉伯语 - 黎巴嫩"
  case "ar-ly"
    outstr="阿拉伯语 - 利比亚"
  case "ar-ma"
    outstr="阿拉伯语 - 摩洛哥"
  case "ar-om"
    outstr="阿拉伯语 - 阿曼"
  case "ar-qa"
    outstr="阿拉伯语 - 卡塔尔"
  case "ar-sa"
    outstr="阿拉伯语 - 沙特阿拉伯"
  case "ar-sy"
    outstr="阿拉伯语 - 叙利亚"
  case "ar-tn"
    outstr="阿拉伯语 - 突尼斯"
  case "ar-ye"
    outstr="阿拉伯语 - 也门"
  case "hy"
    outstr="亚美尼亚语"
  case "az-az"
    outstr="阿泽里语 - 拉丁"
  case "az-az"
    outstr="阿泽里语 - 西里尔语"
  case "eu"
    outstr="巴斯克语"
  case "be"
    outstr="白俄罗斯语"
  case "bg"
    outstr="保加利亚语"
  case "ca"
    outstr="加泰罗尼亚语"
  case "zh"
    outstr="中文"
  case "zh-cn"
    outstr="中文 - 中华人民共和国"
  case "zh-hk"
    outstr="中文 - 中华人民共和国香港特别行政区"
  case "zh-mo"
    outstr="中文 - 中华人民共和国澳门特别行政区"
  case "zh-sg"
    outstr="中文 - 新加坡"
  case "zh-tw"
    outstr="中文 - 台湾地区"
  case "hr"
    outstr="克罗地亚语"
  case "cs"
    outstr="捷克语"
  case "da"
    outstr="丹麦语"
  case "nl"
    outstr="荷兰语"
  case "nl-nl"
    outstr="荷兰语"
  case "nl-be"
    outstr="荷兰语 - 比利时"
  case "en"
    outstr="英语"
  case "en-au"
    outstr="英语 - 澳大利亚"
  case "en-bz"
    outstr="英语 - 伯利兹"
  case "en-ca"
    outstr="英语 - 加拿大"
  case "en-cb"
    outstr="英语 - 加勒比地区"
  case "en-ie"
    outstr="英语 - 爱尔兰"
  case "en-jm"
    outstr="英语 - 牙买加"
  case "en-nz"
    outstr="英语 - 新西兰"
  case "en-ph"
    outstr="英语 - 菲律宾"
  case "en-za"
    outstr="英语 - 南非"
  case "en-tt"
    outstr="英语 - 特立尼达岛"
  case "en-gb"
    outstr="英语 - 英国"
  case "en-us"
    outstr="英语 - 美国"
  case "et"
    outstr="爱沙尼亚语"
  case "fa"
    outstr="波斯语"
  case "fi"
    outstr="芬兰语"
  case "fo"
    outstr="法罗语"
  case "fr"
    outstr="法语"
  case "fr-fr"
    outstr="法语 - 法国"
  case "fr-be"
    outstr="法语 - 比利时"
  case "fr-ca"
    outstr="法语 - 加拿大"
  case "fr-lu"
    outstr="法语 - 卢森堡"
  case "fr-ch"
    outstr="法语 - 瑞士"
  case "gd-ie"
    outstr="盖尔语 - 爱尔兰"
  case "gd"
    outstr="盖尔语 - 苏格兰"
  case "de"
    outstr="德语"
  case "de-de"
    outstr="德语 - 德国"
  case "de-at"
    outstr="德语 - 奥地利"
  case "de-li"
    outstr="德语 - 列支敦士登"
  case "de-lu"
    outstr="德语 - 卢森堡"
  case "de-ch"
    outstr="德语 - 瑞士"
  case "el"
    outstr="希腊语"
  case "he"
    outstr="希伯来语"
  case "hi"
    outstr="印地语"
  case "hu"
    outstr="匈牙利语"
  case "is"
    outstr="冰岛语"
  case "id"
    outstr="印度尼西亚语"
  case "it"
    outstr="意大利语"
  case "it-it"
    outstr="意大利语 - 意大利"
  case "it-ch"
    outstr="意大利语 - 瑞士"
  case "ja"
    outstr="日语"
  case "ko"
    outstr="朝鲜语"
  case "lv"
    outstr="拉脱维亚语"
  case "lt"
    outstr="立陶宛语"
  case "mk"
    outstr="FYRO 马其顿语"
  case "ms-my"
    outstr="马来语 - 马来西亚"
  case "ms-bn"
    outstr="马来语 - 文莱"
  case "mt"
    outstr="马耳他语"
  case "mr"
    outstr="马拉地语"
  case "no"
    outstr="挪威语"
  case "no-no"
    outstr="挪威语 - 博克马尔"
  case "no-no"
    outstr="挪威语 - 尼诺斯克"
  case "pl"
    outstr="波兰语"
  case "pt"
    outstr="葡萄牙语"
  case "pt-pt"
    outstr="葡萄牙语 - 葡萄牙"
  case "pt-br"
    outstr="葡萄牙语 - 巴西"
  case "rm"
    outstr="拉托-罗马语"
  case "ro"
    outstr="罗马尼亚语"
  case "ro-mo"
    outstr="罗马尼亚语 - 摩尔多瓦"
  case "ru"
    outstr="俄语"
  case "ru-mo"
    outstr="俄语 - 摩尔多瓦"
  case "sa"
    outstr="梵语"
  case "sr"
    outstr="塞尔维亚语"
  case "sr-sp"
    outstr="塞尔维亚语 - 西里尔语"
  case "sr-sp"
    outstr="塞尔维亚语 - 拉丁"
  case "tn"
    outstr="茨瓦纳语"
  case "sl"
    outstr="斯洛文尼亚语"
  case "sk"
    outstr="斯洛伐克语"
  case "sb"
    outstr="索布语"
  case "es"
    outstr="西班牙语"
  case "es-es"
    outstr="西班牙语 - 西班牙"
  case "es-ar"
    outstr="西班牙语 - 阿根廷"
  case "es-bo"
    outstr="西班牙语 - 玻利维亚"
  case "es-cl"
    outstr="西班牙语 - 智利"
  case "es-co"
    outstr="西班牙语 - 哥伦比亚"
  case "es-cr"
    outstr="西班牙语 - 哥斯达黎加"
  case "es-do"
    outstr="西班牙语 - 多米尼加共和国"
  case "es-ec"
    outstr="西班牙语 - 厄瓜多尔"
  case "es-gt"
    outstr="西班牙语 - 危地马拉"
  case "es-hn"
    outstr="西班牙语 - 洪都拉斯"
  case "es-mx"
    outstr="西班牙语 - 墨西哥"
  case "es-ni"
    outstr="西班牙语 - 尼加拉瓜"
  case "es-pa"
    outstr="西班牙语 - 巴拿马"
  case "es-pe"
    outstr="西班牙语 - 秘鲁"
  case "es-pr"
    outstr="西班牙语 - 波多黎各"
  case "es-py"
    outstr="西班牙语 - 巴拉圭"
  case "es-sv"
    outstr="西班牙语 - 萨尔瓦多"
  case "es-uy"
    outstr="西班牙语 - 乌拉圭"
  case "es-ve"
    outstr="西班牙语 - 委内瑞拉"
  case "sx"
    outstr="苏图语"
  case "sw"
    outstr="斯瓦希里语"
  case "sv"
    outstr="瑞典语"
  case "sv-se"
    outstr="瑞典语"
  case "sv-fi"
    outstr="瑞典语 - 芬兰"
  case "ta"
    outstr="泰米尔语"
  case "tt"
    outstr="鞑靼语"
  case "th"
    outstr="泰语"
  case "tr"
    outstr="土耳其语"
  case "ts"
    outstr="汤加语"
  case "uk"
    outstr="乌克兰语"
  case "ur"
    outstr="乌尔都语 - 巴基斯坦"
  case "uz-uz"
    outstr="乌兹别克语 - 西里尔"
  case "uz-uz"
    outstr="乌兹别克语 - 拉丁"
  case "vi"
    outstr="越南语"
  case "xh"
    outstr="科萨语"
  case "yi"
    outstr="意第绪语"
  case "zu"
    outstr="祖鲁语"
  case else
    outstr="未知"
  end select
  GetLangchs=outstr & " (" & strLang & ")"
end function

'定义搜索
dim array_Search(6,1)
array_Search(0,0) = "google."
array_Search(0,1) = "Google"

array_Search(1,0) = "baidu."
array_Search(1,1) = "百度"

array_Search(2,0) = "youdao."
array_Search(2,1) = "有道"

array_Search(3,0) = "sogou."
array_Search(3,1) = "搜狗"

array_Search(4,0)	= "bing."
array_Search(4,1)	= "必应"

array_Search(5,0)	= "soso."
array_Search(5,1)	= "SOSO"

array_Search(6,0)	= "yahoo."
array_Search(6,1)	= "雅虎"

'**************************************************
'函数名：referrerParsing
'作  用：取搜索引擎名称及关键字
'**************************************************
sub referrerParsing(referrer, searchEngine, searchKeywords)
	If instr(referrer,"google")<>0 then
		searchEngine		="谷歌"   
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
		searchEngine		="百度"   
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
		searchEngine		="百度"   
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
		searchEngine		="有道"   
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
		searchEngine		="搜狗"   
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
		searchEngine		="搜搜"   
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
		searchEngine		="必应"   
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
		searchEngine		="雅虎"   
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

'编码转换
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