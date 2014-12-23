<!--#include file="inc/plus.inc.asp"-->
<%
on error resume next
public vArea
t1=timer()

isCan = smart_IsCan

newdayCacheName=smart_CacheName & "_NewDay"
if isempty(Application(newdayCacheName)) then Application(newdayCacheName)=cdate("1900-1-1")

LastIPCacheName=smart_CacheName & "_LastIP_" & Siteid
if isempty(Application(LastIPCacheName)) then Application(LastIPCacheName)="#218.246.226.8#"

public Style,vZone,vColor,vSize,vCome,vPage,vLang,vIP,vIPt,vSoft,vOS,theurl,vAgent,truenow,UserSaveNum
dim IsRefurbish,sql,IsNewDay

canCacheName=smart_CacheName & "_can"
if isempty(Application(canCacheName)) then Application(canCacheName)=0

LastTCacheName=smart_CacheName & "_LastTime"
if isempty(Application(LastTCacheName)) then Application(LastTCacheName)=now()

if now()-Application(LastTCacheName) > 0.01 then Application(canCacheName) = 0:Application(LastTCacheName)=now()

if Application(canCacheName) >= 2 then isCan=false

' 网页立即超时，防止漏统计
Response.Expires = 0

if isCan then
    Application.Lock
    Application(canCacheName) = Application(canCacheName) + 1
    Application.UnLock
	'********************** 获取数据 **********************
	theurl = "http://" & Request.ServerVariables("http_host") & finddir(Request.ServerVariables("url"))
	vStyle = Request("style")
	vZone = Request("tzone")
	vColor = Request("tcolor")
	vSize = Request("sSize")
	vCome = Request("referrer")
	'response.write vCome
	vPage = Request.ServerVariables("HTTP_REFERER")
	FromUrl = replace(Trim(Request("fromurl")&""),",","")
	
	IF FromUrl <> "" then
		vCome = FromUrl
		vPage = FromUrl
	end if
	
	if vCome="" then
		vCome = "直接访问"
	end if
	
	vComeHost = findhost(vCome)
	vPageHost = findhost(vPage)
	if right(vPage,1)="/" then vPage=left(vPage,len(vPage)-1)
	vLang = Request.ServerVariables("HTTP_ACCEPT_LANGUAGE")
	vLang = split(vLang,",")(0)
	vLang = lcase(split(vLang,";")(0))
	vIP = Request.ServerVariables("Remote_Addr")
	vIPs = split(vIP,".")
	vIPt = vIPs(0)
	if smart_IPLong>1 then vIPt=vIPt & "." & vIPs(1)
	vAgent = Request.ServerVariables("HTTP_USER_AGENT")
	if instr(vAgent,"Alexa") then
		vAlexa = "1"
	else
		vAlexa = "0"
	end if

	'零时区的当前时间
	truenow = dateadd("h",0 - smart_ZoneServer,now())

	'是否新的一天
	IsNewDay = false 
	if DateValue(Site_TodayDate) < DateValue(dateadd("h",Site_MasterTimeZone-smart_ZoneServer,now())) then IsNewDay=true'是否新的一天
	if cdate(Site_StartTime) < cdate("2000-1-1") then
		conn.Execute ("update tblSite set UserStartTime=(now()-"&smart_ZoneServer&"/24) where Site_ID=" & siteID)
		IsNewDay=true
	end if

	' 新的一天，今天→昨天，今天=0
	if IsNewDay then
		if Application(newdayCacheName)<=now()-1 then
		Application.Lock:Application(newdayCacheName)=now():Application.UnLock 

		' 为流量库添加当天的所有行
		today0hour=dateadd("h",0-Site_MasterTimeZone,datevalue(now()))
		for i= 0 to 23
		  conn.execute ("delete * from tblView where Site_ID="&SiteID&" and UserTime=#"&dateadd("h",i,today0hour)&"#")
		  conn.execute ("insert into tblView (Site_ID,UserTime,UserView,UserViewIP) Values("&Siteid&",'"&dateadd("h",i,today0hour)&"',0,0)")
		next

		' 删除内容信息中较陈旧的
		if smart_AutoDel > 0 then conn.execute "delete * from tblOriginPage where LogPageLastTime <= (now()-"&smart_ZoneServer&"/24-"&smart_AutoDel&")"

		' 更新SITE表的最后日期
		conn.Execute ("update tblSite set UserTodayDate = datevalue(now()+"&Site_MasterTimeZone-smart_ZoneServer&"/24) where Site_ID=" & SiteID)
		conn.Execute ("update tblClient set LogClientYesterday=LogClientToday,LogClientToday=0 where Site_ID=" & SiteID)
		conn.Execute ("update tblOriginPage set LogPageYesterday=LogPageToday,LogPageToday=0 where Site_ID=" & SiteID)
		Application.Lock:Application(newdayCacheName)=cdate("1900-1-1"):Application.UnLock 
		End if
	End if

' 是否刷新
	vUser	= clng(Request.Cookies("oSmart"&smart_CacheName&SiteID)("oSmartstat"))		' 当前用户访问量
	vPageS	= clng(Request.Cookies("smartstat"&smart_CacheName&SiteID)("smartstatPages"))	' 当前用户本次浏览页面数
	vUPageS	= clng(Request.Cookies("smartstat"&smart_CacheName&SiteID)("UserPages"))		' 当前用户浏览页面总数

	IsRefurbish = false
	if instr(Application(LastIPCacheName),"#" & vIP & "#") then IsRefurbish=true	' 如果IP已经存在于保存的列表中，是刷新
'	if vComeHost=vPageHost then IsRefurbish=true				' 如果来路站点和被访问站点是同一个站点，则是刷新
	if IsRefurbish then
		vPageS	= vPageS + 1
	else
		vPageS	= 1
		vUser	= vUser+1
		Response.Cookies("oSmart"&smart_CacheName&SiteID)("oSmartstat")=vuser
		Response.Cookies("oSmart"&smart_CacheName&SiteID).Expires=dateadd("d",100,date() )
		' 更新最近需要防刷的IP
		Application.Lock 
		Application(LastIPCacheName)=vsaveips(Application(LastIPCacheName))
		Application.UnLock
	End if
	Response.Cookies("smartstat"&smart_CacheName&SiteID)("smartstatPages")=vPageS
	Response.Cookies("smartstat"&smart_CacheName&SiteID)("UserPages")=vUPageS + 1
	Response.Cookies("smartstat"&smart_CacheName&SiteID).Expires=dateadd("d",100,date() )

' ==========================================================
'                     写 入 详 细 信 息
' ==========================================================
	
if not IsRefurbish Then ' Client类型
	set smart_cls=new CheckUserAgent
	smart_cls.execute vAgent
	vOs = smart_cls.vos
	vSoft = smart_cls.vsoft

	if vos<> "" then GetFromClient 0,vOs' 0 操作系统
	if vsoft<>"" then							' 1 浏览器
		if instr(vSoft,",") then
			vvsoft=split(vsoft,",")
			for each dsoft in vvsoft
				if trim(dsoft)<> "" then GetFromClient 1,dSoft
			next
		else
			GetFromClient 1,vSoft
		end if
	end if
	GetFromClient 2,vLang				' 2 语言
	GetFromClient 3,vZone				' 3 时区
	GetFromClient 4,vSize				' 4 屏幕大小
	GetFromClient 5,vColor			' 5 屏幕色彩
	GetFromClient 6,vUser				' 6 访问次数
	GetFromClient 9,vAlexa			' 9 ALEXA工具条
	vArea=findArea(vIP)
	vProvince=GetProvince(vArea)
	if vArea<>"" then GetFromClient 7,vArea ' 8 城市
	if vProvince<>"" then GetFromClient 10,vProvince ' 10 省
end if
	GetFromClient 8,vPageS				' 8 浏览的页面数

' ==========================================================

if not IsRefurbish then 
	GetFromDatabase 0,vComeHost,vcome		' 0 来路
	GetFromDatabase 1,vPage,""				' 1 入口
	pSearchEngine	=""
	pSearchKeywords	=""

	Call referrerParsing(vCome, pSearchEngine, pSearchKeywords)
	if pSearchKeywords<> "" then pSearchKeywords=trim(Lcase(pSearchKeywords)) : GetFromDatabase 2,pSearchKeywords,vcome' 2 关键词
	GetFromDatabase 3,vIPt,vIP				' 3 IP
end if
	GetFromDatabase 4,vPage,""					' 4 页面

' ==========================================================
'                     写 入 流 量 信 息
' ==========================================================

	nowHour=cdate(DateValue(truenow) & " " & hour(truenow)&":00:00")      ' 现在是几点
	if IsRefurbish then		' 如果是刷新，则只更新浏览量
		conn.execute("update tblView set UserView=UserView+1 where Site_ID=" & SiteID & " and UserTime=#" & nowHour & "#")
	else				' 如果不是刷新，则更新浏览量和访问量
		conn.execute("update tblView set UserView=UserView+1,UserViewIP=UserViewIP+1 where Site_ID=" & SiteID & " and UserTime=#" & nowHour & "#")
	end if
	
	' 写入最后访问用户到服务器缓存
	if not IsRefurbish then call GetLastUser()

	' 关闭数据库
	call closeconn

    Application.Lock
    Application(canCacheName) = Application(canCacheName) - 1
    Application.UnLock
end if'if isCan
%>