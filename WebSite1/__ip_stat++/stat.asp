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

' ��ҳ������ʱ����ֹ©ͳ��
Response.Expires = 0

if isCan then
    Application.Lock
    Application(canCacheName) = Application(canCacheName) + 1
    Application.UnLock
	'********************** ��ȡ���� **********************
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
		vCome = "ֱ�ӷ���"
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

	'��ʱ���ĵ�ǰʱ��
	truenow = dateadd("h",0 - smart_ZoneServer,now())

	'�Ƿ��µ�һ��
	IsNewDay = false 
	if DateValue(Site_TodayDate) < DateValue(dateadd("h",Site_MasterTimeZone-smart_ZoneServer,now())) then IsNewDay=true'�Ƿ��µ�һ��
	if cdate(Site_StartTime) < cdate("2000-1-1") then
		conn.Execute ("update tblSite set UserStartTime=(now()-"&smart_ZoneServer&"/24) where Site_ID=" & siteID)
		IsNewDay=true
	end if

	' �µ�һ�죬��������죬����=0
	if IsNewDay then
		if Application(newdayCacheName)<=now()-1 then
		Application.Lock:Application(newdayCacheName)=now():Application.UnLock 

		' Ϊ��������ӵ����������
		today0hour=dateadd("h",0-Site_MasterTimeZone,datevalue(now()))
		for i= 0 to 23
		  conn.execute ("delete * from tblView where Site_ID="&SiteID&" and UserTime=#"&dateadd("h",i,today0hour)&"#")
		  conn.execute ("insert into tblView (Site_ID,UserTime,UserView,UserViewIP) Values("&Siteid&",'"&dateadd("h",i,today0hour)&"',0,0)")
		next

		' ɾ��������Ϣ�нϳ¾ɵ�
		if smart_AutoDel > 0 then conn.execute "delete * from tblOriginPage where LogPageLastTime <= (now()-"&smart_ZoneServer&"/24-"&smart_AutoDel&")"

		' ����SITE����������
		conn.Execute ("update tblSite set UserTodayDate = datevalue(now()+"&Site_MasterTimeZone-smart_ZoneServer&"/24) where Site_ID=" & SiteID)
		conn.Execute ("update tblClient set LogClientYesterday=LogClientToday,LogClientToday=0 where Site_ID=" & SiteID)
		conn.Execute ("update tblOriginPage set LogPageYesterday=LogPageToday,LogPageToday=0 where Site_ID=" & SiteID)
		Application.Lock:Application(newdayCacheName)=cdate("1900-1-1"):Application.UnLock 
		End if
	End if

' �Ƿ�ˢ��
	vUser	= clng(Request.Cookies("oSmart"&smart_CacheName&SiteID)("oSmartstat"))		' ��ǰ�û�������
	vPageS	= clng(Request.Cookies("smartstat"&smart_CacheName&SiteID)("smartstatPages"))	' ��ǰ�û��������ҳ����
	vUPageS	= clng(Request.Cookies("smartstat"&smart_CacheName&SiteID)("UserPages"))		' ��ǰ�û����ҳ������

	IsRefurbish = false
	if instr(Application(LastIPCacheName),"#" & vIP & "#") then IsRefurbish=true	' ���IP�Ѿ������ڱ�����б��У���ˢ��
'	if vComeHost=vPageHost then IsRefurbish=true				' �����·վ��ͱ�����վ����ͬһ��վ�㣬����ˢ��
	if IsRefurbish then
		vPageS	= vPageS + 1
	else
		vPageS	= 1
		vUser	= vUser+1
		Response.Cookies("oSmart"&smart_CacheName&SiteID)("oSmartstat")=vuser
		Response.Cookies("oSmart"&smart_CacheName&SiteID).Expires=dateadd("d",100,date() )
		' ���������Ҫ��ˢ��IP
		Application.Lock 
		Application(LastIPCacheName)=vsaveips(Application(LastIPCacheName))
		Application.UnLock
	End if
	Response.Cookies("smartstat"&smart_CacheName&SiteID)("smartstatPages")=vPageS
	Response.Cookies("smartstat"&smart_CacheName&SiteID)("UserPages")=vUPageS + 1
	Response.Cookies("smartstat"&smart_CacheName&SiteID).Expires=dateadd("d",100,date() )

' ==========================================================
'                     д �� �� ϸ �� Ϣ
' ==========================================================
	
if not IsRefurbish Then ' Client����
	set smart_cls=new CheckUserAgent
	smart_cls.execute vAgent
	vOs = smart_cls.vos
	vSoft = smart_cls.vsoft

	if vos<> "" then GetFromClient 0,vOs' 0 ����ϵͳ
	if vsoft<>"" then							' 1 �����
		if instr(vSoft,",") then
			vvsoft=split(vsoft,",")
			for each dsoft in vvsoft
				if trim(dsoft)<> "" then GetFromClient 1,dSoft
			next
		else
			GetFromClient 1,vSoft
		end if
	end if
	GetFromClient 2,vLang				' 2 ����
	GetFromClient 3,vZone				' 3 ʱ��
	GetFromClient 4,vSize				' 4 ��Ļ��С
	GetFromClient 5,vColor			' 5 ��Ļɫ��
	GetFromClient 6,vUser				' 6 ���ʴ���
	GetFromClient 9,vAlexa			' 9 ALEXA������
	vArea=findArea(vIP)
	vProvince=GetProvince(vArea)
	if vArea<>"" then GetFromClient 7,vArea ' 8 ����
	if vProvince<>"" then GetFromClient 10,vProvince ' 10 ʡ
end if
	GetFromClient 8,vPageS				' 8 �����ҳ����

' ==========================================================

if not IsRefurbish then 
	GetFromDatabase 0,vComeHost,vcome		' 0 ��·
	GetFromDatabase 1,vPage,""				' 1 ���
	pSearchEngine	=""
	pSearchKeywords	=""

	Call referrerParsing(vCome, pSearchEngine, pSearchKeywords)
	if pSearchKeywords<> "" then pSearchKeywords=trim(Lcase(pSearchKeywords)) : GetFromDatabase 2,pSearchKeywords,vcome' 2 �ؼ���
	GetFromDatabase 3,vIPt,vIP				' 3 IP
end if
	GetFromDatabase 4,vPage,""					' 4 ҳ��

' ==========================================================
'                     д �� �� �� �� Ϣ
' ==========================================================

	nowHour=cdate(DateValue(truenow) & " " & hour(truenow)&":00:00")      ' �����Ǽ���
	if IsRefurbish then		' �����ˢ�£���ֻ���������
		conn.execute("update tblView set UserView=UserView+1 where Site_ID=" & SiteID & " and UserTime=#" & nowHour & "#")
	else				' �������ˢ�£������������ͷ�����
		conn.execute("update tblView set UserView=UserView+1,UserViewIP=UserViewIP+1 where Site_ID=" & SiteID & " and UserTime=#" & nowHour & "#")
	end if
	
	' д���������û�������������
	if not IsRefurbish then call GetLastUser()

	' �ر����ݿ�
	call closeconn

    Application.Lock
    Application(canCacheName) = Application(canCacheName) - 1
    Application.UnLock
end if'if isCan
%>