<%
' ======================== �� �� �� �� ===========================
public smart_DefaultSite,smart_ZoneServer,smart_IPLong,smart_CheckOnline,smart_CacheName,smart_IsCan,smart_SaveOnlineUser,smart_AutoDel,smart_SaveOnline
'ϵͳ����
smart_SystemName ="SmartStat Ver2.3"
'ϵͳ��Ȩ
smart_SystemCopyright = " Powered By <a href=http://www.cactussoft.cn/>SmartStat</a> Version 2.3"
'Ĭ��վ���ID
smart_DefaultSite = 1
'����������ʱ��
smart_ZoneServer = 8
'IP����λ��
smart_IPLong = 1
'��������û�������룩
smart_CheckOnline	= 40
'��������������
smart_CacheName = "SmartStat-"
'ͳ��������
smart_IsCan = True
'�Զ�ɾ������ǰ��������Ϣ
smart_AutoDel			= 10
'����ʹ�������û�ͳ�Ƶ�վ��
smart_SaveOnlineUser	= "All"

' ===================== ���ô��룬�����޸� =======================
Public SiteID
SiteID = Request("SiteID")
if SiteID = "" then SiteID = smart_DefaultSite
if IsNumeric(Siteid)=False then Response.Redirect "http://www.cactussoft.cn"
smart_SaveOnline = GetCanSave(SiteID,smart_SaveOnlineUser)
function GetCanSave(SiteID,CanSiteID)
	select case lcase(CanSiteID)
    case "all"
		GetCanSave=true
    case "no"
		GetCanSave=false
    case else
		if instr(","&CanSiteID&",",","&SiteID&",") then
			GetCanSave=true
		else
			GetCanSave=false
		end if
	end select
End function
%>