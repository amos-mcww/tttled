<%
' ======================== 参 数 设 置 ===========================
public smart_DefaultSite,smart_ZoneServer,smart_IPLong,smart_CheckOnline,smart_CacheName,smart_IsCan,smart_SaveOnlineUser,smart_AutoDel,smart_SaveOnline
'系统名称
smart_SystemName ="SmartStat Ver2.3"
'系统版权
smart_SystemCopyright = " Powered By <a href=http://www.cactussoft.cn/>SmartStat</a> Version 2.3"
'默认站点的ID
smart_DefaultSite = 1
'服务器所在时区
smart_ZoneServer = 8
'IP保存位数
smart_IPLong = 1
'检测在线用户间隔（秒）
smart_CheckOnline	= 40
'服务器缓存名称
smart_CacheName = "SmartStat-"
'统计器启用
smart_IsCan = True
'自动删除×天前的内容信息
smart_AutoDel			= 10
'允许使用在线用户统计的站点
smart_SaveOnlineUser	= "All"

' ===================== 公用代码，请勿修改 =======================
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