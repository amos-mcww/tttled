<%
set dbRS=conn.Execute("select * from tblSite where Site_ID="& SiteID)
if not dbRS.eof and not dbRS.bof then
	Site_UserSiteName = dbRS("UserSiteName")
	Site_UserSiteUrl = dbRS("UserSiteUrl")
	Site_UserEmail = dbRS("UserEmail")
	Site_UserProvince = dbRS("UserProvince")
	Site_UserCity = dbRS("UserCity")
	Site_UserContent = dbRS("UserContent")
	Site_StartTime = dbRS("UserStartTime")
	Site_MasterTimeZone	= dbRS("UserMasterTimeZone")
	Site_DelRefresh	= dbRS("UserDelRefresh")
	Site_SaveNum = dbRS("UserSaveNum")
	Site_TodayDate = dbRS("UserTodayDate")
End if
%>