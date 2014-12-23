<!--#include file="inc/plus.inc.asp"-->
<%
'判断后台登录是否过期
if GetIsAdmin()=False then
	Response.Write "<script language='javascript'>"
	Response.Write "parent.parent.parent.window.navigate('"&theURL&"login.asp');"
	Response.Write "</script>"
	Response.End
end if
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.1//EN" "http://www.w3.org/TR/xhtml11/DTD/xhtml11.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<title><%=smart_SystemName%></title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<link href="style/style.css" rel="stylesheet" type="text/css">
<link href="style/dtree.css" rel="stylesheet" type="text/css">
<script type="text/javascript" src="inc/dtree.js"></script>
<script type="text/javascript" src="jquery/jquery.js"></script>
<script type="text/javascript" src="jquery/jquery.ui.all.js"></script>
<script type="text/javascript" src="jquery/jquery.layout.js"></script>
<br>
<script type="text/javascript">

	var myLayout; // a var is required because this page utilizes: myLayout.allowOverflow() method

	$(document).ready(function () {

		myLayout = $('body').layout({
			west__size:					200
		,	west__spacing_closed:		20
		,	west__togglerLength_closed:	100
		,	west__togglerAlign_closed:	"top"
		,	west__togglerContent_closed:"M<br>E<br>N<br>U"
		,	west__togglerTip_closed:	"Open & Pin Menu"
		,	west__sliderTip:			"Slide Open Menu"
		,	west__slideTrigger_open:	"mouseover"
		});

 	});

</script>
<style type="text/css">
	/**
	 *	Basic Layout Theme
	 */
	.ui-layout-pane { /* all 'panes' */ 
		border: 1px solid #BBB; 
	} 
	.ui-layout-pane-center { /* IFRAME pane */ 
		padding: 0;
		margin:  0;
	} 
	.ui-layout-pane-west { /* west pane */ 
		padding: 0 5px; 
		background-color: #fff !important;
		overflow: auto;
	} 

	.ui-layout-resizer { /* all 'resizer-bars' */ 
		background: #DDD; 
		} 
		.ui-layout-resizer-open:hover { /* mouse-over */
			background: #9D9; 
		}

	.ui-layout-toggler { /* all 'toggler-buttons' */ 
		background: #AAA; 
		} 
		.ui-layout-toggler-closed { /* closed toggler-button */ 
			background: #CCC; 
			border-bottom: 1px solid #BBB; 
		} 
		.ui-layout-toggler .content { /* toggler-text */ 
			font: 14px bold Verdana, Verdana, Arial, Helvetica, sans-serif;
		}
		.ui-layout-toggler:hover { /* mouse-over */ 
			background: #DCA; 
			} 
			.ui-layout-toggler:hover .content { /* mouse-over */ 
				color: #009; 
				}

	/* class to make the 'iframe mask' visible */
	.ui-layout-mask {
		opacity: 0.2 !important;
		filter:	 alpha(opacity=20) !important;
		background-color: #666 !important;
	}


	body {
		font-family: Geneva, Arial, Helvetica, sans-serif;
	}

	ul { /* basic menu styling */
		margin:		1ex 0;
		padding:	0;
		list-style:	none;
		position:	relative;
	}
	li {
		padding: 0.15em 1em 0.3em 5px;
	}

</style>
</head>
<body>
<div class="ui-layout-north">
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td height="60" bgcolor="#00309c"><table width="100%" border="0" cellspacing="0" cellpadding="5">
      <tr>
        <td width="50%"><a href="http://www.cactussoft.cn" target="_blank"><img src="images/logo.gif" border="0"></a></td>
        <td><table border="0" align="right" cellpadding="5" cellspacing="0">
          <tr>
            <td><a href="mailto:avram@tom.com" class="smartLink">技术支持</a></td>
            <td class="SmartWhite">|</td>
            <td><a href="http://www.cactussoft.cn" title="仙人掌软件" target="_blank" class="smartLink">关于</a></td>
          </tr>
        </table></td>
      </tr>
    </table></td>
  </tr>
  <tr>
    <td height="22" background="images/smart-stat-bg.gif">&nbsp;</td>
  </tr>
</table>
</div>
<iframe id="mainFrame" name="mainFrame" class="ui-layout-center" width="100%" height="600" frameborder="0" scrolling="auto" src="smart-stat-homepage.asp"></iframe>
<div class="ui-layout-west">
<script type="text/javascript">
		<!--
		//id, pid, name, url, title, target, icon, iconOPne, open,
		d = new dTree('d');
		d.config.target="mainFrame";		
		d.config.folderLinks = false;

// ----------------- INTRODUCTION ------------------//

		d.add(0,-1,'<B>SmartStat V2.0</B>','','');
		
		d.add(1,0,'概况','smart-stat-homepage.asp','');
		d.add(2,0,'在线用户','smart-stat-onlineuser.asp');
		d.add(3,0,'访问明细','smart-stat-lastuser.asp');

		d.add(5,0,'时段分析','','','','','',false);
		d.add(501,5,'时段分析','smart-stat-report-hour.asp','');
		d.add(502,5,'日段分析','smart-stat-report-day.asp','');
		d.add(503,5,'月分析','smart-stat-report-month.asp','');

		d.add(6,0,'搜索引擎','','','','','',false);
		d.add(601,6,'搜索引擎','smart-stat-report-searchengines.asp?TypeTo=5','');
		d.add(602,6,'关键字','smart-stat-report-searchkeywords.asp?TypeTo=2','');

		d.add(7,0,'客户端','','I','','','',false);
		d.add(701,7,'操作系统','smart-stat-report-client.asp?TypeTo=0','');
		d.add(702,7,'浏览器','smart-stat-report-client.asp?TypeTo=1','');
		d.add(703,7,'语言','smart-stat-report-client.asp?TypeTo=2','');
		d.add(704,7,'时区','smart-stat-report-client.asp?TypeTo=3','');
		d.add(705,7,'分辨率','smart-stat-report-client.asp?TypeTo=4','');
		d.add(706,7,'颜色','smart-stat-report-client.asp?TypeTo=5','');
		d.add(707,7,'访问者地区','smart-stat-report-area.asp?TypeTo=10','');
		d.add(708,7,'Alexa工具条','smart-stat-report-client.asp?TypeTo=9','');

		d.add(8,0,'受欢迎度','','I','','','',false);
		d.add(801,8,'回头客','smart-stat-report-client.asp?TypeTo=6','');
		d.add(802,8,'浏览深度','smart-stat-report-client.asp?TypeTo=8','');

		d.add(9,0,'内容分析','','I','','','',false);
		d.add(901,9,'来路','smart-stat-report-operativesystems.asp?TypeTo=0','');
		d.add(902,9,'入口网址','smart-stat-report-operativesystems.asp?TypeTo=1','');
		d.add(903,9,'页面浏览','smart-stat-report-operativesystems.asp?TypeTo=4','');

		d.add(10,0,'系统管理','','I','','','',false);
		d.add(1001,10,'设置','smart-stat-settings.asp','');
		d.add(1002,10,'获取代码','smart-stat-take.asp','');
		d.add(1003,10,'站点初始化','smart-stat-initialization.asp','');
		d.add(1004,10,'退出','logout.asp','');

		document.write(d);

		//-->
</script>
<p><a href="javascript: d.openAll();">Expand all</a> | <a href="javascript: d.closeAll();">Collapse all</a></p>
</div>
</body>
</html>