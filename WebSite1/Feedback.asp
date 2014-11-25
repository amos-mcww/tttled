<!--#include file="Clkj_Inc/clkj_inc.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<meta http-equiv="X-UA-Compatible" content="IE=EmulateIE7" />
<title>Feedback - <%=clkj_co_name%></title>
<meta name="keywords" content="<% call navdes(31,0)%>" />
<meta name="description" content="<% call navdes(31,1)%>" />
<link href="Clkj_Template/Clkj_moban_1/Clkj_css/Trade.css" rel="stylesheet" type="text/css" />
</head>
<script language="javascript">
 function E()
       {
         var temp = document.getElementById("mail");
          var myreg = /^([a-zA-Z0-9]+[_|\-|\.]?)*[a-zA-Z0-9]+@([a-zA-Z0-9]+[_|\-|\.]?)*[a-zA-Z0-9]+\.[a-zA-Z]{2,3}$/;
          if(!myreg.test(temp.value))
          {
               alert('Please enter a valid E-mail！');
			   temp.value="";	   
               return false;
          }
      }

</script>
<body>
<div class="sem">
<!--#include file="Clkj_Template/Clkj_moban_1/Clkj_Include/Trade_top.asp"-->
<div class="sem-tag"><a href="/">Home</a> » Feedback</div>
<div class="cb"></div>
    <div class="sem-mid">
 <!--#include file="Clkj_Template/Clkj_moban_1/Clkj_Include/Trade_left.asp"-->
        <div class="sem-mid-cont">
        	<div class="sem-mid-cont-bt">Feedback</div>
            <div class="cb"></div>
            <div class="sem-mid-cont-1"> <table width="100%" border="0" cellpadding="0" cellspacing="0" style="border:#FFFFFF solid 1px;">
              <form name="message" action="Clkj_Inc/clkj_message.asp?class=message" method="post">
                <tr>
                  <td width="140" height="120" align="right" valign="middle" class="border1"><span class="cl-1">*</span> Content:</td>
                  <td width="416" align="left" valign="middle" class="border2">&nbsp;
                    <textarea name="content" cols="50" rows="7" id="content"></textarea></td>
                </tr>
                <tr>
                  <td height="30" align="right" valign="middle" class="border1">Name:</td>
                  <td align="left" valign="middle" class="border2">&nbsp;
                    <input name="Name" type="text" id="Name" size="20" /></td>
                </tr>
                <tr>
                  <td height="30" align="right" valign="middle" class="border1"><span class="cl-1">*</span> E-mail:</td>
                  <td align="left" valign="middle" class="border2">&nbsp;
                    <input name="mail" type="text" id="mail" size="30" onBlur="E();" />
                  </td>
                </tr>
                <tr>
                  <td height="30" align="right" valign="middle" class="border1">Mobile:</td>
                  <td align="left" valign="middle" class="border2">&nbsp;
                    <input name="Phone" type="text" id="Phone" size="20" /></td>
                </tr>
                <tr>
                  <td height="35" align="right" valign="middle" class="border1"><!--DWLayoutEmptyCell-->&nbsp;</td>
                  <td align="left" valign="middle" class="border2">&nbsp;
                    <input type="submit" name="Submit" value="Submit" />
                    </td>
                </tr>
              </form>
            </table></div>
            <div class="cb"></div>
        </div>
    
    </div>
    <div class="cb"></div>
 <!--#include file="Clkj_Template/Clkj_moban_1/Clkj_Include/Trade_bot.asp"-->
</div>
</body>
</html>