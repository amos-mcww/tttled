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
                  <td width="140" height="30" align="right" valign="middle" class="border1"><span class="cl-1">*</span>Title:</td>
                  <td width="416" align="left" valign="middle" class="border2">&nbsp;
                    <input name="title" type="text" id="title" size="40" value="<%=clkj_prtitlefeed%>"/></td>
                </tr>
                <tr>
                  <td height="120" align="right" valign="middle" class="border1"><span class="cl-1">*</span> Content:</td>
                  <td align="left" valign="middle" class="border2">&nbsp;
                    <textarea name="content" cols="50" rows="7" id="content"></textarea></td>
                </tr>
                <tr>
                  <td height="30" align="right" valign="middle" class="border1"> Company:</td>
                  <td align="left" valign="middle" class="border2">&nbsp;
                    <input name="Company" type="text" id="Company" size="40" /></td>
                </tr>
                <tr>
                  <td height="30" align="right" valign="middle" class="border1"><span class="cl-1">*</span> Name:</td>
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
                  <td height="30" align="right" valign="middle" class="border1"><span class="cl-1">*</span> Phone:</td>
                  <td align="left" valign="middle" class="border2">&nbsp;
                    <input name="Phone" type="text" id="Phone" size="20" /></td>
                </tr>
                <tr>
                  <td height="30" align="right" valign="middle" class="border1"> Fax:</td>
                  <td align="left" valign="middle" class="border2">&nbsp;
                    <input name="Fax" type="text" id="Fax" size="20" /></td>
                </tr>
                <tr>
                  <td height="30" align="right" valign="middle" class="border1"><span class="cl-1">*</span> Country/Region:</td>
                  <td align="left" valign="middle" class="border2">&nbsp;
                  <input type="text" name="Region" id="Region" />
                  </td>
                </tr>
                <tr>
                  <td height="30" align="right" valign="middle" class="border1">Home Page:</td>
                  <td align="left" valign="middle" class="border2">&nbsp;
                    <input name="Home" type="text" id="Home" size="40" /></td>
                </tr>
                <tr>
                  <td height="35" align="right" valign="middle" class="border1"><span class="cl-1">*</span> Code</td>
                  <td align="left" valign="middle" class="border2"> &nbsp;
                  <input type="text" size="8" name="yzm" style="width:40px;" > <img src="Clkj_Inc/Clkj_code.asp" alt="验证码" onclick="this.src='Clkj_Inc/Clkj_code.asp?rnd=' + Math.random();" /> </td>
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