<div class="sem-mid-left" id="lefts">
<%if clkj_new_lei<>"" and clkj_new_les="" then%>

<div class="sem-mid-left-bt"><%=clkj_new_lei%></div>
            <div class="sem-mid-left-1">
            	<div class="sem-mid-left-1-1">
                	<ul>
					 <%=Abouttitle%>

                    </ul>

                </div>

            </div>
            <div class="cb"></div>
			
<%elseif clkj_new_lei="About us" and clkj_new_les="index" then%>
<div class="sem-mid-left-bt"><%=clkj_new_lei%></div>
            <div class="sem-mid-left-1">
            	<div class="sem-mid-left-1-1">
                	<ul>
					 <%=Abouttitle%>

                    </ul>

                </div>

            </div>
            <div class="cb"></div>

<%elseif clkj_page_id="gallery" then%>
    <script language="javascript">
    function showsubmenu(sid)
    {
    whichEl = eval("submenu" + sid);
    imgmenu = eval("imgmenu" + sid);
    if (whichEl.style.display == "none")
    {
    eval("submenu" + sid + ".style.display=\"\";");
    imgmenu.background="./Clkj_Template/Clkj_moban_1/Clkj_images/inc_3.gif";
    }
    else
    {
    eval("submenu" + sid + ".style.display=\"none\";");
    imgmenu.background="./Clkj_Template/Clkj_moban_1/Clkj_images/inc_4.gif";
    }
    }
    </script>


            	<div class="sem-mid-left-bt">GALLERYS</div>
                <div class="sem-mid-left-1">
                	<div class="sem-mid-left-1-1">
                    	<ul>
    					<% call gallery_leftnav()%>

                        </ul>

                    </div>

                </div>
                <div class="cb"></div>
<%else%>
<script language="javascript">
function showsubmenu(sid)
{
whichEl = eval("submenu" + sid);
imgmenu = eval("imgmenu" + sid);
if (whichEl.style.display == "none")
{
eval("submenu" + sid + ".style.display=\"\";");
imgmenu.background="./Clkj_Template/Clkj_moban_1/Clkj_images/inc_3.gif";
}
else
{
eval("submenu" + sid + ".style.display=\"none\";");
imgmenu.background="./Clkj_Template/Clkj_moban_1/Clkj_images/inc_4.gif";
}
}
</script>


        	<div class="sem-mid-left-bt">Categories</div>
            <div class="sem-mid-left-1">
            	<div class="sem-mid-left-1-1">
                	<ul>
					<% call leftnav()%>

                    </ul>

                </div>

            </div>
            <div class="cb"></div>
<%end if%>
            <div class="sem-mid-left-bt">News</div>
            <div class="cb"></div>
            <div class="sem-mid-left-2">
            	<ul>
			<%=News%>
                </ul>
            </div>
            <div class="cb"></div>
            <div class="sem-mid-left-3">
            	<div class="sem-mid-left-3-1">Our newsletter</div>
                <div class="cb"></div>
                <div class="sem-mid-left-3-2">
                <form name="f1" method="post" action="Clkj_Inc/clkj_message.asp?class=mails">
                	<ul>
		<li><input type="text" name="mails" style="font-size:11px; color:#666; height:20px; line-height:20px; padding:3px; width:170px;" value="Enter your E-mail"  autocomplete="off" onfocus="if(this.value!='Enter your E-mail'){this.style.color='#000000'}else{this.value='';this.style.color='#000000'}" onblur="if(this.value==''){this.value='Enter your E-mail';this.style.color='#666666'}"/></li>
        <li><input type="image" src="./Clkj_Template/Clkj_moban_1/Clkj_images/submit.gif" style="float:right;" /></li>


                    </ul>
             </form>
                </div>


            </div>
            </div>
