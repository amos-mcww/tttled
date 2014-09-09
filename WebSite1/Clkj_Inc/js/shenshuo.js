// JavaScript Document

function showsubmenu(sid)
{
whichEl = eval("submenu" + sid);
//imgmenu = eval("imgmenu" + sid);
if (whichEl.style.display == "none")
{
eval("submenu" + sid + ".style.display=\"\";");
//imgmenu.background="./Clkj_Template/Clkj_moban_1/Clkj_images/inc_3.gif";
}
else
{
eval("submenu" + sid + ".style.display=\"none\";");
//imgmenu.background="./Clkj_Template/Clkj_moban_1/Clkj_images/inc_4.gif";
}
}
