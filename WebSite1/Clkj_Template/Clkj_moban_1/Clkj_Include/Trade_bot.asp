<div class="sem-bot">
    	<ul>
        <li><%=bottom%> <a href="sitemap.xml">Sitemap</a></li>
        <%if clkj_open_key=1 and clkj_new_les="index" then%>
        <li><%=clkj_bottom_key%></li>
        <%end if%>
         <%if clkj_open_link=1 and clkj_new_les="index" then%>
        <li><%=links%></li>
        <%end if%>
        <li><%=clkj_bottom_m%></li></ul> 
</div>
<%call allclose()%>