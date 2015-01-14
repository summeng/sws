
    <TABLE width=960 border=0 align="center" cellPadding=0 cellSpacing=0 background="../images/sdimg/<%=sdfg%>/btmbg.png"  style="BORDER-COLLAPSE: collapse">
      <TBODY>
        <TR> 
          <TD height=20 width="758"  colspan="3"> </TD>
        </TR>
        <TR><TD width="30"> </TD>          
        <TD width="920" class="p9black">&nbsp;&nbsp;&nbsp;&nbsp;声明： <b><%=rs("sdname")%></b>店铺，是由<b><%=rs("name")%></b>利用[<%=webname%>]交易平台自行开办的，店铺信息的合法性和真实性由<b><%=rs("name")%></b>负责，[<%=webname%>]不作任何担保。您若采用<b><%=rs("sdname")%></b>店铺信息并接受<b><%=rs("name")%></b>提供的服务，要自行鉴别真伪和合法性。对因此造成的任何损失[<%=webname%>]概不负责，亦不承担任何法律责任。<br>
      </TD><TD width="30"> </TD>
        </TR>
        <TR> <%conn.execute "UPDATE DNJY_user SET sdhits=sdhits+1 WHERE username='"&com&"'" '点击计数%>
          <TD height=30 width="920" colspan="3" align="center"><%=rs("sdname")%> <img src="../images/sdimg/ws.jpg" align="middle" border=0 alt="真诚合作"> <a href="/<%=strInstallDir%>"><%=webname%></a></TD>
        </TR>
        <TR> <%conn.execute "UPDATE DNJY_user SET sdhits=sdhits+1 WHERE username='"&com&"'" '点击计数%>
          <TD height=30 width="920" colspan="3" align="center">已有[<font face="Haettenschweiler"><b><%=rs("tong")%></b></font>]人访问</TD>
        </TR>
      </TBODY>
</TABLE>
</body>
<%Call CloseRs(rs)
Call CloseDB(conn)%>
