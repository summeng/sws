
    <TABLE width=960 border=0 align="center" cellPadding=0 cellSpacing=0 background="../images/sdimg/<%=sdfg%>/btmbg.png"  style="BORDER-COLLAPSE: collapse">
      <TBODY>
        <TR> 
          <TD height=20 width="758"  colspan="3"> </TD>
        </TR>
        <TR><TD width="30"> </TD>          
        <TD width="920" class="p9black">&nbsp;&nbsp;&nbsp;&nbsp;������ <b><%=rs("sdname")%></b>���̣�����<b><%=rs("name")%></b>����[<%=webname%>]����ƽ̨���п���ģ�������Ϣ�ĺϷ��Ժ���ʵ����<b><%=rs("name")%></b>����[<%=webname%>]�����κε�������������<b><%=rs("sdname")%></b>������Ϣ������<b><%=rs("name")%></b>�ṩ�ķ���Ҫ���м�����α�ͺϷ��ԡ��������ɵ��κ���ʧ[<%=webname%>]�Ų������಻�е��κη������Ρ�<br>
      </TD><TD width="30"> </TD>
        </TR>
        <TR> <%conn.execute "UPDATE DNJY_user SET sdhits=sdhits+1 WHERE username='"&com&"'" '�������%>
          <TD height=30 width="920" colspan="3" align="center"><%=rs("sdname")%> <img src="../images/sdimg/ws.jpg" align="middle" border=0 alt="��Ϻ���"> <a href="/<%=strInstallDir%>"><%=webname%></a></TD>
        </TR>
        <TR> <%conn.execute "UPDATE DNJY_user SET sdhits=sdhits+1 WHERE username='"&com&"'" '�������%>
          <TD height=30 width="920" colspan="3" align="center">����[<font face="Haettenschweiler"><b><%=rs("tong")%></b></font>]�˷���</TD>
        </TR>
      </TBODY>
</TABLE>
</body>
<%Call CloseRs(rs)
Call CloseDB(conn)%>
