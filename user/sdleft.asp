<%dim rs0
set rs0=server.createobject("adodb.recordset")
set rs0=conn.execute("select count(id) from [DNJY_info] where username='"&com&"'")
n=rs0(0)
rs0.close
set rs0=Nothing%>
	<table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF" id="line">
      <tr>
        <td width="100%" valign="top"><table width="100%" border="0" cellpadding="0" cellspacing="0">
          <tr>
            <td><div id="lt">��˾����</div></td>
          </tr>
        </table>
          <table width="100%" border="0" cellspacing="0" cellpadding="0" style="margin:3px auto">
            <tr>
              <td><div id="lefti"> <ul><marquee onmouseover="this.stop()" onmouseout="this.start()" direction="up" scrollamount="2" scrolldelay="0" behavior="scroll" height="80" align="center"><%if rs("comgg")="" then%>��ӭ����<%=rs("sdname")%>��ҳ<%else%><%=rs("comgg")%><%End if%></marquee></ul></div></td>
            </tr>
          </table></td>
      </tr>
    </table>
	<table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF" id="line">
      <tr>
        <td width="100%" valign="top"><table width="100%" border="0" cellpadding="0" cellspacing="0">
          <tr>
            <td><div id="lt">�������</div></td>
          </tr>
        </table>
          <table width="100%" border="0" cellspacing="0" cellpadding="0" style="margin:3px auto">
            <tr>
              <td><div id="lefti"><ul><table width=180 border=0 cellpadding=0 cellspacing=0><tr><td width=150>�̼ҵȼ���VIP��Ա<br>&nbsp;&nbsp;&nbsp;&nbsp; ������<%=rs("name")%>(<%=rs("username")%>)<br>&nbsp;&nbsp;&nbsp;&nbsp; ���֣�<%=rs("jf")%><br>&nbsp;&nbsp; ��Ʒ����<%=n%><br>&nbsp;����ʱ�䣺<%=formatdatetime(rs("zcdata"),1)%><br></td></tr><tr><TD height=1 colspan=2 background=images/sdimg/<%=sdfg%>/left_line01.gif><IMG height=1 src=images/sdimg/<%=sdfg%>/1x1_pix.gif width=10></TD></TR></table></ul></div></td>
            </tr>
          </table></td>
      </tr>
    </table>

	<table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF" id="line">
      <tr>
        <td width="100%" valign="top"><table width="100%" border="0" cellpadding="0" cellspacing="0">
          <tr>
            <td><div id="lt">���Ų�Ʒ</div></td>
          </tr>
        </table>
          <table width="100%" border="0" cellspacing="0" cellpadding="0" style="margin:3px auto">
            <tr>
              <td><div id="lefti"><ul>
 <%
dim ae1,rsxx
ae1=0
set rsxx = Server.CreateObject("ADODB.RecordSet")
sql="select * from [DNJY_info] where yz=1  and username='"&com&"' order by llcs "&DN_OrderType&",ID "&DN_OrderType&"" 	
rsxx.open sql,conn,3,3
do while not rsxx.eof
%>
           <table width="100%" border="0" cellpadding="0" cellspacing="0">
             <tr>
               <td>��<a  href="../<%=x_path("",rsxx("id"),"",rsxx("Readinfo"))%>" target="_blank"><%=CutStr(rsxx("biaoti"),24,"...")%></a></td>
             </tr>
           </table>
    <%
                ae1=ae1+1
                if ae1>=10 then exit do
                rsxx.movenext
                loop
                rsxx.close
                set rsxx=nothing
                %>
			  </ul></div></td>
            </tr>
          </table></td>
      </tr>
    </table>

	<table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF" id="line">
      <tr>
        <td width="100%" valign="top"><table width="100%" border="0" cellpadding="0" cellspacing="0">
          <tr>
            <td><div id="lt">��ϵ����</div></td>
          </tr>
        </table>
          <table width="100%" border="0" cellspacing="0" cellpadding="0" style="margin:3px auto">
            <tr>
              <td><div id="lefti"><ul>
			   �绰��<%=rs("dianhua")%><br> �ֻ���<%=rs("CompPhone")%><br> ���棺<%=rs("fax")%><br> E-mail��<%=rs("email")%><br> ��ַ��<a href="http://<%=rs("http")%>" target="_blank"><%=rs("http")%></a><br>�ʱࣺ<%=rs("youbian")%><br>��ַ��<%=rs("dizhi")%><br><a href="#" ONCLICK="window.open('user_mail.asp?username=<%=rs("username")%>&email=<%=rs("email")%>&mailzt=<%=rs("sdname")%>','Sample','toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=yes,resizable=no,copyhistory=no,width=650,height=300,left=300,top=20')">����ҷ��ʼ�</a><br><a href="../messh.asp?name=<%=rs("username")%>&<%=C_ID%>">��վ�ڶ��Ÿ���</a><br><%if rs("qq")<>"" then %> <a target=blank href=tencent://message/?Uin=<%=rs("qq")%>&Menu=yes><img border="0" SRC=http://wpa.qq.com/pa?p=1:<%=rs("qq")%>:7 alt="��QQ��������Ϣ"></a><%end if%>
			  </ul></div></td>
            </tr>
          </table></td>
      </tr>
    </table>
	<table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF" id="line">
      <tr>
        <td width="100%" valign="top"><table width="100%" border="0" cellpadding="0" cellspacing="0">
          <tr>
            <td><div id="lt">��������</div></td>
          </tr>
        </table>
          <table width="100%" border="0" cellspacing="0" cellpadding="0" style="margin:3px auto">
            <tr>
              <td><div id="lefti"><ul>
<%Dim rsl
i=0
set rsl = Server.CreateObject("ADODB.RecordSet")
sql="select * from DNJY_comlink where username='"&com&"'" 	
rsl.open sql,conn,1,1
do while not rsl.eof
%>
           <table width="100%" border="0" cellpadding="0" cellspacing="0">
             <tr>
               <td>��<a  href="<%=rsl("url")%>" target="_blank"><%=CutStr(rsl("web"),24,"...")%></a></td>
             </tr>
           </table>
    <%
                i=i+1
                if i>=8 then exit do
                rsl.movenext
                loop
                rsl.close
                set rsl=nothing
                %>
			  </ul></div></td>
            </tr>
          </table></td>
      </tr>
    </table>
