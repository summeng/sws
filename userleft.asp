<!--#include file=usercookies.asp-->
<%
'=====================================
'���칩����Ϣ��վ����ϵͳ
'��������Ƽ������ҳ�Ʒ
'�ͷ�����http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>
<%
dim sdname,mpname,mpfg,sj,vip
username=request.cookies("dick")("username")
Call OpenConn
set rs=server.createobject("adodb.recordset")
sql = "select * from [DNJY_user] where username='"&username&"'"
rs.open sql,conn,1,1
sdname=rs("sdname")
mpname=rs("mpname")
mpfg=rs("mpfg")
vip=rs("vip")
if rs("vipdq")<>"" Then
sj=DateDiff("d",now(),""&rs("vipdq")&"")
end If
%>
<meta http-equiv="Content-Language" content="zh-cn">
<meta name="GENERATOR" content="Microsoft FrontPage 6.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title><%=title%>-�ҵ�״̬</title>
<style type="text/css">
<!--
body {
	margin-left: 0px;
	margin-top: 0px;
}
-->
</style>

<table width="214" border="0" cellPadding="0" cellSpacing="0" borderColor="#111111" style="border-collapse:collapse">
<td width="215" height="360" vAlign="top">
      <table width="210" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td  height="7"><a href="<%=adduserinfo%>?<%=CT_ID%>"><img src="img/addxx.jpg" width=214 height=58 border=0></a></td>
        </tr>
      </table>
      <TABLE cellSpacing=0 cellPadding=0 border=0 class="ty1">        
        <TR>
          <TD align=middle  class="ty3" height=30><b>�� Ա �� ��</b></TD></TR>
        <TR>
          <TD width="210" height=200 valign="top">	
	<table width="80%" border="0" align="center" cellpadding="2" cellspacing="0" class="font_10_e_black">
      <tr>
        <td height="10" colspan="2"><%response.write "��Ա<font color=#FF0000>"&username&"</font>�Ĺ�������<br>"%>
  <%if vip>0 then
  response.write "�ҵĻ�Ա����<font color=#FF0000>VIP��Ա</font><br>"
if sj>0 then
response.Write "�ҵ�VIP��Ա�ʸ�<font color=#FF0000> ʣ��"&sj&"</b>��</font><br>"
elseif sj=0 then
response.Write "�ҵ�VIP��Ա�ʸ�<font color=""#414141"">���յ���</font><br>"
elseif sj<0 then
response.Write "�ҵ�VIP��Ա�ʸ�<font color=""#FF0000""> �Ѿ�����</font><br>"
end If
else
response.write "�ҵĻ�Ա����<font color=#FF0000>��ͨ��Ա</font><br>"
response.write "�������������Ϊ<a href='upgradevip.asp?"&C_ID&"'><font color=#0000ff>VIP��Ա</font></a>"
end if%> </td>
      </tr>

      <tr>
        <td align="center" valign="top">[ <a href="user.asp?<%=C_ID%>">�ҵ�״̬</a></td>
        <td width="84"><div align="center"><a href="user_cwmx.asp?<%=C_ID%>" >������ϸ</a> ]</div></td>
      </tr>
      <tr>
        <td align="center" valign="top"><div align="center">[ <a href="user_zhcz.asp?<%=C_ID%>">�ʻ���ֵ</a></div></td>
        <td width="84"><div align="center"><a href="user_order.asp?<%=C_ID%>">��������</a> ]</div></td>
      </tr>

      <tr>
        <td align="center" valign="top"><div align="center">[ <a href="user_props.asp?<%=C_ID%>" >�������</a></div></td>
        <td width="84"><div align="center"><a href="props.asp?<%=C_ID%>">װ��ת��</a> ]</div></td>
      </tr>
      <tr>
        <td align="center" valign="top"><div align="center">[ <a href="user_xxgl.asp?<%=C_ID%>">��Ϣ����</a></div></td>
        <td width="84"><div align="center"><a href="user_hfgl.asp?<%=C_ID%>">����ظ�</a> ]</div></td>
      </tr>
      <tr>
        <td align="center" valign="top"><div align="center">[ <a href="UserNews_AddSelClass.asp?<%=C_ID%>">��������</a></div></td>
        <td width="84"><div align="center"><a href="UserNews_ModSelClass.asp?<%=C_ID%>">��������</a> ]</div></td>
      </tr>
       <tr>
        <td align="center" valign="top"><div align="center">[ <a href="messs.asp?<%=C_ID%>">���ն���</a></div></td>
        <td width="84"><div align="center"> <a href="messh.asp?<%=C_ID%>">���Ͷ���</a> ]</div></td>
      </tr>	  
      <tr>
        <td align="center" valign="top"><div align="center">[ <% if sdname<>"" then %><a href="user/com.asp?com=<%=rs("username")%>&<%=C_ID%>" target="_blank">�ҵ�����</a><%Else%><font color="#FF0000">��������</font><%end if%></div></td>
        <td width="84"><div align="center"><% if mpname<>"" then %><a href="company.asp?username=<%=rs("username")%>&<%=C_ID%>" target="_blank">�ҵ���Ƭ</a><%Else%><font color="#FF0000">������Ƭ</font><%end if%> ]</div></td>
      </tr>   
      <tr>
        <td align="center" valign="top"><div align="center">[ <a href="usersdeditzl.asp?<%=C_ID%>">�������</a></div></td>
        <td width="84"><div align="center"> <a href="usersdeditzl.asp?<%=C_ID%>">������Ƭ</a> ]</div></td>
      </tr> 
      <tr>
        <td align="center" valign="top"><div align="center">[ <a href="certificate.asp?<%=C_ID%>">֤�չ���</a></div></td>
        <td width="84"><div align="center"><% if sdname<>"" Or mpname<>"" then %><a href="comgg.asp?<%=C_ID%>">�����־</a><%else%><font color="#FF0000">�����־</font><%end if%> ]</div></td>
      </tr> 	  
 
      <tr>
        <td align="center" valign="top"><div align="center">[ <a href="user_data.asp?<%=C_ID%>">�޸�����</a></div></td>
        <td width="84"><div align="center"><a href="user_pass.asp?<%=C_ID%>">�޸�����</a> ]</div></td>
      </tr>

       <tr>
        <td align="center" valign="top"><div align="center">[ <a href="user_favorites.asp?<%=C_ID%>">��Ϣ�ղ�</a></div></td>
        <td width="84"><div align="center"> <a href="user_favshops.asp?<%=C_ID%>">�����ղ�</a> ]</div></td>
      </tr>
      <tr>
        <td align="center" valign="top"><div align="center">[ <a href="#" ONCLICK="window.open('user/contribute.asp','Sample','toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=yes,resizable=no,copyhistory=no,width=780,height=700,left=300,top=20')">��ҪͶ��</a></div></td>
        <td width="84"><div align="center"><a href="manuscript.asp?<%=C_ID%>">��Ͷ���</a> ]</div></td>
      </tr>
      <tr>
        <td align="center" valign="top"><div align="center">[ <a href="Message_board.asp?mybook=ok&<%=C_ID%>">�ҵ�����</a></div></td>
        <td width="84"><div align="center"><a href="usermail_loglist.asp?isse=0&<%=C_ID%>">�ʼ���־</a> ]</div></td>
      </tr>
      <tr>
        <td align="center" valign="top"><div align="center">[ <a href="user_link.asp?<%=C_ID%>">��������</a></div></td>
        <td width="84"><div align="center"><a href="userrecommend.asp">�Ƽ���Ա</a> ]</div></td>
      </tr>
	  <%If  isnull(rs("QQopenid")) Then
	  session("binding")="yes"%>
      <tr>
        <td align="center" valign="top" colspan="2"><div align="center"><a title="ʹ��QQ�ſ��ٵ�¼" href="/<%=strInstallDir%>api/qq/redirect_to_login.asp"  target="_blank"><font color=#FF0000><b>��QQ�ſ�ݵ�¼</b></font></a></td>
      </tr>
	  <%End if%>
      <tr>
        <td height="10" colspan="2"><div align="center"><a href="user_sk.asp"><font color=#FF0000><b>��վ�տ��ʺ�</b></font></a></div></td>
      </tr>
      <tr>
        <td align="center" valign="top" colspan="2"><div align="center"><a href="user/userout.asp?<%=C_ID%>"><b>��ȫ�˳�</b></a></td>
      </tr>    </table></td>
    
  </tr>

</table>
