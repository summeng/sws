<!--#include file="../conn.asp"-->
<!--#include file="../config.asp"-->
<!--#include file="../SqlIn.Asp"-->
<!--#include file=cookies.asp-->
<%
'=====================================
'���칩����Ϣ��վ����ϵͳ
'��������Ƽ������ҳ�Ʒ
'�ͷ�����http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>
<title>�༭�ظ�</title>
<link rel="stylesheet" type="text/css" href="../css/css.css">
<BODY>
<script language=javascript src=../Include/mouse_on_title.js></script>
<table width="95%" border="0" cellspacing="1" cellpadding="1" bordercolor="6699cc" align="center">
  <tr> 
    <td bgcolor="#FFFFFF">
	<table border="0" cellpadding="0" cellspacing="0" width="100%">
        <tr> 
          <td> 
            
            <table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr> 
                <td><%dim id,nr,referer
				referer = Request.ServerVariables("HTTP_REFERER")
id=request("id")
if id="" then
	response.write "<script language='javascript'>"
	response.write "alert('�޴˻ظ���ţ��뵥����ȷ�������أ�');"
	response.write "location.href='"&(referer)&"';"
	response.write "</script>"
	response.end
end if
Call OpenConn
	'�޸Ļظ�����
if request("send")="ok" then
	set rs=server.createobject("adodb.recordset")
	sql = " select * from DNJY_reply  where id="&id
	rs.open sql,conn,1,3
		if not (rs.eof and rs.bof) then
		rs("neirong")=request.form("nr")
		rs("hfyz")=request("hfyz")
		rs.update
		end if

	rs.close
    response.Write "<script language=javascript>alert('�޸ĳɹ�!');location.replace('"&(referer)&"');</script>"
		response.end
end if

	'��ʾ��ϸ����
	set rs = server.createobject("adodb.recordset")
	sql = "select * from DNJY_reply where id="&id
	rs.open sql,conn,1,1

		if rs.eof and rs.bof then 
		response.write "<script language='javascript'>"
		response.write "alert('�޴˻ظ����뵥����ȷ�������أ�');"
		response.write "location.href='"&(referer)&"';"
		response.write "</script>"
		response.end
		end if

		if not (rs.eof and rs.bof) then 
			nr=replace(rs("neirong"),"<BR>",vbCRLF)
		%>
   <table width="98%" border="1" cellpadding="3" bordercolor="#333333" style="border-collapse: collapse;">
	<tr class=backs><td colspan=2 class=td height=18>�ظ���Ϣ����</td></tr>
		 <form name="repl" method="post" action='chkcomment.asp?id=<%=rs("id")%>'>
		 <tr><TD align="right" WIDTH=20% height=15>�ظ���IP��ַ</TD><td><a href='http://www.ip138.com/ips.asp?ip=<%=rs("ip")%>' target=_blank alt=�鿴��Դ><%=rs("ip")%></a></td></tr>
		 <tr><TD align="right" >�ظ�����</TD><td><%=rs("hfsj")%></td></tr>		 
		 <tr><TD align="right" >�ظ�������</TD><td><%=rs("username")%></td></tr>
		
		 <tr><TD align="right" >�ظ�����</TD><td><textarea style="overflow:auto" name="nr" cols="60" rows="8"><%=nr%></textarea></td></tr>
		<tr><TD align="right" >�Ƿ�����</TD><td><input type="radio" name="hfyz" value="1" <%if rs("hfyz")="1" then%>checked<%end if%>>
			  ����<input type="radio" name="hfyz" value="0" <%if rs("hfyz")="0" then%>checked<%end if%>>
			  ���� </td></tr>
			<TR><TD align="right" >&nbsp;<INPUT TYPE="hidden" name=send value=ok></TD><TD>
				<input type="submit" name="action" value=" �� �� "></TD></TR>
			</form></TABLE>
		<%
		end if	
	rs.close
	set rs=nothing
	Call CloseDB(conn)%>
	</td>
              </tr>
             
            </table></td>
        </tr>
      </table> </td>
  </tr>
</table>

