<!--#include file="conn.asp"-->
<!--#include file="config.asp"-->
<%
'=====================================
'���칩����Ϣ��վ����ϵͳ
'��������Ƽ������ҳ�Ʒ
'�ͷ�����http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>
<%
Dim Action,citynotice,activity,id
Call OpenConn
If request.cookies("administrator")="" Then
activity=" and activity=1"
Else
activity=""
End If
id=clng(request.querystring("id"))
	set rs=server.createobject("adodb.recordset")
	sql="select * from DNJY_announce where id="&id&""&activity&""
	rs.open sql,conn,1,1
	if rs.eof Then
    Response.Write "<font color=000000>�������Ի��޲鿴Ȩ��</font>"
	Response.end
    else%>
<HTML>
 <HEAD>
  <TITLE>��վ����</TITLE>
 </HEAD>
 <BODY>
<LINK rel="stylesheet" type="text/css" href="../css/style.css">
  <table width="98%" height="98%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#799AE1">
    <tr>
      <td height="20" bgcolor="#799AE1"><div align="center"><font color="#FFFFFF" style="font-size:14px">�� վ �� �� </font></div></td>
    </tr>
    <tr>
      <td bgcolor="#FFFFFF"><br />
        <table width="98%" height="98%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#D6DFF7">
            <tr>             
              <td height="50" align="center" bgcolor="#E3EBf0"><%=rs("title")%></td>
          </tr>
            <tr>             
              <td bgcolor="#E3EBF9" valign="top"><div style="margin-top:10px;margin-bottom:10px;margin-left:10px;margin-right:10px;"><%=rs("content")%></div></td>
          </tr>
            <tr>             
              <td height="40" align="right" bgcolor="#E3EBF9">����ʱ�䣺<%=rs("addtime")%></td>
          </tr>
          </table>
      <br/></td>
    </tr>
  </table>
<%sql="update DNJY_announce set hits=hits+1 where id="&id
conn.execute sql
End if
Call CloseRs(rs)
Call CloseDB(conn)%>
 </BODY>
</HTML>