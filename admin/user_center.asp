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
<%Call OpenConn
dim id
id=replace(trim(request("id")),"'","��")
set rs = Server.CreateObject("ADODB.RecordSet")
sql="select id,username from [DNJY_user] where id="&cstr(id) 
rs.open sql,conn,1,1
if rs.eof then
response.write "<li>��û���û����ݣ�"
response.end
end if
Response.cookies("dick")("username")=rs("username")
Response.cookies("dick")("domain")=Request.ServerVariables("SERVER_NAME")
Call CloseRs(rs)
Call CloseDB(conn)
response.redirect "../user_xxgl.asp"
%>
