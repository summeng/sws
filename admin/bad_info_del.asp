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
<%
dim id,xid,i,u,str1,str2,tupian,objFSO,fileExt,sql1,rs1,xxid,bdback
bdback = Request.ServerVariables("HTTP_REFERER")
xid=trim(request("xid"))
id=trim(request("selectedid"))
if trim(id)="" then
response.write "<script language=JavaScript>" & chr(13) & "alert('û��ѡ���¼��');" & "history.back()" & "</script>"
response.end
end if
str1=split(xid,",")
str2=split(id,",")
Call OpenConn
set rs1 = Server.CreateObject("ADODB.RecordSet")
if xid<>"" then
for i=0 to ubound(str1)
sql1="delete from [DNJY_Bad_info] where id="&id
rs1.open sql1,conn,1,3
next
else

for u=0 to ubound(str2)
sql1="delete from [DNJY_Bad_info] where id="&trim(str2(u))
rs1.open sql1,conn,1,3
next
end if

response.write "ɾ���ٱ���Ϣ�ɹ���"
response.redirect (bdback)
response.end
Call CloseRs(rs)
Call CloseRs(rs1)
Call CloseDB(conn)
%>
