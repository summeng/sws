<!--#include file="../conn.asp"-->
<!--#include file="../SqlIn.Asp"-->
<!--#include file=../config.ASP-->
<!--#include file="../Include/head.asp"-->
<!--#include file="../Include/Function.asp"-->
<!--#include file="../Include/label.asp"-->
<%
'=====================================
'���칩����Ϣ��վ����ϵͳ
'��������Ƽ������ҳ�Ʒ
'�ͷ�����http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>
<%
dim objFSO,fileExt,sql1,rs1,sql2,delpas,username
id=request("id")
username=request.cookies("dick")("username")
If request.cookies("dick")("username")="" then
response.Write "����������ȷ�ϻ�Ա��¼���ܲ�ѯ��"
response.end
End if
Call OpenConn
if request("key")="cs" And username<>"" then
%>
<form method="POST" action="?username=<%=username%>&cskey=ok&id=<%=id%>">
<p align="center">��ѯ����Ϣ�������ֻ�����<br>���������ʻ��Ͽ۳�<%=jf_ck%>�����</p>
<p align="center">Ŀǰ���ʻ���<%=conn.Execute("Select jf From DNJY_user Where username='"&username&"'")(0)%>�����</p>
<p align="center"> <input type="submit" value="ȷ����ѯ" name="B1"></p>
<p align="center"><a href="../help.asp?id=1&<%=C_ID%>" target=_blank>׬ȡ���ַ���</a></p>
<!--<p align="center"><a href="../help.asp?id=0&<%=C_ID%>" target=_blank>׬ȡ���ַ���</a></p><!--������-->
</form>
<%Elseif request("cskey")="ok" And conn.Execute("Select jf From DNJY_user Where username='"&username&"'")(0)>=jf_ck then
set rs=server.createobject("adodb.recordset")
sql="select CompPhone,name from [DNJY_info] where id="&cstr(id)
rs.open sql,conn,1,1
if not(rs.eof or rs.bof) then
conn.execute "UPDATE DNJY_user SET jf=jf-"&jf_ck&" WHERE username='"&username&"'" 'ͬʱ���û����ݿ���ٻ���
response.Write "<TABLE align=""center""><TR><TD><b>����Ϣ��ϵ�ˣ�</b>"&rs("name")&"<br><b>��ϵ�ֻ����룺</b>"&rs("CompPhone")&"<p>"
response.Write "Ŀǰ���ʻ��ࣺ"&conn.Execute("Select jf From DNJY_user Where username='"&username&"'")(0)&"�����</TD></TR></TABLE>"
response.Write "<p align=""center"">����ˢ�±�ҳ�������ظ��۷֣�</p>"
Call CloseRs(rs)
Call CloseDB(conn)
response.end
Else
response.Write "������������Ļ��ֲ��㣬�޷���ѯ��"
end If
else
response.Write "���������޷���ѯ������ԭ��֮һ��1�����Ļ��ֲ��㡣2����Աδ��¼��"
end If
Call CloseRs(rs)
Call CloseDB(conn)%>
