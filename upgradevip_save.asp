<!--#include file="conn.asp"-->
<!--#include file="SqlIn.Asp"-->
<!--#include file=CONfig.ASP-->
<!--#include file="Include/Function.asp"-->
<!--#include file="Include/label.asp"-->
<!--#include file=usercookies.asp-->
<!--#include file="Include/tt_auto_lock.asp" -->
<%
'=====================================
'���칩����Ϣ��վ����ϵͳ
'��������Ƽ������ҳ�Ʒ
'�ͷ�����http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>
<%
dim rs2,username,aliname,ywje,ywlx,czbz,vipys,Amount,dqts,vip
username=trim(request.Form("username"))
if username="" then
response.write "<li>��������"
response.end
end If
Call OpenConn
set rs2=server.createobject("adodb.recordset")
rs2.open"select username from DNJY_user where  username='"&username&"'",conn,1,1
if rs2.eof then
Response.Write ("<script language=javascript>alert('���û�������!');history.go(-1);</script>")
Call CloseRs(rs2)
Call CloseDB(conn)
response.end
end If
Call CloseRs(rs2)

username=trim(request.Form("username"))
aliname=trim(request("aliname"))
ywje=CInt(request.Form("ywje"))
vipys=CInt(request.Form("vipys"))
vipje=CInt(request.Form("vipje"))
ywlx=trim(request.Form("ywlx"))
czbz=HTMLEncode(trim(request.Form("czbz")))
Amount=CInt(request.Form("Amount"))
vip=CInt(request.Form("vip"))

if vip=1 then
Response.Write ("<script language=javascript>alert('������VIP��Ա����������!');history.go(-1);</script>")
response.end
end If

if vipys<vipsj then
Response.Write ("<script language=javascript>alert('��վҪ��VIP�ʸ�����һ���Թ���"&vipsj&"����!');history.go(-1);</script>")
response.end
end If

if Amount<0 Or Amount<vipje*vipys then
Response.Write ("<script language=javascript>alert('���ʻ������㣬���ֵ��������VIP��Ա�ʸ�!');history.go(-1);</script>")
response.end
end If

set rs=server.createobject("adodb.recordset")
sql="select * from [DNJY_Financial] "
rs.open sql,conn,1,3
rs.addnew
rs("username")=username
rs("aliname")=aliname
rs("ywje")=vipje*vipys
rs("ywlx")=ywlx
rs("czbz")=czbz
rs("czz")=username
rs("admincl")=1
rs("addTime")=now()
rs.update
Call CloseRs(rs)

conn.execute "UPDATE DNJY_user SET Amount=Amount-"&vipje*vipys&" WHERE username='"&username&"'" 'ͬʱ���û������ʻ����
conn.execute "UPDATE DNJY_user SET tAmount=tAmount+"&vipje*vipys&" WHERE username='"&username&"'" 'ͬʱ���û����������ܽ��
dqts=vipys*30
dqts= dateadd("d", dqts, now)
conn.execute "UPDATE DNJY_user SET vip=1 WHERE username='"&username&"'" 'ͬʱ����ΪVIP�ʸ�
conn.execute "UPDATE DNJY_user SET vipdq='"&dqts&"' WHERE username='"&username&"'" 'ͬʱ����VIP�ʸ���ʱ��
response.Write "<script language=javascript>alert('�����ɹ�!');location.replace('upgradevip.asp');</script>"
Call CloseDB(conn)%>

