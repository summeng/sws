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
<%call checkmanage("03")
If request("ywje")="" then
Response.Write ("<script language=javascript>alert('����Ϊ��!');history.go(-1);</script>")
 response.end
 end If
dim rs2,username,aliname,ywje,ywlx,czbz,admincl,ddhm,zffs,ddshsj
username=trim(request("username"))
aliname=trim(request("aliname"))
ywje=clng(Formatnumber(request("ywje"),2,-1,-1,0))
username=trim(request("username"))
aliname=trim(request("aliname"))
ywlx=trim(request("ywlx"))
czbz=trim(request("czbz"))
admincl=trim(request("admincl"))
ddhm=trim(request("ddhm"))
zffs=trim(request("zffs"))
czbz=trim(request("czbz"))
ddhm=trim(request("ddhm"))
zffs=trim(request("zffs"))
if username="" then
response.write "<li>��������"
response.end
end If
Call OpenConn
set rs2=server.createobject("adodb.recordset")
rs2.open"select username,Amount from DNJY_user where  username='"&username&"'",conn,1,1
if rs2.eof then
Response.Write ("<script language=javascript>alert('���û�������!');history.go(-1);</script>")
 response.end
 end If
If (ywlx="֧��" Or ywlx="����") And ywje>clng(Formatnumber(rs2("Amount"),2,-1,-1,0)) Then
Response.Write ("<script language=javascript>alert('�˻�Ա�ʻ����㣬�޷�����!');history.go(-1);</script>")
response.end
End if
Call CloseRs(rs2)

set rs=server.createobject("adodb.recordset")
sql="select * from [DNJY_Financial] "
rs.open sql,conn,1,3
rs.addnew
rs("username")=username
rs("aliname")=aliname
rs("ywje")=ywje
rs("ywlx")=ywlx
rs("czbz")=czbz
rs("czz")=request.cookies("administrator")
If admincl="" Then
rs("admincl")=1
else
rs("admincl")=admincl
End if
rs("addTime")=now()
rs.update
Call CloseRs(rs)
If ywlx="����" then
conn.execute "UPDATE DNJY_user SET Amount=Amount+"&ywje&" WHERE username='"&username&"'" 'ͬʱ���û������ʻ����
Else
conn.execute "UPDATE DNJY_user SET Amount=Amount-"&ywje&" WHERE username='"&username&"'" 'ͬʱ���û������ʻ����
End If
If ddhm<>"" then
ddshsj=DN_Now
conn.execute "UPDATE DNJY_order SET zfqr=1 WHERE ddhm='"&ddhm&"'" 'ͬʱ�����û���������״̬
conn.execute "UPDATE DNJY_order SET zffs="&zffs&" WHERE ddhm='"&ddhm&"'" 'ͬʱ�����û��������ʽ
conn.execute "UPDATE DNJY_order SET admincl=1 WHERE ddhm='"&ddhm&"'" 'ͬʱ�����û��������״̬
conn.execute "UPDATE DNJY_order SET ddshsj="&ddshsj&" WHERE ddhm='"&ddhm&"'" 'ͬʱ�����û���������ʱ��
End If
Call CloseDB(conn)
response.Write "<script language=javascript>alert('�����ɹ���ˢ��ҳ��ɼ��½��!');window.close();</script>"
%>

