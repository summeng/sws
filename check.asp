<!--#include file="conn.asp"-->
<!--#include file="config.asp"-->
<!--#include file="Include/head.asp"-->
<!--#include file="Include/Function.asp"-->
<!--#include file="SqlIn.Asp"-->
<!--#include file=Include/md5.asp-->
<%
'=====================================
'���칩����Ϣ��վ����ϵͳ
'��������Ƽ������ҳ�Ʒ
'�ͷ�����http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>

<html>
<head>
<title><%=webname%> | ע���Ա�ʼ�ȷ��</title>
<meta http-equiv="Content-Type" content="text/html; charSet=gb2312">
<link rel="stylesheet" type="text/css" href="/<%=strInstallDir%>css/css.css">
</head>
<body leftmargin="0" Topmargin="0" marginwidth="0" marginheight="0">

<%Call OpenConn
Dim ranid,email,mailname,username,recommend

'ȡ�����ȷ����
ranid=trim(Request("ranid"))
mailname=trim(Request("mailname"))
recommend=trim(Request("recommend"))

'�ж�������Ƿ���ȷ,�����ȷ�򷵻�TRUE
Function Isranid(ranid)
	Isranid=FALSE
	Set rs=Server.Createobject("adodb.recordset") 
	sql="select * from DNJY_usertemp where ranid='"&ranid&"' and email='"&mailname&"'"
	rs.open sql,conn,1,1
		If NOT(rs.bof AND rs.eof) Then
			Isranid=TRUE
			email=rs("email")
			username=rs("username")
		End If
	'rs.close
End Function

'����������ȷ�����
If Isranid(ranid) Then
	'��ȷ�����EMAIL����ʽ���ݱ�
	Call addmail()
Else
response.write "1���������ԣ��ʼ�ȷ��ʧ�ܣ�2��Ҳ��������QQ��������䰲ȫ��ʾ����ȷ�ϵ�ԭ��ʵ���ϵ�һ���ѳɹ�ȷ�ϣ������Ա��¼�����������Ա��ϵ��"
End If

Sub addmail()

'��ӻ�Ա����ʽ���ݱ�

Dim rst
set rst=server.createobject("adodb.recordset")
sql="select * from [DNJY_usertemp] where username='"&username&"'"
rst.open sql,conn,1,1
if rst.eof or rst.bof then'�ж��û����Ƿ����
'����û���������
else
set rs=server.createobject("adodb.recordset")
sql="select * from [DNJY_user]"
rs.open sql,conn,1,3
rs.addnew
rs("username")=rst("username")
rs("password")=rst("password")
rs("problem")=rst("problem")
rs("answer")=rst("answer")
rs("name")=rst("name")
rs("email")=rst("email")
rs("maillist")=rst("maillist")
rs("idcard")=rst("idcard")
rs("userip")=rst("userip")
rs("Sex")=rst("Sex")
rs("dianhua")=rst("dianhua")
rs("CompPhone")=rst("CompPhone")
rs("youbian")=rst("youbian")
rs("dizhi")=rst("dizhi")
rs("mpname")=rst("mpname")
rs("fax")=rst("fax")
rs("http")=rst("http")
rs("qq")=rst("qq")
rs("weixin")=rst("weixin")
rs("zcdata")=now()
rs("jftime")=now()
rs("vipdq")=now()
rs("a")=rst("a")
rs("b")=rst("b")
rs("c")=rst("c")
rs("d")=rst("d")
rs("jf")=rst("jf")
rs("hb")=rst("hb")
rs("city_oneid")=rst("city_oneid")
rs("city_twoid")=rst("city_twoid")
rs("city_threeid")=rst("city_threeid")
rs("city_one")=rst("city_one")
rs("city_two")=rst("city_two")
rs("city_three")=rst("city_three")
rs("ranid")=rst("ranid")
If zdsh=1 Then rs("useryz")=1
rs("mailyz")=1
If recommend<>"" Then rs("tjname")=recommend'�Ƽ���
rs("qqopenid")=rst("qqopenid")
rs("QQtoken")=rst("QQtoken")
rs("UserFace")=rst("UserFace")
rs("SinaId")=rst("SinaId")
rs("weiboNick")=rst("weiboNick")
rs("alipayID")=rst("alipayID")
rs.update 
End if


'ϵͳ��ע���û����ͻ�ӭվ�ڶ���
Dim rs3,sql3,admin
set rs3=server.createobject("adodb.recordset")
sql3="select top 1 username from [DNJY_admin]"
rs3.open sql3,conn,1,1
admin=rs3("username")
Call CloseRs(rs3)

If recommend<>"" Then'ͬʱ���Ƽ������ӻ��ֺͷ��Ͷ���֪ͨ
Dim rsjf
    conn.execute "UPDATE DNJY_user SET jf=jf+"&jf_5&" WHERE username='"&recommend&"'" 
	set rsjf=server.createobject("adodb.recordset")
	sql="select * from DNJY_Message"
	rsjf.open sql,conn,1,3
	rsjf.addnew
	  	rsjf("name")=recommend		
	  	rsjf("neirong")="��ϲ�����Ƽ�������"&username&"ע�᱾վ��Ա����վ���������˻���"&jf_5&"�֡�"
	  	rsjf("riqi")=now()
	  	rsjf("fname")=admin
	rsjf.update	
	rsjf.close	
	set rsjf=nothing	
End If

Response.cookies("reg")("regkey")=""


'ɾ����ʱ���ݱ��е��ʼ�
Set rs=Server.Createobject("adodb.recordset")     
sql="delete from DNJY_usertemp where username='"&username&"'"     
rs.open sql,conn,1,3 
rst.close
set rst=Nothing
'��ʾ��ӭ��Ϣ
%>
<center><FORM method="post" action="http://<%=web%>">
    <table border="1" cellspacing="0" cellpadding="0" width="550" bordercolor="#7C96B8" bordercolordark="#FFFFFF" bordercolorlight="#7C96B8" height="132" >
      <tr> 
<td height="20" colspan="2" bgcolor="#7C96B8" background="img/mmTo.gIf" width="546" align="center"><DIV style="FILTER: dropshadow(color=#FFFFFF, offx=1, offy=1, positive=1); WIDTH: 100%; CURSOR: hAND; POSITION: relative"><font color="#336699">���Ѿ��ɹ�<font color=#FF0000>����</font>ȷ�������Ļ�Ա��Ϣ��</font></div></td>
      </tr>
      <tr> 
<td height="84" colspan="2" width="546"><font color="#336699">&nbsp;</font>��л����ʽע�᡺<%=webname%>����Ա��������Ϊ<%=webname%>���ͥ������Ա��

��<%=webname%>������Ϊ������</td>
      </tr>
      <tr> 
<td colspan="2" align="center" bgcolor="#7C96B8" height="20" background="img/mmTo.gIf" width="546"> <input name="bank" type="Submit" value="������ҳ" class="Tips_bo"> 
</td>
      </tr>
    </table>
  </FORM>
</center>
<%End Sub%>

<%'�ر����ٶ���
Set rs=nothing
conn.close
Set conn=nothing %>

