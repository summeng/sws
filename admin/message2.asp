<!--#include file="../conn.asp"-->
<!--#include file="../SqlIn.Asp"-->
<!--#include file="../config.asp"-->
<!--#include file=cookies.asp-->
<%
'=====================================
'���칩����Ϣ��վ����ϵͳ
'��������Ƽ������ҳ�Ʒ
'�ͷ�����http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>
<%call checkmanage("09")%>
<html>
<head>
<title>����վ����Ϣ</title>
<link rel="stylesheet" href="../css/manage.css" type="text/css">
<BODY>
<script language="JavaScript" type="text/JavaScript">
function CheckAll(form)  {
  for (var i=0;i<form.elements.length;i++)    {
    var e = form.elements[i];
    if (e.name != 'chkall')  e.checked = form.chkall.checked; 
   }
  }
</script>
<style type="text/css">
/*��ʾ��Ϣ*/
.info {position:relative;background:#;color:#666; text-decoration:none;font-size:12px;width:0px;height:0px;text-align:center;border:0px solid #ccc;line-height:18px;}/*�������ӵ�����,һ��Ҫ����Ϊrelative����ʹ��ʾ�����������*/
.info:hover {background:#eee;color:#333;}
.info span {display: none }/*���������µ�spanΪ����״̬*/
.info:hover span /*����hover�µ�span����Ϊ����״̬,��������ʾ���λ��*/{display:block;position:absolute;top:0px;left:25px;width:200px;
border:1px solid #DC5826; background:#FFF3B4; color:#333;padding:5px;text-align:left;}
</style>
</head>
<script language=javascript src=../Include/mouse_on_title.js></script>
<%dim admin,action,name,username,id,huifu,vbcrlf,where,neirong,url
admin=request.cookies("administrator")
action=request("send")
Call OpenConn
if action<>"ok" then 
name=request("name")
username=request("username")
id=request("id")
if id<>"" then
Set rs = conn.Execute("select * from DNJY_Message where id="&id) 
name=rs("fname")
huifu=vbcrlf&vbcrlf&vbcrlf&vbcrlf&"======="&rs("fname")&"��"&rs("riqi")&"���͵���Ϣԭ��======="&vbcrlf&replace(rs("neirong"),"<font color=darkgray>","")
end if
%>

	<table width="98%" border="1"  style="border-collapse: collapse; border-style: dotted; border-width: 0px"  bordercolor="#333333" cellspacing="0" cellpadding="2">
	<form action=Message2.asp method=post name="hand">
	<tr class=backs><td colspan=2 class=td height=18 align="center">����վ����Ϣ</td></tr>
	<tr><td width=20% align=right height="18">�� �� ˭&nbsp;</td><td>
	<input type=text 
	<%
	if request("username")<>"" then response.write "value='"&username&"'"
	if request("fw")<>"yes" then response.write "value='"&name&"'"
	%>"
	name="name" size=18 maxlength=16> <a class="info" href="#"><img src="../images/memo.gif"  BORDER="0" /><span>����д��վע���Ա��IDС�ؼ������͸������ˣ�����д��д��ALL</span></a></td></tr>
	<tr><td width=20% align=right height="18">��Ϣ����&nbsp;</td><td>
	<textarea rows="10" name="neirong" style="width=95%; overflow:auto;"><%=huifu%></textarea>
	<tr><td colspan=2>
	<INPUT name="id" TYPE="hidden" value="<%=request("id")%>">
	<INPUT name="send" TYPE="hidden" value="ok">
	<INPUT name="back" TYPE="hidden" value="<%=request("back")%>">
	<INPUT name=action TYPE="submit" value="������Ϣ"></td></tr>
	</form>
	</table>
<%

set rs=nothing
end if

if action="ok" then
	if trim(request.form("name"))="" then
	response.write "<script language='javascript'>"
	response.write "alert('�����ˣ�û����д�ռ��ˣ�');"
	response.write "location.href='javascript:history.go(-1)';"
	response.write "</script>"
	response.end
	end if

	if request.form("neirong")=""  then
	response.write "<script language='javascript'>"
	response.write "alert('�����ˣ�û����дҪ���͵����ݣ�');"
	response.write "location.href='javascript:history.go(-1)';"
	response.write "</script>"
	response.end
	end if

	if len(request.form("neirong"))>500  then
	response.write "<script language='javascript'>"
	response.write "alert('�����ˣ�����д������̫���ˣ�����������ύ��');"
	response.write "location.href='javascript:history.go(-1)';"
	response.write "</script>"
	response.end
	end if

	if  request("name")<>"ALL" then
	set rs=server.createobject("adodb.recordset")
	rs.Open "select username from [DNJY_user] where username='"&request("name")&"'",conn,1,1	
	if rs.eof and rs.bof then
	response.write "<script language='javascript'>"
	response.write "alert('�����ˣ�û���ҵ�"&request("name")&"��Ա������������ύ��');"
	response.write "location.href='javascript:history.go(-1)';"
	response.write "</script>"
	response.end
	end if
	end if
'set rs=server.CreateObject("adodb.recordset")
'rs.Open "select * from DNJY_Message  where id="&request("id")&"",conn,1,3
'rs("hf")=1
'rs.Update
'rs.Close
'set rs=nothing

	sql="select * from DNJY_Message"
	set rs=server.createobject("adodb.recordset")
	rs.open sql,conn,1,3
	rs.addnew
	  	rs("name")=trim(request.form("name"))
		neirong=request.form("neirong")
		neirong=replace(neirong,"=======","<font color=darkgray>======")
	  	rs("neirong")=neirong
	  	rs("riqi")=now()
	  	rs("fname")=""&admin&""

	rs.update	
	rs.close
	conn.close
	set rs=nothing
	set conn=nothing


	where=request("where")
	if where="" then
	url="Message2.asp"
	else
	url=where
	end if

	response.write "<script language='javascript'>"
	response.write "alert('վ����Ϣ���ͳɹ���');"
	if request("back")="user" then
	response.write "location.href='user1.asp?action=useredit&id="&request("name")&"';"
	elseif request("back")="kkk1" then
	response.write "location.href='Message1.asp';"
	elseif request("back")="kkk3" then
	response.write "location.href='Message3.asp';"
	elseif request("back")="hand5" then
	response.write "location.href='Message5.asp';"
	else
	response.write "location.href='Message2.asp';"
	end if
	response.write "</script>"
	response.end
end if%>

</body>
</html>
