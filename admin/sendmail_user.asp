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
<%call checkmanage("08")%>
<script language=javascript src=../Include/mouse_on_title.js></script>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" type="text/css" href="../css/css.css">
<BODY>
<script language=javascript>
<!--
function CheckAll(form){
for (var i=0;i<form.elements.length;i++){
var e = form.elements[i];
if (e.name != 'chkall') e.checked = form.chkall.checked; 
}
}
-->
</script>
</head>
<BODY>
<BODY>

<%Call OpenConn
dim action,pages,totaluser,allPages,page,email,keywords,CurrentPage,n
action=request("action")
if action="sendmail" then
call sendmail()
else
%>
<style type="text/css">
<!--
body {
	background-color: #E3EBF9;
}
.style1 {
	color: #FFFFFF;
	font-weight: bold;
	font-size: 12px;
}
.style2 {font-size: 12px}
-->
</style>
<form action=sendmail_user.asp method=post name=email>
<table width="98%" border="1"  style="border-collapse: collapse; border-style: dotted; border-width: 0px"  bordercolor="#DADAE9" cellspacing="0" cellpadding="2">
<tr>
<td colspan=4 bgcolor="#BDBEDE" height=28 align=center><B>���ĵ�����־��Ա�ʼ��б�</b></td></tr>

<%  
pages = 20
sql="select * from DNJY_user where email<>'' and maillist=1   order by id "&DN_OrderType&""
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,3
if rs.eof and rs.bof then 
response.write "<tr><td colspan=4 align=center height=50>û�ж�����־��Ա��δ�����ʼ��б�</td></tr>"
response.end 
Call CloseRs(rs)
end if

totaluser=RS.RecordCount
rs.pageSize = pages
allPages = rs.pageCount
page = Request("page")
If not isNumeric(page) then page=1
if isEmpty(page) or clng(page) < 1 then
page = 1
elseif clng(page) >= allPages then
page = allPages 
end if
				If totaluser Mod pages=0 Then  
					n= totaluser \ pages  
				Else
					n= totaluser \ pages+1  
				End If
rs.AbsolutePage = page

response.write "<tr><td align=center width='5%' bgcolor='#DADAE9'>ѡ��</td><td width=30% align=center bgcolor='#DADAE9'>��ԱID</td><td width=30% align=center bgcolor='#DADAE9'>��Ա����</td><td align=center bgcolor='#DADAE9'>��ԱEmail</td></tr>"

Do While Not rs.eof and pages>0 
response.write "<tr><td height=25 bgcolor='#FAFAFA'><input type='checkbox' value='"&rs("eMail")&"' name=email></td><td bgcolor='#FAFAFA'>"&rs("UserName")&"</td><td bgcolor='#FAFAFA'>"&rs("Name")&"</td><td bgcolor='#FAFAFA'>"&rs("email")&"</td></tr>"

pages = pages - 1
rs.movenext
if rs.eof then exit do
loop
%>

<tr><td colspan=4 bgcolor='#DADAE9'>
<input type='checkbox' name=chkall onclick='CheckAll(this.form)'>ȫѡ 
<input type=hidden name=action value=sendmail>
<input type=submit value="Ⱥ���ʼ�">
</td></tr>
<table>
</form>
<%
call listPages()
end if

sub listPages() 
'��ҳ
'if allpages<=1 then exit sub
response.write "<br>�ܼ�¼��<font color=red>"&totaluser&"</font> &nbsp; ÿҳ<font color=red>20</font>&nbsp; "
Response.Write "<font class='contents'> ҳ�Σ�</font><font class='contents'>"&Page&"</font><font class='contents'>/"&n&"ҳ</font> " 
if page = 1 then
response.write "<font color=darkgray>��ҳ ǰҳ</font>"
else
response.write "<a href=sendmail_user.asp?keywords="&keywords&"&page=1>��ҳ</a> <a href=sendmail_user.asp?keywords="&keywords&"&page="&page-1&">ǰҳ</a>"
end if
if page = allpages then
response.write "<font color=darkgray> ��ҳ ĩҳ</font>"
else
response.write " <a href=sendmail_user.asp?keywords="&keywords&"&page="&page+1&">��ҳ</a> <a href=sendmail_user.asp?keywords="&keywords&"&page="&allpages&">ĩҳ</a>"
end if
end sub


sub sendmail()

email=request("email")
if email="" or isnull(email) then
response.write "<script language='javascript'>"
response.write "alert('�����ˣ���ʲôҲû��ѡ��');"
response.write "location.href='javascript:history.go(-1)';"
response.write "</script>"
response.end
end if
email=replace(email,", ",";")
response.redirect "sendmail.asp?"&email
end sub
%>

</body>
</html>
<%
conn.close
set conn=nothing
%>