<!--#include file="conn.asp"-->
<!--#include file="SqlIn.Asp"-->
<!--#include file=top.asp-->
<!--#include file=usercookies.asp-->
<%
'=====================================
'���칩����Ϣ��վ����ϵͳ
'��������Ƽ������ҳ�Ʒ
'�ͷ�����http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>
<%dim mess%>
<link rel="stylesheet" type="text/css" href="/<%=strInstallDir%>css/css.css">
<link rel="stylesheet" type="text/css" href="../1.CSS">
<style type="text/css">
<!--
.style1 {color: #FF0000}
.style5 {color: #FF000; font-weight: bold; }
-->
</style>
</head>

<body background="img/bg1.gif" leftmargin="0" topmargin="0">

<div align="center">
  <center>
  <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="980" height="426">
    <tr>
      <td width="214" height="354" valign="top" bgcolor="#FFFFFF"><div align="center">
      
 <!--#include file=userleft.asp--></div></td>
 <td width="5">&nbsp;</td>
 <td width="760" height="354" valign="top" bgcolor="#FFFFFF">
 <%Dim rsmess,pages,allPages,page,neirong,riqi,isnew,fname,delid

if request.cookies("dick")("username")="" then
response.Write "<center>���ȵ�¼</center>"
response.End
end if
username=request.cookies("dick")("username")

If request("action")="msgdel" Then
   Set rs = conn.Execute("select * from DNJY_Message where id="&request("delid")) 
	if rs.eof and rs.bof then 
	response.write "<script language='javascript'>"
	response.write "alert('�����ˣ����ݿ����ʧ�ܣ�');"
	response.write "location.href='javascript:history.go(-1)';"
	response.write "</script>"
	response.end
    elseif ucase(rs("name"))="ALL" then
	response.write "<script language='javascript'>"
	response.write "alert('�����ˣ�������ɾ��������Ϣ������������Ϣ���ڹ�����Ϣ��');"
	response.write "location.href='javascript:history.go(-1)';"
	response.write "</script>"
	response.end
    elseif Ucase(request.cookies("dick")("username"))<>Ucase(rs("name")) then
	response.write "<script language='javascript'>"
	response.write "alert('�����ˣ�������ɾ��������Ϣ������������Ϣ���ڹ�����Ϣ��');"
	response.write "location.href='javascript:history.go(-1)';"
	response.write "</script>"
	response.end
    else
    conn.Execute ("update DNJY_Message set del='1' where id="&request("delid")) 
	response.write "<script language='javascript'>"
	response.write "alert('�����ɹ�����ѡ��Ϣ�ѱ�ɾ����');"
	response.write "location.href='?"&C_ID&"';"
	response.write "</script>"
	response.end
    end if
conn.close
set conn=nothing
End If

set rs=server.createobject("adodb.recordset")
sql = "select mess from [DNJY_user] where username='"&username&"'"
rs.open sql,conn,1,1
mess=rs("mess")
Call CloseRs(rs)

If mess=0 then
response.write "<table width=100% border=0 align=center><tr><td><b>��ӭ��Ϣ��</b><p>&nbsp;&nbsp;&nbsp;&nbsp;�𾴵� "&username&" ����!ף������Ϊ["&webname&"]��Ա�����������ɷ����Ͳ�ѯ��Ϣ�����ס��վ��ַ"&web&"�������û��������룬�Ժ󷢲���Ϣ����ݡ�����ѡ��վ�ṩ��VIP��Ա��������Ϣ���г����񣬽������Լ�����������̼ҵ��̡���ҵ��Ƭ�ͷ���������Ϣ�ȣ�ʹ������Ϣ�õ����õ��Ƽ���������λ����ʾ������������ġ�����֧�������Ƿ�չ�Ķ�������ӭ�����������Ѿ����ؼҿ�����</td></tr></table>"
conn.execute "UPDATE DNJY_user SET mess=1 WHERE username='"&request.cookies("dick")("username")&"'" '��ǻ�ӭ�������Ķ�
End if
set rsmess=server.CreateObject("adodb.recordset")
rsmess.Open "select * from DNJY_Message where del='0' and (name='ALL' or name='"&username&"') order by name asc, riqi "&DN_OrderType&"",conn,1,1

if rsmess.eof and rsmess.bof  then 
If mess<>0 then
	response.write "<table width=100% border=0 align=center class=""ty1""><tr><td><marquee width=100% height=10>��ʱû��վ����Ϣ</marquee></td></tr></table>"
	Else
response.write ""
End if
	else
response.write "<marquee><font color=red>��Լÿһ�ֿռ䣬�뼰ʱɾ�����ö��ţ�лл��������Ҫ�����뼰ʱת�Ʊ��棬ȷ����ȫ��</font></marquee>"
pages = 10			'ÿҳ��¼��
rsmess.pageSize = pages
allPages = rsmess.pageCount		'����һ���ֶܷ���ҳ
page = Request("page")
'if������ڻ������Ŵ���
if isEmpty(page) or clng(page) < 1 then
page = 1
elseif clng(page) >= allPages then
page = allPages 
end if
rsmess.AbsolutePage = page

do while not rsmess.eof and  pages>0
neirong=rsmess("neirong")
riqi=rsmess("riqi")
isnew=rsmess("isnew")
fname=rsmess("fname")
id=rsmess("id")
%>

<form name=msgdel method=post action=saveprofile.asp?action=msgdel&<%=C_ID%>>
<table width="100%" border="1" cellpadding="4" cellspacing="0" bordercolor="#C0C0C0" align=center bordercolordark="#FFFFFF" style="word-break:break-all" class="ty1">
<tr>
<td onMouseOver="bgColor='#e7e7e7'" onMouseOut="bgColor='#FFFFFF'">
<table border=0 width=100%>
<tr><td width=150 rowspan=2>
<a href='messh.asp?id=<%=id%>&fname=<%=fname%>&<%=C_ID%>'><img alt="�ظ�����Ϣ" src=images/m_replyp.gif border=0></a> &nbsp;  
<span  style="cursor:hand" onclick="{if(confirm('��ʾ����Ϣɾ���󲻿ɻָ���\n\n��ȷʵɾ��������Ϣ�� ')){location.href='?action=msgdel&delid=<%=id%>&<%=C_ID%>';}}"><img ALT="ɾ������Ϣ" src=images/m_delete.gif border=0></span>&nbsp;&nbsp;
</td><td>
<%if rsmess("name")="ALL" then
response.write "<font color=blue>����һ����������</font>"
end if
%>
</td></tr>
<tr><td>�����ˣ�<%=fname%>  &nbsp;  ����ʱ�䣺<%=riqi%></td></tr>
<tr><td colspan=2><hr color="#FFFFFF" WIDTH=100%>
<%=HTMLEncode(replace(neirong,vbCRLF,"<BR>"))%>
</td></tr>
</table>
</td></tr>
</table>
</form>
<%
pages = pages - 1
rsmess.movenext
if rsmess.eof then exit do
loop

conn.execute "UPDATE DNJY_Message SET isnew ='1' WHERE name ='"&request.cookies("dick")("username")&"'" '��Ƕ������Ķ�
response.write "<table width=100% border=0 align=center><tr><td height=50 valign=top>"
if page = 1 then
response.write "<font color=darkgray>��ҳ ǰҳ</font>"
else
response.write "<a href="&request.ServerVariables("script_name")&"?action=my_msg&page=1&"&C_ID&">��ҳ</a> <a href="&request.ServerVariables("script_name")&"?action=my_msg&page="&page-1&"&"&C_ID&">ǰҳ</a>"
end if
if page = allpages then
response.write "<font color=darkgray> ��ҳ ĩҳ</font>"
else
response.write " <a href="&request.ServerVariables("script_name")&"?action=my_msg&page="&page+1&"&"&C_ID&">��ҳ</a> <a href="&request.ServerVariables("script_name")&"?action=my_msg&page="&allpages&"&"&C_ID&">ĩҳ</a>"
end if
response.write " ��"&page&"ҳ ��"&allpages&"ҳ"
response.write "</td></tr></table>"

end if
rsmess.close
set rsmess=nothing
%></td>
      <td width="4" valign="top" bgcolor="#e6e6e6"></td>
    </tr>
    <tr>
      <td height="26" colspan="3"><!--#include file=footer.asp--></td>
    </tr>
  </table>
  </center>
</div>
</body>
</html>
