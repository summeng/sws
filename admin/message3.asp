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
<title>վ�ڶ��Ź���</title>
<link rel="stylesheet" href="../css/manage.css" type="text/css">
<BODY>
</head>
<script language=javascript src=../Include/mouse_on_title.js></script>
<script language="JavaScript">
<!--
function CheckAll(form)
{
for (var i=0;i<form.elements.length;i++)
{
var e = form.elements[i];
if (e.name != 'chkall')
e.checked = form.chkall.checked;
}
}

//-->
</script>
	<table width="98%" border="1"  style="border-collapse: collapse; border-style: dotted; border-width: 0px"  bordercolor="#333333" cellspacing="0" cellpadding="2">
	<tr class=backs><td class=td height=18>վ�ڶ���</td></tr>
<%dim pages,allPages,page,neirong,riqi,isnew,del,fname,name,id,temp
Call OpenConn
set rs=Server.CreateObject("ADODB.recordset")
sql="select * from DNJY_Message order by riqi "&DN_OrderType&""
rs.open sql,conn,1,3
if rs.eof and rs.bof then 
response.write "<tr><td>û��վ�ڶ�����Ϣ��</td></tr>"
else
response.write "<marquee><font color=red>��ʱɾ�����ö��ţ���Ҫ�����뼰ʱת�Ʊ��棬ȷ����ȫ��</font></marquee>"
pages = 10			'ÿҳ��¼��
rs.pageSize = pages
allPages = rs.pageCount		'����һ���ֶܷ���ҳ
page = Request("page")
'if������ڻ������Ŵ���
if isEmpty(page) or clng(page) < 1 then
page = 1
elseif clng(page) >= allPages then
page = allPages 
end if
rs.AbsolutePage = page%>
<form name="form1" method="post" action="Message4.asp">
<%do while not rs.eof and pages>0
neirong=rs("neirong")	'����
riqi=rs("riqi")		'����
isnew=rs("isnew")	'�Ƿ��Ķ�
del=rs("del")		'�Ƿ�ɾ��
fname=rs("fname")   '������
name=rs("name")		'�ռ���
id=rs("id")		'���
if isnew="0" then
	temp="�ռ���δ�Ķ�"
elseif isnew="1" then
	if del="0" then 
	temp="<font color=red>�ռ������Ķ�</font>"
	elseif del="1" then
	temp="<font color=red>�ռ�����ɾ��</font>"
	end if	
end if
if name="ALL" then temp="��������"
if pages<10 then response.write "<tr><td height=15></td></tr>"
%>


<tr><td onMouseOver="bgColor='#FFFFFF'" onMouseOut="bgColor='#e7e7e7'">
ѡ��<input name="delid3" type="checkbox" id="delid3" value="<%=rs("id")%>">
&nbsp;&nbsp;
<span alt=����ɾ������Ϣ style="cursor:hand" onclick="{if(confirm('�ò������ɻָ���\n\nȷʵɾ��������Ϣ�� ')){location.href='Message4.asp?delid3=<%=id%>';}}"><img src=../images/m_delete.gif border=0 align=center></span><span alt=ת����������Ա style="cursor:hand" onclick="{location.href='Message2.asp?back=kkk3&fw=yes&id=<%=id%>';}"><img src=../images/m_fw.gif border=0 align=center></span>&nbsp;&nbsp;
�����ˣ�<a href="Message2.asp?back=kkk1&id=<%=id%>"><font color=red><%=fname%></font></a>&nbsp;&nbsp;
�ռ��ˣ�<a href="Message2.asp?back=kkk1&name=<%=name%>"><font color=red><%=name%></font></a> &nbsp; &nbsp; &nbsp; ����ʱ�䣺<%=riqi%> &nbsp; &nbsp; &nbsp; <b><%=temp%></b>
<hr color=#e7e7e7><BR>
<%=replace(neirong,vbCRLF,"<BR>")%><br>&nbsp;
</td></tr>


<%
pages = pages - 1
rs.movenext
if rs.eof then exit do
loop
response.write "</table>"%>
<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%">
<tr>
<td height="20" colspan="15" align="center" bgcolor="#F3F3F3">	
  <p align="center">
    <input type="checkbox" name="chkall" onClick="CheckAll(this.form)">
    ȫ��ѡ�� 
    <input type="submit" name="Submit" value="ɾ����ѡ��Ϣ" onClick="{if(confirm('�ò������ɻָ���\n\nȷʵɾ��ѡ������Ϣ��')){return true;}return false;}">
  </p>
 </td> 
</tr>
</form>
</table>
<%
call listpages()
end if
Call CloseRs(rs)%>

<%
'��ҳ
sub listPages() 
'if allpages <= 1 then exit sub 
if page = 1 then
response.write "<font color=darkgray>��ҳ ǰҳ</font>"
else
response.write "<a href="&request.ServerVariables("script_name")&"?page=1>��ҳ</a> <a href="&request.ServerVariables("script_name")&"?page="&page-1&">ǰҳ</a>"
end if
if page = allpages then
response.write "<font color=darkgray> ��ҳ ĩҳ</font>"
else
response.write " <a href="&request.ServerVariables("script_name")&"?page="&page+1&">��ҳ</a> <a href="&request.ServerVariables("script_name")&"?page="&allpages&">ĩҳ</a>"
end if
response.write " ��"&page&"ҳ ��"&allpages&"ҳ"
end sub
Call CloseDB(conn)%>


</td></tr></table>
</body>
</html>

