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
<title>�ռ������</title>
<link rel="stylesheet" href="../css/manage.css" type="text/css">
<BODY>
</head>
<script language=javascript src=../Include/mouse_on_title.js></script>
<table width="98%" border="1"  style="border-collapse: collapse; border-style: dotted; border-width: 0px"  bordercolor="#333333" cellspacing="0" cellpadding="2">
<tr class=backs><td class=td height=18>�ռ���</td></tr>
<%Dim admin,pages,allPages,page,neirong,riqi,isnew,fname,id
admin=request.cookies("administrator")
Call OpenConn
set rs=Server.CreateObject("ADODB.recordset")
sql="select * from DNJY_Message where name='"&admin&"' order by riqi "&DN_OrderType&""
rs.open sql,conn,1,3
if rs.eof and rs.bof then 
response.write "<tr><td>�ռ�����û����Ϣ��</td></tr>"
else
response.write "<marquee><font color=red>��ʱɾ�����ö��ţ�ϵͳ��ʱɾ�����ţ���Ҫ�����뼰ʱת�Ʊ��棬ȷ����ȫ��</font></marquee>"
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
rs.AbsolutePage = page

do while not rs.eof and pages>0
neirong=rs("neirong")
riqi=rs("riqi")
isnew=rs("isnew")
fname=rs("fname")
id=rs("id")
if pages<10 then response.write "<tr><td height=15></td></tr>"
%>
<tr><td onMouseOver="bgColor='#FFFFFF'" onMouseOut="bgColor='#e7e7e7'">
<span alt=�ظ�����Ϣ style="cursor:hand" onclick="{location.href='Message2.asp?back=kkk1&id=<%=id%>';}"><img src=../images/m_replyp.gif border=0 align=center></span>&nbsp;&nbsp;
<span alt=ת����������Ա style="cursor:hand" onclick="{location.href='Message2.asp?back=kkk1&fw=yes&id=<%=id%>';}"><img src=../images/m_fw.gif border=0 align=center></span>&nbsp;&nbsp;
<span alt=����ɾ������Ϣ style="cursor:hand" onclick="{if(confirm('�ò������ɻָ���\n\nȷʵɾ��������Ϣ�� ')){location.href='Message4.asp?delid=<%=id%>';}}"><img src=../images/m_delete.gif border=0 align=center></span>&nbsp;&nbsp;
�����ˣ�<font color=red><%=fname%></font> &nbsp; &nbsp; &nbsp; ����ʱ�䣺<%=riqi%>&nbsp;&nbsp;<%if rs("hf")=1 then%><font color=red>�ѻظ�</font><%else%><b>δ�ظ�</b><%end if%>
<hr color=#e7e7e7><BR>
<%=replace(neirong,vbCRLF,"<BR>")%><br>&nbsp;
</td></tr>
<%
pages = pages - 1
rs.movenext
if rs.eof then exit do
loop
response.write "</table>"
call listpages()
end if
Call CloseRs(rs)
Call CloseDB(conn)

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
%>


</td></tr></table>
</body>
</html>

