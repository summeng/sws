<!--#include file="../conn.asp"-->
<!--#include file="../config.asp"-->
<!--#include file="../Include/Function.asp"-->
<%
'=====================================
'���칩����Ϣ��վ����ϵͳ
'��������Ƽ������ҳ�Ʒ
'�ͷ�����http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>
<%Dim id,city_one,city_two,city_three,issue,issueid,editionname,Releasetime,lr
id=trim(request("id"))
city_one=request("city_one")
city_two=request("city_two")
city_three=request("city_three")
issue=request("issue")
lr=trim(request("lr"))
Call OpenConn
If id="" Then
response.write "���޷��ϰ���"
response.end
End if
set rs=server.CreateObject("ADODB.Recordset")
sql="select * from DNJY_zz_edition where state="&DN_True&" and id="&id&""
rs.open sql,conn,1,1
if rs.bof or rs.eof then
response.write"��������û�з����������棡"
response.end
End if
editionname=rs("editionname")
Releasetime=rs("Releasetime")
issueid=rs("issueid")
Response.Cookies("issueid")=id
Call CloseRs(rs)
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE><%=webname%>��������־��<%=city_one&city_two&city_three%>�� ��<%=issue%>�� <%=editionname%>��</TITLE>
<META http-equiv="Content-Type" content="text/html; charset=gb2312">
<META http-equiv=Content-Language content=zh-cn>
<META content="<%=webname%>������־" name="keywords">
<META content="<%=webname%>������־,���ӹ��,����DM���" name="description">
<STYLE>
.stylea {
	FILTER: progid:DXImageTransform.Microsoft.Matrix(sizingMethod='auto expand',m11=0,m12=1,m21=-1,m22=0)
}
.styleb {
	FILTER: progid:DXImageTransform.Microsoft.Matrix(sizingMethod='auto expand',m11=0,m12=-1,m21=1,m22=0)
}
</STYLE>

<META content="MSHTML 6.00.2900.3020" name=GENERATOR></HEAD>
<BODY oncontextmenu="return false" onselectstart="return false" bottomMargin=0 leftMargin=0 topMargin=0 rightMargin=0  style="text-align:center;">
<DIV style="width: 100px; height: 0px; z-index: 1" id="layer1">
<TABLE height=30 cellSpacing=0 cellPadding=0 width=100% border=0>
	<tr> 
          <TD><A title="������������ת90�Ȳ鿴" href="show_largerpic.asp?lr=l&id=<%=id%>&issue=<%=issue%>&i=0&city_one=<%=city_one%>&city_two=<%=city_two%>&city_three=<%=city_three%>"><IMG 
            height=30 src="img/lr_l90.gif" width=120 border=0></A></TD>
          <TD><A title="�����鿴��־" 
            href="show_largerpic.asp?id=<%=id%>&issue=<%=issue%>&i=0&city_one=<%=city_one%>&city_two=<%=city_two%>&city_three=<%=city_three%>"><IMG 
            height=30 src="img/normal.gif" width=81 border=0></A></TD>
          <TD><A title="������������ת90�Ȳ鿴"
            href="show_largerpic.asp?lr=r&id=<%=id%>&issue=<%=issue%>&i=0&city_one=<%=city_one%>&city_two=<%=city_two%>&city_three=<%=city_three%>"><IMG 
            height=30 src="img/lr_r90.gif" width=120 border=0></A></TD>         
	</tr>
</TABLE></DIV>
<DIV class=drag id=mov style="margin:5px 0px;text-align:center;padding:0px;"><IMG <%if lr="l" then%>class='stylea'<%elseif lr="r" then%>class='styleb'<%end if%> alt=<%=editionname%> src="show_thumbnail.asp?dt=ok" galleryimg="no"></div>
<%Conn.Execute("Update DNJY_zz_edition Set Click=Click+1 where id="&id&"")
Conn.Execute("Update DNJY_zz_issue Set Click=Click+1 where id="&issueid&"")
Call CloseDB(conn)%>










































</BODY></HTML>
