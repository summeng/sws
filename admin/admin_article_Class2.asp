<!--#include file="../conn.asp"-->
<!--#include file="../config.asp"-->
<!--#include file="../SqlIn.Asp"-->
<!--#include file=cookies.asp-->
<!--#include file="../Include/Function.asp"-->
<!--#include file="../Include/label.asp"-->
<%
'=====================================
'���칩����Ϣ��վ����ϵͳ
'��������Ƽ������ҳ�Ʒ
'�ͷ�����http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
call checkmanage("06")
Call OpenConn
dim two,twoid,action,oneid,idd,i,dname,dbpath,mname,yname,adid,fdel,iPage,pic,sql1,sql2,sql3
twoid=strint(request.QueryString("twoid"))
oneid=strint(request.QueryString("oneid"))
idd=strint(request.QueryString("idd"))
action=HtmlEncode(request.querystring("action"))
dname=HtmlEncode(request("dname"))
dbpath="../"

set rs=server.CreateObject("adodb.recordset")
rs.open "select * from DNJY_news_class where oneid="&oneid&" and twoid=0",conn,1,3
mname=rs("name")
Call CloseRs(rs)

select case action
case "add" 
set rs=server.CreateObject("adodb.recordset")
if dname<>"" then
rs.open "select * from DNJY_news_class where oneid="&oneid&" and twoid<>0 and threeid=0 and name='"&dname&"'",conn,1,3
if Not rs.eof Then
response.Write "<script language=javascript>alert('�˷�������["&dname&"]�Ѿ����ڣ�������ѡ��һ��!');location.replace('?oneid="&oneid&"');</script>"	
set rs=Nothing
response.end	
end If
set rs=server.CreateObject("adodb.recordset")
rs.open "select * from DNJY_news_class where oneid="&oneid&" order by twoid "&DN_OrderType&"",conn,1,3
two=rs("twoid")+1
rs.close
rs.open "DNJY_news_class",conn,1,3
rs.AddNew
rs("oneid")=oneid
rs("twoid")=two
rs("threeid")=0
rs("name")=dname
rs("sorting")=strint(request("sorting"))
rs.Update
Call CloseRs(rs)
end if

case "edit"
   set rs=server.CreateObject("adodb.recordset")
	rs.open "select name from DNJY_news_class where oneid="&oneid&" and twoid>0 and threeid=0 and name='"&dname&"' and id<>"&idd&"",conn,1,1
	if not rs.eof Then
	response.Write "<script language=javascript>alert('�˷�������["&dname&"]�Ѿ����ڣ�������ѡ��һ��!');location.replace('?oneid="&oneid&"');</script>"	
	set rs=Nothing
    response.end	
	end If
set rs=server.CreateObject("adodb.recordset")
rs.open "select * from DNJY_news_class where threeid=0 and twoid="&twoid&" and oneid="&oneid,conn,1,3
yname=rs("name")
rs("name")=dname
rs("sorting")=strint(request("sorting"))
rs.Update
Call CloseRs(rs)
sql="update DNJY_news set c_two='"&dname&"' where  c_oneid="&oneid&" and c_twoid="&twoid&" "
conn.execute sql


case "del"
set rs=server.CreateObject("adodb.recordset")
sql="select * from DNJY_news_class where oneid="&oneid&" and threeid>0 and twoid ="&twoid
rs.open sql,conn,3,2
if not rs.eof Then
Response.Write ("<script language=javascript>alert('�˷����¼�����ǿգ�����ɾ���¼��������ɾ��������!');history.go(-1);</script>")
set rs=Nothing
response.End
End If
set rs=server.CreateObject("adodb.recordset")
conn.execute ("delete from DNJY_news_class where oneid="&oneid&" and twoid="&twoid&"")
'ɾ��������������¼�����ļ���ʼ
Dim rsdel,sqldel,contentimg,photoimg,Strsimg
sql="select * from DNJY_news where c_oneid="&oneid&" and c_twoid="&twoid
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,3
if not rs.eof Then
do while not rs.eof
contentimg=content_pic(rs("content"))
If contentimg<>"" Then'ɾ��������ͼƬ
contentimg=split(contentimg,"|")
For Each Strsimg In contentimg
DoDelFile(Strsimg)
Next
End If
photoimg=rs("s_photo")
If photoimg<>"" Then'ɾ�������ϴ����ͼƬ
photoimg=split(photoimg,"|")
For Each Strsimg In photoimg
DoDelFile("/"&strInstallDir&Replace(Strsimg,"_s",""))'ɾ��ԭͼ
DoDelFile("/"&strInstallDir&Strsimg)'ɾ������ͼ
Next
End If
set rsdel=server.createobject("adodb.recordset")
sqldel="delete from [DNJY_news] where id="&rs("id")'ɾ���������
rsdel.open sqldel,conn,1,3
rs.movenext
if rs.eof then exit do
Loop
Call CloseRs(rs)                                 
end If
'ɾ��������������¼�����ļ���ʼEND

Call CloseDB(conn)
response.Redirect "admin_article_Class2.asp?oneid="&oneid
end select
%>
<HTML>
<HEAD>
<LINK rel="stylesheet" type="text/css" href="../css/style.css">
<META http-equiv="Content-Type" content="text/html; charset=gb2312">
<META name="GENERATOR" content="Microsoft FrontPage 4.0">
<META name="ProgId" content="FrontPage.Editor.Document">
<TITLE>�������</TITLE>
</HEAD>

<BODY background="images/background.gif">
<TABLE width="98%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#799AE1">
  <TR> 
    <TD height="20" bgcolor="#799AE1" align="center"><FONT color="#FFFFFF" style="font-size:14px">��&nbsp;��&nbsp;�� �� ��&nbsp;Ŀ&nbsp;�� ��</FONT></TD>
  </TR>
  <TR> 
    <TD bgcolor="#FFFFFF"> &nbsp;[<A href="admin_article_Class1.asp">��ҳ</A>]-&gt;<%=mname%><BR>
	<TABLE width="98%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="#D6DFF7">
        <TR align="center" bgcolor="#FFFFFF" height="20"> 
          <TD width="30">���</TD>
		  <TD width="30">��ĿID��</TD>
          <TD>��������</TD>
          <TD>��ʾ˳��</TD>
          <TD>�������</TD>
        </TR>
        <%set rs=server.CreateObject("adodb.recordset")
		  rs.Open "select * from DNJY_news_class where oneid="&oneid&" and twoid<>0 and threeid=0 order by sorting",conn,1,1
		  dim follows
		  if rs.EOF and rs.BOF then
          response.write"<tr bgcolor=#FFFFFF><td colspan='6'><p align='center'><font color='red'>���޷��࣡</font></td></tr></table><br>"
		  follows=0
		  else
		  do while not rs.EOF
		  i=i+1
		  %>
        <FORM name="form1" method="post" action="admin_article_Class2.asp?action=edit&oneid=<%=int(rs("oneid"))%>&twoid=<%=int(rs("twoid"))%>&idd=<%=rs("id")%>">
          <TR bgcolor="#FFFFFF" align="center">
		  <TD><%=i%></TD>
		  <TD><%=rs("twoid")%></TD>
            <TD><INPUT name="dname" type="text" size="30" value="<%=trim(rs("name"))%>">
            &nbsp;[<A href="admin_article_Class3.asp?oneid=<%=rs("oneid")%>&twoid=<%=rs("twoid")%>">��Ŀ¼</A>]</TD>
            <TD><INPUT name="sorting" type="text" size="3" value="<%=trim(rs("sorting"))%>"></TD>            
            <TD><INPUT type="submit" name="Submit" value="�� ��">
                &nbsp; <INPUT type="button" name="DEL" onClick="{if(confirm('ȷ��Ҫɾ�����������\n��ɾ���˷���������������º�ͼƬ\n�˲��������Իָ���')){location.href='?name=<%=rs("name")%>&action=del&oneid=<%=rs("oneid")%>&twoid=<%=rs("twoid")%>';}return false;}" value="ɾ��" >
            </TD>
          </TR>
        </FORM>
        <%
		rs.MoveNext
          loop
          follows=rs.RecordCount
          end If%>
      </TABLE>
	<BR></TD>
  </TR>
</TABLE>
<BR>
<TABLE width="98%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#799AE1">
  <TR> 
    <TD height="20" bgcolor="#799AE1" align="center"><FONT color="#FFFFFF" style="font-size:14px">�� �� ��&nbsp;��&nbsp;�� �� �� ��</FONT></TD>
  </TR>
  <TR> 
    <TD bgcolor="#FFFFFF">
	<BR>
	<TABLE width="98%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="#D6DFF7">
        <TR align="center" bgcolor="#FFFFFF" height="20"> 
          <TD width="30">��� </TD>
          <TD>��������</TD>
          <TD>��ʾ˳��</TD>
          <TD>ȷ������</TD>
        </TR>
        <FORM name="form1" method="post" action="admin_article_Class2.asp?action=add&oneid=<%=oneid%>">
          <TR align="center" bgcolor="#FFFFFF"> 
            <TD><%=rs.RecordCount+1%></TD>
	        <TD><INPUT name="dname" type="text" id="dname" size="30"></TD>
	        <TD><INPUT name="sorting" type="text" value="0" size="3"></TD>
	        <TD><INPUT type="submit" name="Submit3" value="�� ��" onClick="if(dname.value==''){alert('����������Ϊ��');return false;}"></TD>
          </TR>
        </FORM>
      </TABLE>
	<BR></TD>
  </TR>
</TABLE>
</BODY>                                                                                                                          
<%Call CloseRs(rs)
Call CloseDB(conn)%>                                                                                                                           
</HTML>  
