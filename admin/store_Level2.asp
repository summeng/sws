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
%>
<%call checkmanage("02")
Call OpenConn
dim twoid,action,id,idd,i,dname,dbpath,mname,two,yname,adid,fdel,iPage,pic,sql1,sql2,sql3
twoid=strint(request.QueryString("twoid"))
id=strint(request.QueryString("id"))
idd=strint(request.QueryString("idd"))
action=HtmlEncode(request.querystring("action"))
dname=HtmlEncode(request("dname"))
dbpath="../"

set rs=server.CreateObject("adodb.recordset")
rs.open "select * from DNJY_hy_type where id="&id&" and twoid=0",conn,1,3
mname=rs("name")
Call CloseRs(rs)

select case action
case "add" 
set rs=server.CreateObject("adodb.recordset")
if dname<>"" then
rs.open "select * from DNJY_hy_type where id="&id&" and twoid<>0 and threeid=0 and name='"&dname&"'",conn,1,3
if Not rs.eof Then
response.Write "<script language=javascript>alert('�˷�������["&dname&"]�Ѿ����ڣ�������ѡ��һ��!');location.replace('?id="&id&"');</script>"	
set rs=Nothing
response.end	
end If
set rs=server.CreateObject("adodb.recordset")
rs.open "select * from DNJY_hy_type where id="&id&" order by twoid "&DN_OrderType&"",conn,1,3
two=rs("twoid")+1
rs.close
rs.open "DNJY_hy_type",conn,1,3
rs.AddNew
rs("id")=id
rs("name")=dname
rs("indexid")=strint(request("indexid"))
rs("indexshow")=HtmlEncode(request("indexshow"))
rs("titlecolor")=HtmlEncode(request("colorshow"))
rs("twoid")=two
rs("threeid")=0
rs.Update
Call CloseRs(rs)
end if

case "edit"
   set rs=server.CreateObject("adodb.recordset")
	rs.open "select name from DNJY_hy_type where id="&id&" and name='"&dname&"' and i<>"&idd&"",conn,1,1
	if not rs.eof Then
	response.Write "<script language=javascript>alert('�˷�������["&dname&"]�Ѿ����ڣ�������ѡ��һ��!');location.replace('?id="&id&"');</script>"	
	set rs=Nothing
    response.end	
	end If
set rs=server.CreateObject("adodb.recordset")
rs.open "select * from DNJY_hy_type where threeid=0 and twoid="&twoid&" and id="&id,conn,1,3
yname=rs("name")
rs("name")=dname
rs("indexid")=strint(request("indexid"))
rs("indexshow")=HtmlEncode(request("indexshow"))
rs("titlecolor")=HtmlEncode(request("colorshow"))
rs.Update
Call CloseRs(rs)
sql="update DNJY_info set type_two='"&dname&"' where  type_oneid="&id&" and type_twoid="&twoid&" "
conn.execute sql


case "del"
set rs=server.CreateObject("adodb.recordset")
sql="select * from DNJY_hy_type where id="&id&" and threeid>0 and twoid ="&twoid
rs.open sql,conn,3,2
if not rs.eof Then
Response.Write ("<script language=javascript>alert('�˷����¼�����ǿգ�����ɾ���¼��������ɾ��������!');history.go(-1);</script>")
set rs=Nothing
response.End
End If
set rs=server.CreateObject("adodb.recordset")
conn.execute ("delete from DNJY_hy_type where id="&id&" and twoid="&twoid&"")
'ɾ�������������Ϣ������ļ���ʼ
sql="select * from DNJY_info where type_oneid="&id&" and type_twoid="&twoid
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,3,2
Set fdel = CreateObject("Scripting.FileSystemObject")	
if not rs.eof then
 For iPage = 1 To rs.recordcount
 adid=rs("id")
 pic=rs("tupian")
rs.delete
rs.update
sql1="delete from DNJY_Message where adid ="&adid
conn.execute(sql1)
rs.movenext     
   '''''''''''ɾ��html��ʼ''''''''''''''''
				 if fdel.fileExists(Server.MapPath("../html/"&trim(adid)&".html")) then
                 fdel.DeleteFile(Server.MapPath("../html/"&trim(adid)&".html"))
                 end if
				 if pic<>"0" And fdel.fileExists(Server.MapPath("../UploadFiles/infopic/"&pic&"")) then
				 fdel.DeleteFile(Server.MapPath("../UploadFiles/infopic/"&pic&""))
				 end if             

'''''''''''ɾ��html����'''''''''''''''
'ɾ����ػظ���ʼ
sql1="delete from [DNJY_reply] where xxid="&adid&" "
rs.open sql1,conn,1,3
'ɾ����ػظ�END
'ɾ����ؾٱ���ʼ
sql2="delete from [DNJY_Bad_info] where infoid="&adid&" "
rs.open sql2,conn,1,3
'ɾ����ؾٱ�END
'ɾ������ղؿ�ʼ
sql3="delete from [DNJY_favorites] where scid='"&adid&"' "
rs.open sql3,conn,1,3
'ɾ������ղ�END
if rs.eof then exit for
next
set fdel=nothing
rs.close:set rs=nothing                                 
end if
conn.close:set conn=Nothing
'ɾ�������������Ϣ������ļ�END
response.Redirect "store_Level2.asp?id="&id
end select
%>
<HTML>
<HEAD>
<LINK rel="stylesheet" type="text/css" href="../css/style.css">
<SCRIPT language="javascript" src="admin.js"></SCRIPT>
<META http-equiv="Content-Type" content="text/html; charset=gb2312">
<META name="GENERATOR" content="Microsoft FrontPage 4.0">
<META name="ProgId" content="FrontPage.Editor.Document">
<TITLE>�������</TITLE>
<style type="text/css">
<!--
body {
	background-color: #E7EEF5;
	margin-left: 0px;
	margin-top: 0px;
	margin-right: 0px;
	margin-bottom: 0px;
}
-->
</style>
</HEAD>

<BODY background="images/background.gif">
<IFRAME width="260" height="165" id="colourPalette" src="nc_selcolor.htm" style="visibility:hidden; position: absolute; left: 0px; top: 0px;border:1px gray solid" frameborder="0" scrolling="no" ></IFRAME>
<table width="98%" border="1" align="center" cellpadding="3" cellspacing="0" bordercolorlight="#335078" bordercolordark="#FFFFFF" bgcolor="#C5D5E4"  style="font-size:12px; letter-spacing:1px;">
  <tr> 
    <td height="20" bgcolor="#C5D5E4" align="center"><span style="color: #FF0000"><strong><font style="font-size:14px">�� �� �� �� �� Ŀ �� ��</font></strong></span></td>
  </tr>
  <TR> 
    <TD bgcolor="#EAEEF4"> &nbsp;[<A href="store_Level1.asp">��ҳ</A>]-&gt;<%=mname%><BR>
	<table width="98%" border="1" align="center" cellpadding="3" cellspacing="0" bordercolorlight="#335078" bordercolordark="#FFFFFF" bgcolor="#C5D5E4"  style="font-size:12px; letter-spacing:1px;">
        <TR align="center" bgcolor="#FFFFFF" height="20"> 
          <TD width="30" bgcolor="#D9E4EE">���</TD>
          <TD bgcolor="#D9E4EE">��������</TD>
          <TD bgcolor="#D9E4EE">��ʾ˳��</TD>
          <TD bgcolor="#D9E4EE">��ҳ��ʾ</TD>
		  <TD bgcolor="#D9E4EE">��ҳ��ɫ</TD>
          <TD bgcolor="#D9E4EE">�������</TD>
        </TR>
        <%set rs=server.CreateObject("adodb.recordset")
		  rs.Open "select * from DNJY_hy_type where id="&id&" and twoid<>0 and threeid=0 order by twoid",conn,1,1
		  dim follows
		  if rs.EOF and rs.BOF then
          response.write"<tr bgcolor=#FFFFFF><td colspan='6'><p align='center'><font color='red'>���޷��࣡</font></td></tr></table><br>"
		  follows=0
		  else
		  do while not rs.EOF
		  i=i+1
		  %>
        <FORM name="form1" method="post" action="store_Level2.asp?action=edit&id=<%=int(rs("id"))%>&twoid=<%=int(rs("twoid"))%>&idd=<%=rs("i")%>">
          <TR bgcolor="#FFFFFF" align="center">
		  <TD bgcolor="#F4F7F9"><%=i%></TD>
            <TD bgcolor="#F4F7F9"><INPUT name="dname" type="text" size="30" value="<%=trim(rs("name"))%>">
            &nbsp;[<A href="store_Level3.asp?id=<%=rs("id")%>&twoid=<%=rs("twoid")%>">��Ŀ¼</A>]</TD>
            <TD bgcolor="#F4F7F9"><INPUT name="indexid" type="text" size="3" value="<%=trim(rs("indexid"))%>"></TD>
            <TD bgcolor="#F4F7F9"><SELECT name="indexshow">
<OPTION value="yes" <%if rs("indexshow")="yes" then%>selected<%end if%>>��ʾ</OPTION>
<OPTION value="no" <%if rs("indexshow")="no" then%>selected<%end if%>>����</OPTION>
                    </SELECT></TD>
            <TD bgcolor="#F4F7F9"><INPUT name="colorshow" type="text" style="background:<%=rs("titlecolor")%>" onClick="Getcolor(this)" value="<%=rs("titlecolor")%>" size="7" maxlength="7"></TD>
            <TD bgcolor="#F4F7F9"><INPUT type="submit" name="Submit" value="�� ��">
                &nbsp; <INPUT type="button" name="DEL" onClick="{if(confirm('ȷ��Ҫɾ�����������\n��ɾ���˷�����������Ϣ\n�˲��������Իָ���')){location.href='?name=<%=rs("name")%>&action=del&id=<%=rs("id")%>&twoid=<%=rs("twoid")%>';}return false;}" value="ɾ��" >
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
<table width="98%" border="1" align="center" cellpadding="3" cellspacing="0" bordercolorlight="#335078" bordercolordark="#FFFFFF" bgcolor="#C5D5E4"  style="font-size:12px; letter-spacing:1px;">
  <tr> 
    <td height="20" bgcolor="#C5D5E4" align="center"><span style="color: #FF0000"><strong><font style="font-size:14px">�� �� �� �� �� �� �� Ŀ</font></strong></span></td>
  </tr>
  <TR> 
    <TD bgcolor="#EAEEF4">
	<BR>
	<TABLE width="98%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="#D6DFF7">
        <TR align="center" bgcolor="#FFFFFF" height="20"> 
          <TD width="30" bgcolor="#D9E4EE">��� </TD>
          <TD bgcolor="#D9E4EE">��������</TD>
          <TD bgcolor="#D9E4EE">��ʾ˳��</TD>
          <TD bgcolor="#D9E4EE">��ҳ��ʾ</TD>
          <TD bgcolor="#D9E4EE">��ҳ��ɫ</TD>
          <TD bgcolor="#D9E4EE">ȷ������</TD>
        </TR>
        <FORM name="form1" method="post" action="store_Level2.asp?action=add&id=<%=id%>">
          <TR align="center" bgcolor="#FFFFFF"> 
            <TD bgcolor="#F4F7F9"><%=rs.RecordCount+1%></TD>
	        <TD bgcolor="#F4F7F9"><INPUT name="dname" type="text" id="dname" size="30"></TD>
	        <TD bgcolor="#F4F7F9"><INPUT name="indexid" type="text" value="0" size="3"></TD>
	        <TD bgcolor="#F4F7F9"><SELECT name="indexshow">
                <OPTION value="yes">��ʾ</OPTION>
                <OPTION value="no" selected>����</OPTION>
              </SELECT></TD>
	        <TD bgcolor="#F4F7F9"><INPUT name="colorshow" type="text" style="background:#000000" onClick="Getcolor(this)" value="#000000" size="7" maxlength="7"></TD>
	        <TD bgcolor="#F4F7F9"><INPUT type="submit" name="Submit3" value="�� ��" onClick="if(dname.value==''){alert('����������Ϊ��');return false;}"></TD>
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
