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
<%Call OpenConn
dim action,id,idd,i,dname,dbpath,yname,adid,follows,fdel,iPage,pic,sql1,sql2,sql3
id=strint(request.QueryString("id"))
idd=strint(request.QueryString("idd"))
action=HtmlEncode(request.querystring("action"))
dname=HtmlEncode(request("dname"))
dbpath="../"

select case action
case "add"
set rs=server.CreateObject("adodb.recordset")
 rs.open "select * from DNJY_type where name='"&dname&"' and threeid=0 and twoid=0",conn,1,1
         if rs.EOF and rs.BOF then
           rs.close
    rs.open "select * from DNJY_type where threeid=0 and twoid=0 order by id "&DN_OrderType&"",conn,1,1
	if rs.eof and rs.bof then
	id=1
	else
    id=rs("id")+1
	end if
	rs.close
	rs.open "DNJY_type",conn,1,3
	rs.AddNew
if id=0 then
   rs("id")=1
   else
   rs("id")=id
    end If
    
rs("name")=dname
rs("indexid")=strint(request("indexid"))
rs("indexshow")=request("indexshow")
rs("titlecolor")=request("colorshow")
rs("twoid")=0
rs("threeid")=0
rs.Update
Call CloseRs(rs)
else
response.write "<script>window.alert('�������һ��Ŀ¼����:["&dname&"]�Ѿ����ˣ�');history.go(-1);</script>"
end if
case "edit"
   set rs=server.CreateObject("adodb.recordset")
	rs.open "select name from DNJY_type where twoid=0 and threeid=0  and name='"&dname&"' and i<>"&idd&"",conn,1,1
	if not rs.eof Then
	response.Write "<script language=javascript>alert('�˷�������["&dname&"]�Ѿ����ڣ�������ѡ��һ��!');location.replace('?');</script>"	
	set rs=Nothing
    response.end	
	end If
set rs=server.CreateObject("adodb.recordset")
rs.open "select * from DNJY_type where twoid=0 and id="&id,conn,1,3
yname=rs("name")
rs("name")=dname
rs("indexid")=strint(request("indexid"))
rs("indexshow")=request("indexshow")
rs("titlecolor")=request("colorshow")
rs.Update
Call CloseRs(rs)
sql="update DNJY_info set type_one='"&dname&"' where type_oneid="&id&" "
conn.execute sql

case "del"
set rs=server.CreateObject("adodb.recordset")
sql="select * from DNJY_type where twoid>0 and id ="&id
rs.open sql,conn,3,2
if not rs.eof Then
response.Write "<script language=javascript>alert('�˷����¼�����ǿգ�����ɾ���¼��������ɾ��������!');location.replace('?');</script>"
set rs=Nothing
response.End
End If
conn.execute ("delete from DNJY_type where id="&id&"")
'ɾ�������������Ϣ������ļ���ʼ
sql="select * from DNJY_info where type_oneid = "&id
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,3,2
Set fdel = CreateObject("Scripting.FileSystemObject")	
if not rs.eof then
 For iPage = 1 To rs.recordcount
 adid=rs("id")
 pic=rs("tupian")
rs.delete
rs.update
sql1="delete from DNJY_Message where adid ="&adid&""    
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
             
'''''''''''ɾ��html����'''''''''''''''
if rs.eof then exit for
next
set fdel=nothing
Call CloseRs(rs)                                 
end if
Call CloseDB(conn)
'ɾ�������������Ϣ������ļ�END
response.Redirect "type_Level1.asp" 
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
</HEAD>

<BODY background="images/background.gif">
<IFRAME width="260" height="165" id="colourPalette" src="nc_selcolor.htm" style="visibility:hidden; position: absolute; left: 0px; top: 0px;border:1px gray solid" frameborder="0" scrolling="no" ></IFRAME>

<TABLE width="98%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#799AE1">
  <TR> 
    <TD height="20" bgcolor="#799AE1" align="center"><FONT color="#FFFFFF" style="font-size:14px">һ&nbsp;��&nbsp;�� �� ��&nbsp;Ŀ&nbsp;�� ��</FONT></TD>
  </TR>
  <TR> 
    <TD bgcolor="#FFFFFF"> &nbsp;��ҳ <BR>
	<TABLE width="98%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="#D6DFF7">
        <TR align="center" bgcolor="#FFFFFF" height="20"> 
          <TD width="30">���</TD>
          <TD>��������</TD>
          <TD>��ʾ˳��</TD>
          <TD>��ҳ��ʾ</TD>
          <TD>��ҳ��ɫ</TD>
          <TD>�������</TD>
        </TR>
        <%set rs=server.CreateObject("adodb.recordset")
		  rs.Open "select * from DNJY_type where twoid=0 and id>0 order by id",conn,1,1		  
		  if rs.EOF and rs.BOF then
          response.write"<tr bgcolor=#FFFFFF><td colspan='6'><p align='center'><font color='red'>���޷��࣡������ӷ��ࡣ������ùٷ���Ĭ�Ϸ�����ã���</font><a href=""Classify_in_out.asp?Action=xxfl"" TARGET=right><FONT color=""#0000FF"" >����ٷ�Ĭ����Ϣ���ࡣ</font></a>��������������룬����һ����ɾ���ǳ��鷳��<a href=""http://test.ip126.com/type_list.asp"" target=_blank><FONT color=""#0000FF"" >�ٷ�Ĭ�Ϸ���ο�</font></a></td></tr></table><br>"
		  follows=0
		  else
		  do while not rs.EOF
		  i=i+1
		  %>
        <FORM name="form1" method="post" action="type_Level1.asp?action=edit&id=<%=int(rs("id"))%>&idd=<%=rs("i")%>">
          <TR bgcolor="#FFFFFF" align="center">
		  <TD><%=i%></TD>
            <TD><INPUT name="dname" type="text" size="30" value="<%=rs("name")%>">
            &nbsp;[<A href="type_Level2.asp?id=<%=rs("id")%>">��Ŀ¼</A>]</TD>
            <TD><INPUT name="indexid" type="text" size="3" value="<%=rs("indexid")%>"></TD>
            <TD><SELECT name="indexshow">
<OPTION value="yes" <%if rs("indexshow")="yes" then%>selected<%end if%>>��ʾ</OPTION>
<OPTION value="no" <%if rs("indexshow")="no" then%>selected<%end if%>>����</OPTION>
                    </SELECT></TD>
            <TD><INPUT name="colorshow" type="text" style="background:<%=rs("titlecolor")%>" onClick="Getcolor(this)" value="<%=rs("titlecolor")%>" size="7" maxlength="7"></TD>
            <TD><INPUT type="submit" name="Submit" value="�� ��">
                &nbsp; <INPUT type="button" name="DEL" onClick="{if(confirm('ȷ��Ҫɾ�����������\n��ɾ���˷�����������Ϣ\n�˲��������Իָ���')){location.href='?name=<%=rs("name")%>&action=del&id=<%=rs("id")%>';}return false;}" value="ɾ��" >
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
    <TD height="20" bgcolor="#799AE1" align="center"><FONT color="#FFFFFF" style="font-size:14px">�� �� һ&nbsp;��&nbsp;�� �� �� ��</FONT></TD>
  </TR>
  <TR> 
    <TD bgcolor="#FFFFFF">
	<BR>
	<TABLE width="98%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="#D6DFF7">
        <TR align="center" bgcolor="#FFFFFF" height="20"> 
          <TD width="30">��� </TD>
          <TD>��������</TD>
          <TD>��ʾ˳��</TD>
          <TD>��ҳ��ʾ</TD>
          <TD>��ҳ��ɫ</TD>
          <TD>ȷ������</TD>
        </TR>
        <FORM name="form1" method="post" action="type_Level1.asp?action=add">
          <TR align="center" bgcolor="#FFFFFF"> 
            <TD><%=rs.RecordCount+1%></TD>
	        <TD><INPUT name="dname" type="text" id="dname" size="30"></TD>
	        <TD><INPUT name="indexid" type="text" value="0" size="3"></TD>
	        <TD><SELECT name="indexshow">
                <OPTION value="yes" selected>��ʾ</OPTION>
                <OPTION value="no">����</OPTION>
            </SELECT></TD>
	         <TD><INPUT name="colorshow" type="text" style="background:#000000" onClick="Getcolor(this)" value="#000000" size="7" maxlength="7"></TD>
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
