<!--#include file="../conn.asp"-->
<!--#include file="../SqlIn.Asp"-->
<!--#include file=cookies.asp-->
<!--#include file="../config.asp"-->
<!--#include file="../Include/head.asp"-->
<!--#include file="../Include/Function.asp"-->
<!--#include file="../Include/label.asp"-->
<!--#include file="../Include/md5.asp"-->
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
If left(web,4)="www." Then web=Mid(""&web&"",4)
dim action,idd,dcity,mname,twoid,dcityadmin,dcitypass,dbpath,ycity,fdel,iPage,adid,pic,sql1,sql2,sql3,two
twoid=strint(request.QueryString("twoid"))
id=strint(request.QueryString("id"))
idd=strint(request.QueryString("idd"))
action=HtmlEncode(request.querystring("action"))
dcity=HtmlEncode(request("city"))
dcityadmin=HtmlEncode(request("cityadmin"))
dcitypass=HtmlEncode(request.form("citypass"))
dname=LCase(HtmlEncode(Request.form("dname")))
if dcity="" and action<>"" then
	response.write "<script LANGUAGE='javascript'>alert('��������д��');history.go(-1);</script>"
	response.end
end If
WebSetting =""
For i=0 To 12
	WebSetting = WebSetting & HtmlEncode(Request("w"&i)) & "�ӡӡ�"
Next
WebSetting = Left(WebSetting,Len(WebSetting)-3)
dbpath="../"

set rs=server.CreateObject("adodb.recordset")
rs.open "select * from DNJY_city where id="&id&" and twoid=0",conn,1,3
mname=rs("city")
Call CloseRs(rs)


select case action
case "add"
If dcity<>"" Then  
   set rs=server.CreateObject("adodb.recordset")
	rs.open "select city from DNJY_city where id="&id&" and threeid=0  and city='"&dcity&"' ",conn,1,1
	if not rs.eof Then
	response.Write "<script language=javascript>alert('�˵�������["&dcity&"]�Ѿ����ڣ�������ѡ��һ��!');location.replace('javascript:history.go(-1)');</script>"	
	set rs=Nothing
    response.end	
	end If
end If	
If dname<>"" Then  
   set rs=server.CreateObject("adodb.recordset")
	rs.open "select dname from DNJY_city where dname='"&dname&"' ",conn,1,1
	if not rs.eof Then
	response.Write "<script language=javascript>alert('��������["&dname&"]�Ѿ����ڣ�������ѡ��һ��!');location.replace('javascript:history.go(-1)');</script>"		
	set rs=Nothing
    response.end	
	end If
end If
If dcityadmin<>"" Then  
    set rs=server.CreateObject("adodb.recordset")
	sql="select a.cityadmin,b.username from DNJY_city a,DNJY_user b where a.cityadmin='"&dcityadmin&"' or b.username='"&dcityadmin&"'"
	rs.open sql,conn,1,1
	if not rs.eof Then
	response.Write "<script language=javascript>alert('�˵�������Ա�û�����["&dcityadmin&"]�ڵ�������Ա��ǰ̨��Ա���Ѿ����ڣ�������ѡ��һ��!');location.replace('javascript:history.go(-1)');</script>"	
	set rs=Nothing
    response.end	
	end If
end If 

If dcityadmin<>"" And dcitypass="" Then  
	response.write "<script LANGUAGE='javascript'>alert('��д��������Ա������д���룡');history.go(-1);</script>"
	response.end
end If 		
if dcity<>"" then 
set rs=server.CreateObject("adodb.recordset")
rs.open "select * from DNJY_city where id="&id&" and twoid<>0 and threeid=0 and city='"&dcity&"'",conn,1,3
if rs.eof then
rs.close
rs.open "select * from DNJY_city where id="&id&" order by twoid "&DN_OrderType&"",conn,1,3
two=rs("twoid")+1
rs.close
rs.open "DNJY_city",conn,1,3
rs.AddNew
rs("id")=id
rs("city")=dcity
rs("dname")=dname
rs("cityadmin")=dcityadmin
rs("citypass")=md5(dcitypass)
rs("indexid")=strint(request("indexid"))
rs("indexshow")=HtmlEncode(request("indexshow"))
Rs("WebSetting")=WebSetting
rs("twoid")=two
rs("threeid")=0
i=rs("i")
city_one=conn.Execute("Select city From DNJY_city Where id="&id&"")(0)
rs.Update
Call CloseRs(rs)
end if  
end If

If dcityadmin<>"" then
set rs = Server.CreateObject("ADODB.RecordSet")'ͬʱע��ǰ̨��Ա
sql="select username,password,useryz,zcdata,vipdq,city_oneid,city_one,city_twoid,city_two,city_threeid,cid from DNJY_user"
rs.open sql,conn,1,3
rs.addnew
rs("username")=dcityadmin
rs("password")=md5(dcitypass)
rs("useryz")=1
rs("city_oneid")=id
rs("city_one")=city_one
rs("city_twoid")=two
rs("city_two")=dcity
rs("city_threeid")=0
rs("cid")=i
rs("zcdata")=now()
rs("vipdq")=now()
rs.update
Call CloseRs(rs)
End If

case "edit"
If dcity<>"" Then  
   set rs=server.CreateObject("adodb.recordset")
	rs.open "select city from DNJY_city where id="&id&" and city='"&dcity&"' and i<>"&idd&"",conn,1,1
	if not rs.eof Then
	response.Write "<script language=javascript>alert('�˵�������["&dcity&"]�Ѿ����ڣ�������ѡ��һ��!');location.replace('javascript:history.go(-1)');</script>"		
	set rs=Nothing
    response.end	
	end If
end If	
If dname<>"" Then  
   set rs=server.CreateObject("adodb.recordset")
	rs.open "select dname from DNJY_city where dname='"&dname&"' and i<>"&idd&"",conn,1,1
	if not rs.eof Then
	response.Write "<script language=javascript>alert('��������["&dname&"]�Ѿ����ڣ�������ѡ��һ��!');location.replace('javascript:history.go(-1)');</script>"	
	set rs=Nothing
    response.end	
	end If
    end If

If dcityadmin<>"" Then  
    set rs=server.CreateObject("adodb.recordset")
	sql="select a.cityadmin,b.username from DNJY_city a,DNJY_user b where (a.cityadmin='"&dcityadmin&"' or b.username='"&dcityadmin&"') and a.i<>"&idd&" and b.cid<>"&idd&" "
	rs.open sql,conn,1,1
	if not rs.eof Then
	response.Write "<script language=javascript>alert('�˵�������Ա�û�����["&dcityadmin&"]�ڵ�������Ա��ǰ̨��Ա���Ѿ����ڣ�������ѡ��һ��!');location.replace('javascript:history.go(-1)');</script>"	
	set rs=Nothing
    response.end	
	end If
end If 

If dcityadmin<>"" And dcitypass="" Then  
	response.write "<script LANGUAGE='javascript'>alert('��д��������Ա������д���룡');history.go(-1);</script>"
	response.end
end If
		
set rs=server.CreateObject("adodb.recordset")
rs.open "select * from DNJY_city where threeid=0 and twoid="&twoid&" and id="&id,conn,1,3
ycity=rs("city")
rs("city")=dcity
rs("dname")=dname
rs("cityadmin")=dcityadmin
rs("indexid")=strint(request("indexid"))
rs("indexshow")=HtmlEncode(request("indexshow"))
Rs("WebSetting")=WebSetting
if dcitypass<>rs("citypass") Or (dcitypass<>"" And IsNull(rs("citypass"))) then
rs("citypass")=md5(dcitypass)
end If
If dcityadmin="" Then rs("citypass")=""
rs("AppId")=HtmlEncode(Request.form("AppId"))
rs("AppKey")=HtmlEncode(Request.form("AppKey"))
rs("API_QQCallBack")=HtmlEncode(Request.form("API_QQCallBack"))
i=rs("i")
city_one=conn.Execute("Select city From DNJY_city Where id="&id&"")(0)
rs.Update
Call CloseRs(rs)
sql="update DNJY_info set city_two='"&dcity&"' where city_threeid=0 and city_twoid="&twoid&" and city_oneid="&id
conn.execute sql

set rs = Server.CreateObject("ADODB.RecordSet")'���ǰ̨û�˻�Աͬʱע��ǰ̨��Ա
sql="select username,password,useryz,zcdata,vipdq,city_oneid,city_one,city_twoid,city_two,city_threeid,cid from DNJY_user where username='"&dcityadmin&"'"
rs.open sql,conn,1,3
if (rs.eof or rs.bof) And dcityadmin<>"" then
rs.addnew
rs("username")=dcityadmin
rs("password")=md5(dcitypass)
rs("useryz")=1
rs("city_oneid")=id
rs("city_one")=city_one
rs("city_twoid")=twoid
rs("city_two")=dcity
rs("city_threeid")=0
rs("cid")=i
rs("zcdata")=now()
rs("vipdq")=now()
rs.update
End If
Call CloseRs(rs)

set rs = Server.CreateObject("ADODB.RecordSet")'���ǰ̨�д˻�Աͬʱ�޸�ǰ̨��Ա
sql="select username,password,useryz,zcdata,vipdq,city_oneid,city_one,city_twoid,city_two,city_threeid,cid from DNJY_user where username='"&dcityadmin&"'"
rs.open sql,conn,1,3
if (Not rs.eof and Not rs.bof) And dcityadmin<>"" then
rs("username")=dcityadmin
rs("password")=md5(dcitypass)
rs("useryz")=1
rs("city_oneid")=id
rs("city_one")=city_one
rs("city_twoid")=twoid
rs("city_two")=dcity
rs("city_threeid")=0
rs("cid")=i
rs("zcdata")=now()
rs("vipdq")=now()
rs.update
End If
Call CloseRs(rs)

case "del"
set rs=server.CreateObject("adodb.recordset")
sql="select * from DNJY_city where id="&id&" and threeid>0 and twoid ="&twoid
rs.open sql,conn,3,2
if not rs.eof Then
Response.Write ("<script language=javascript>alert('�˵����¼�����ǿգ�����ɾ���¼��������ɾ��������!');history.go(-1);</script>")
set rs=Nothing
response.End
End If

set rs=server.CreateObject("adodb.recordset")
conn.execute ("delete from DNJY_city where id="&id&" and twoid="&twoid&"")
'ɾ��������������Ϣ������ļ���ʼ
sql="select * from DNJY_info where city_oneid="&id&" and city_twoid="&twoid
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,3,2
Set fdel = CreateObject("Scripting.FileSystemObject")	
if not rs.eof then
 For iPage = 1 To rs.recordcount
 adid=rs("id")
 pic=rs("tupian")
rs.delete
rs.update
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
end if
rs.close
'ɾ��������������Ϣ������ļ�END

'ɾ���������������Ź��漰����ļ���ʼ
set rs=server.CreateObject("adodb.recordset")
sql="select * from DNJY_news where city_oneid="&id&" and city_twoid="&twoid
rs.open sql,conn,3,2
if not rs.eof then
 For iPage = 1 To rs.recordcount
 adid=rs("id")
 pic=rs("firstImageName")
rs.delete
rs.update

rs.movenext  
   '''''''''''ɾ��html��ʼ''''''''''''''''
   				 if fdel.fileExists(Server.MapPath("../news/"&trim(adid)&".html")) then
                 fdel.DeleteFile(Server.MapPath("../news/"&trim(adid)&".html"))
                 end if
				 if pic<>"" and lcase(left(pic,7))<>"http://" And fdel.fileExists(Server.MapPath("../"&pic&"")) then
				 fdel.DeleteFile(Server.MapPath("../"&pic&""))
				 end if           
'''''''''''ɾ��html����'''''''''''''''
if rs.eof then exit for
next
end If
rs.close
'ɾ���������������Ź��漰����ļ�END

'ɾ�����������Ķ���114��ʼ
set rs=server.CreateObject("adodb.recordset")
sql="select * from DNJY_114 where city_oneid ="&id&" and city_twoid="&twoid 
rs.open sql,conn,3,2
if not rs.eof then
For iPage = 1 To rs.recordcount
rs.delete
rs.update
rs.movenext 
if rs.eof then exit for end if
next
end If
rs.close
'ɾ�����������Ķ���114END

'ɾ���������������Կ�ʼ
set rs=server.CreateObject("adodb.recordset")
sql="select * from DNJY_gbook where city_oneid ="&id&" and city_twoid="&twoid
rs.open sql,conn,3,2
if not rs.eof then
For iPage = 1 To rs.recordcount
rs.delete
rs.update
rs.movenext 
if rs.eof then exit for end if
next
end If
rs.close
'ɾ����������������END
'ɾ�������������������ӿ�ʼ
set rs=server.CreateObject("adodb.recordset")
sql="select * from DNJY_link where city_oneid ="&id&" and city_twoid="&twoid
rs.open sql,conn,3,2
if not rs.eof then
For iPage = 1 To rs.recordcount
rs.delete
rs.update
rs.movenext 
if rs.eof then exit for end if
next
end If

'ɾ�����������ĺ�����վEND
set fdel=nothing
rs.close:set rs=nothing                                 
conn.close:set conn=nothing 
response.Redirect "?id="&id
end select
%>
<HTML>

<HEAD>
<META http-equiv="Content-Type" content="text/html; charset=gb2312">
<META name="GENERATOR" content="Microsoft FrontPage 4.0">
<META name="ProgId" content="FrontPage.Editor.Document">
<TITLE>�����������</TITLE>
<LINK rel="stylesheet" type="text/css" href="../css/style.css">
<script language="javascript">
<!--//
function show(v){
if(document.getElementById(v).style.display=='none')
document.getElementById(v).style.display='';
else
document.getElementById(v).style.display='none';
}			
//-->
</script>
</HEAD>
<BODY>

<TABLE width="98%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#799AE1">
  <TR> 
    <TD height="20" bgcolor="#799AE1" align="center"><FONT color="#FFFFFF" style="font-size:14px">��&nbsp;��&nbsp;�� �� �� �� �� ��</FONT></TD>
  </TR>
  <TR> 
    <TD bgcolor="#FFFFFF"> &nbsp;&nbsp;[<A href="city_Level1.asp">��ҳ</A>]-&gt;<%=mname%><BR>
	  <TABLE width="98%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="#D6DFF7">
        <TR align="center" bgcolor="#FFFFFF" height="20"> 
          <TD width="30">���</TD>
		  <TD width="30">ID��</TD>
          <TD width="200">��������</TD>
          <TD width="40">��ʾ˳��</TD>
          <TD width="50">��ҳ��ʾ</TD>
		  <TD width="35">������</TD>
          <TD width="100">��������Ա</TD>
          <TD width="100">������������<bR>��8λ���£�</TD>
          <TD width="150">�������</TD>
        </TR>
        <%set rs=server.CreateObject("adodb.recordset")
		  rs.Open "select * from DNJY_city where id="&id&" and twoid<>0 and threeid=0 order by twoid",conn,1,1
		  dim follows
		  if rs.EOF and rs.BOF then
          response.write"<tr bgcolor=#FFFFFF><td colspan='9'><p align='center'><font color='red'>���޵������࣡</font></td></tr></table><br>"
		  follows=0
		  else
		  do while not rs.EOF
				WebSetting = Empty
				If Rs("WebSetting")<>"" And Not IsNull(Rs("WebSetting")) Then
					WebSetting=Split(Rs("WebSetting"),"�ӡӡ�")
				Else
					ReDim WebSetting(12)
				End If
		  i=i+1
		  %>
        <FORM name="form1" method="post" action="?action=edit&id=<%=int(rs("id"))%>&twoid=<%=int(rs("twoid"))%>&idd=<%=rs("i")%>">
          <TR bgcolor="#FFFFFF" align="center"> 
            <TD><%=i%></TD>
			<TD><%=trim(rs("twoid"))%></TD>
            <TD>
              <INPUT name="city" type="text" id="city" value="<%=trim(rs("city"))%>" size="30">
              [<A href="city_Level3.asp?id=<%=rs("id")%>&twoid=<%=rs("twoid")%>">��Ŀ¼</A>]</TD>
            <TD>
              <INPUT name="indexid" type="text" size="3" value="<%=trim(rs("indexid"))%>">
            </TD>
            <TD>
              <SELECT name="indexshow">
                <OPTION value="yes" <%if rs("indexshow")="yes" then%>selected<%end if%>>��ʾ</OPTION>
                <OPTION value="no" <%if rs("indexshow")="no" then%>selected<%end if%>>����</OPTION>
              </SELECT>
            </TD>
            <TD>
              <%if rs("dname")<>"" Then
              response.write"<font color=#ff0000>��</font>"
              Else
              response.write"��"
			  end if%>
            </TD>  			
            <TD>
              <INPUT name="cityadmin" type="text" id="cityadmin" value="<%=trim(rs("cityadmin"))%>" size="15">
            </TD>
            <TD>
              <INPUT name="citypass" type="password" id="citypass" value="<%=trim(rs("citypass"))%>" size="15">
            </TD>
            <TD>
              <INPUT type="submit" name="Submit" value="�� ��">
              &nbsp; 
              <INPUT type="button" name="DEL" onClick="{if(confirm('ȷ��Ҫɾ���������������\n��ɾ���˷�����������Ϣ\n�˲��������Իָ���')){location.href='?city=<%=rs("city")%>&id=<%=rs("id")%>&twoid=<%=rs("twoid")%>&action=del';}return false;}" value="ɾ��" >  
			  &nbsp;&nbsp;<input type="button" value="�� ��" onClick="show('s<%=Rs("twoid")%>');">
            </TD>
          </TR>
          <TR bgcolor="#FFFFFF" align="center" id="s<%=Rs("twoid")%>" style="display:none">
           <TD colspan="9">
					 <table width="700" border="0" cellpadding="0" cellspacing="0" align="center">
							<TR bgcolor="#FFFFFF" align="center">
							 <TD align="right" bgcolor="#C2D3FC">��վLOGO���ƣ�</TD>
							 <TD align="left" bgcolor="#C2D3FC"><font color="ff0000"><%=Rs("id")%>_<%=Rs("twoid")%>_<%=Rs("threeid")%>.jpg</font> ��200��80������/UploadFiles/logoĿ¼�£�</TD>
					       </TR>
					       </TR>
							<TR bgcolor="#FFFFFF" align="center">
							 <TD align="right" bgcolor="#C2D3FC"><font color="ff0000">�ر�ע�⣺</font></TD>
							 <TD align="left" bgcolor="#C2D3FC" width="500"><font color="ff0000">���º�ɫ��Ŀ��ֻҪ��д����һ���վ��ϵ��Ϣ�Ƚ���ʾ����վ�ģ���˷�վ�����е����ݽ�����д���������򣬷�վ������Ϣ��ʾ��ȱ��ȫ�����ȫ������ʾ��վ�ġ�</font></TD>
					       </TR>
					 		<TR bgcolor="#FFFFFF" align="center">
							 <TD align="right" bgcolor="#C2D3FC">����������</TD>
							 <TD align="left" bgcolor="#C2D3FC"><input name="dname" type="text" id="dname" maxlength="50" value="<%=Rs("dname")%>"> ǰ����"http://"���󲻴�"/" <font color='red'>���������÷�վ����Ҫ��д��</font></TD>
					       </TR>

							<TR bgcolor="#FFFFFF" align="center">
							 <TD align="right" bgcolor="#C2D3FC">��վ���ƣ�</TD>
							 <TD align="left" bgcolor="#C2D3FC"><input type="text" name="w0" type="text" id="w0" value="<%=WebSetting(0)%>"></TD>
						  </TR>
							<TR bgcolor="#FFFFFF" align="center">
							 <TD align="right" bgcolor="#C2D3FC">������ʾ��</TD>
							 <TD align="left" bgcolor="#C2D3FC"><input type="text" name="w1" type="text" id="w1" value="<%=WebSetting(1)%>"> ���ڱ�����ʾ��������վ����ͬ</TD>
						  </TR>
							<TR bgcolor="#FFFFFF" align="center">
							 <TD align="right" bgcolor="#C2D3FC">��վ˵����</TD>
							 <TD align="left" bgcolor="#C2D3FC"><input name="w2" type="text" id="w2" value="<%=WebSetting(2)%>"></TD>
						  </TR>
							<TR bgcolor="#FFFFFF" align="center">
							 <TD align="right" bgcolor="#C2D3FC">�����ؼ��ʣ�</TD>
							 <TD align="left" bgcolor="#C2D3FC"><input name="w3" type="text" id="w3" value="<%=WebSetting(3)%>"> ���������������Ĺؼ����ݣ��ؼ����á�,���ŷָ�</TD>						  </TR>

							<TR bgcolor="#FFFFFF" align="center">
							 <TD align="right" bgcolor="#C2D3FC">�ʼ���ַ��</TD>
							 <TD align="left" bgcolor="#C2D3FC"><input name="w4" type="text" id="w4" value="<%=WebSetting(4)%>"></TD>
						  </TR>
							<TR bgcolor="#FFFFFF" align="center">
							 <TD align="right" bgcolor="#C2D3FC">��ϵ�绰��</TD>
							 <TD align="left" bgcolor="#C2D3FC"><input name="w5" type="text" id="w5" value="<%=WebSetting(5)%>"></TD>
						  </TR>
							<TR bgcolor="#FFFFFF" align="center">
							 <TD align="right" bgcolor="#C2D3FC">��ϵ�ֻ���</TD>
							 <TD align="left" bgcolor="#C2D3FC"><input name="w6" type="text" id="w6" value="<%=WebSetting(6)%>"></TD>
						  </TR>
							<TR bgcolor="#FFFFFF" align="center">
							 <TD align="right" bgcolor="#C2D3FC">������룺</TD>
							 <TD align="left" bgcolor="#C2D3FC"><input name="w7" type="text" id="w7" value="<%=WebSetting(7)%>"></TD>
						  </TR>
							<TR bgcolor="#FFFFFF" align="center">

							 <TD align="right" bgcolor="#C2D3FC">��˾���ƣ�</TD>
							 <TD align="left" bgcolor="#C2D3FC"><input name="w12" type="text" id="w12" value="<%=WebSetting(12)%>"></TD>
						  </TR>							<TR bgcolor="#FFFFFF" align="center">
							 <TD align="right" bgcolor="#C2D3FC">��ϵ��ַ��</TD>
							 <TD align="left" bgcolor="#C2D3FC"><input name="w8" type="text" id="w8" value="<%=WebSetting(8)%>"></TD>
						  </TR>
							<TR bgcolor="#FFFFFF" align="center">
							 <TD align="right" bgcolor="#C2D3FC">�������룺</TD>
							 <TD align="left" bgcolor="#C2D3FC"><input name="w9" type="text" id="w9" value="<%=WebSetting(9)%>"></TD>
						  </TR>
							<TR bgcolor="#FFFFFF" align="center">
							 <TD align="right" bgcolor="#C2D3FC">����QQ��</TD>
							 <TD align="left" bgcolor="#C2D3FC"><input name="w10" type="text" id="w10" value="<%=WebSetting(10)%>"> ���QQ�������Ķ��Ÿ�����QQ����������ǳ�һһ��Ӧ</TD>
						  </TR>
							<TR bgcolor="#FFFFFF" align="center">
							 <TD align="right" bgcolor="#C2D3FC">QQ�ǳƣ�</TD>
							 <TD align="left" bgcolor="#C2D3FC"><input name="w11" type="text" id="w11" value="<%=WebSetting(11)%>"> ����ǳ������Ķ��Ÿ������ǳ��������QQ��һһ��Ӧ</TD>
						  </TR>
							<TR bgcolor="#FFFFFF" align="center">
							 <TD align="right" bgcolor="#C2D3FC">����������������</TD>
							 <TD align="left" bgcolor="#C2D3FC">��������������������������������������������������������������������</font></TD>
					       </TR>
							<TR bgcolor="#FFFFFF" align="center">
							 <TD align="right" bgcolor="#C2D3FC"></TD>
							 <TD align="left" bgcolor="#C2D3FC"><font color="0000ff">��ԱQQ��ݵ�¼��ڲ���</font></TD>
					       </TR>
							<TR bgcolor="#FFFFFF" align="center">
							 <TD align="right" bgcolor="#C2D3FC"><font color="0000ff">QQ��¼AppID��</font></TD>
							 <TD align="left" bgcolor="#C2D3FC"><input name="AppId" type="password" size="40" id="AppId" maxlength="50" value="<%=trim(Rs("AppId"))%>"> <font color="red">opensns.qq.com ���뵽��appid,<a href="http://connect.qq.com/" target="_blank">�������</a>��</font></TD>
					       </TR>
							<TR bgcolor="#FFFFFF" align="center">
							 <TD align="right" bgcolor="#C2D3FC"><font color="0000ff">QQ��¼AppKey��</font></TD>
							 <TD align="left" bgcolor="#C2D3FC"><input name="AppKey" type="password" size="40" id="AppKey" maxlength="50" value="<%=Rs("AppKey")%>"> <font color="red">opensns.qq.com ���뵽��appkey��</font></TD>
					       </TR>
							<TR bgcolor="#FFFFFF" align="center">
							 <TD align="right" bgcolor="#C2D3FC"><font color="0000ff">QQ��¼����ת�ĵ�ַ��</font></TD>
							 <TD align="left" bgcolor="#C2D3FC"><input name="API_QQCallBack" type="text" size="60" id="API_QQCallBack"  readonly maxlength="150" value="<%="http://"&Rs("dname")&"/"&strInstallDir&"api/qq/user.asp"%>"> <font class="tips">QQ��¼�ɹ�����ת�ĵ�ַ,���ɸġ�</font></TD>
					       </TR>
							<TR bgcolor="#FFFFFF" align="center">
							 <TD align="right"></TD>
							 <TD align="left" height="25"></font></TD>
					       </TR>
					 </table>				 
					 
					 </TD>
          </TR>
        </FORM>
        <%
		rs.MoveNext
          loop
          follows=rs.RecordCount
          end if
          Call CloseRs(rs)
		  Call CloseDB(conn)%>
      </TABLE>
	<BR></TD>
  </TR>
</TABLE>
<BR>
<TABLE width="98%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#799AE1">
  <TR> 
    <TD height="20" bgcolor="#799AE1" align="center"><FONT color="#FFFFFF" style="font-size:14px">�� �� ��&nbsp;��&nbsp;�� ��</FONT></TD>
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
          <TD>��������<br><font color='red'>���������÷�վ����Ҫ��д����������</font></TD>
          <TD>��������Ա<br><font color='red'>���������÷�վ����ɲ���д��</font></TD>
          <TD>������������<br><font color='red'>���������÷�վ����ɲ���д��</font><bR>��8λ���£�</TD>
          <TD>ȷ������</TD>
        </TR>
        <FORM name="form1" method="post" action="?action=add&id=<%=id%>">
          <TR align="center" bgcolor="#FFFFFF"> 
            <TD><%=follows+1%></TD>
            <TD>
              <INPUT name="city" type="text" id="city" size="30">
            </TD>
            <TD>
              <INPUT name="indexid" type="text" value="0" size="3">
            </TD>
            <TD>
              <SELECT name="indexshow">
                <OPTION value="yes">��ʾ</OPTION>
                <OPTION value="no" selected>����</OPTION>
              </SELECT>
            </TD>
           <TD>
              <INPUT name="dname" type="text" size="20" value="">
            </TD>
            <TD>
              <INPUT name="cityadmin" type="text" id="cityadmin" size="15">
            </TD>
            <TD>
              <INPUT name="citypass" type="password" id="citypass" size="15">
            </TD>
            <TD>
              <INPUT type="submit" name="Submit3" value="�� ��">
            </TD>
          </TR>
        </FORM>
      </TABLE>
	<BR></TD>
  </TR>
</TABLE>
</BODY> 
</HTML> 
