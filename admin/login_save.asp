<!--#include file="../conn.asp"-->
<!--#include file="../SqlIn.Asp"-->
<!--#include file="../config.asp"-->
<!--#include file=../Include/md5.asp-->
<%
'=====================================
'���칩����Ϣ��վ����ϵͳ
'��������Ƽ������ҳ�Ʒ
'�ͷ�����http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>
 <link href="../css/style.css" rel="stylesheet" type="text/css">
<%dim username,password
If Request.Cookies("status")="" Then
    Response.Write "�������Cookieû�п������޷��������б�����!"
    response.end
 end If	
	if session("dick_err")="" then
	session("dick_err")=0	
	end If

If trim(request.form("admin_codee"))="" Then
response.Write "<script language=javascript>alert('��������֤��!');location.replace('login.asp');</script>"
response.end
else
if CInt(trim(request.form("admin_codee")))<>CInt(Session("ttpSN")) then
response.Write "<script language=javascript>alert('��֤�벻��ȷ!');location.replace('login.asp');</script>"
Session("ttpSN")=""
response.end
end If
end If
Call OpenConn
	username=replace(trim(request("username")),"'","��")
	password=md5(replace(trim(request("password")),"'","��"))

	set rs=server.createobject("adodb.recordset")
	sql="select * from [DNJY_admin] where Password='"&Password&"' and UserName='"&UserName&"'"
	rs.open sql,conn,1,1
 	if not rs.bof and not rs.eof then
 		if password=rs("password") Then			
        Response.cookies("administrator")=rs("username")
		Response.cookies("dick")("username")=rs("username")'ֱ�ӵ�¼ǰ̨
        Response.cookies("dick")("domain")=Request.ServerVariables("SERVER_NAME")'ֱ�ӵ�¼ǰ̨
        Response.cookies("admindick")=1		  
        'Response.Cookies("dick").Expires=DateAdd("n",20,now())'����cookies("dick")20��������Ч
        Response.cookies("domain")=Request.ServerVariables("SERVER_NAME")
		session("admin_dick")=now()

       Response.Redirect "index.asp"
 		Else
		    session("dick_err")=session("dick_err")+1
            response.write " <b>���ȷ��ʧ��"&session("dick_err")&"�Σ���������"&htdlip&"�ν�����IP���޷����ʣ���<a href=login.asp><b><font color=#Ff0000>���µ�½</font></b></a></b> "
            
            response.end
 		end if
	Else
	        session("dick_err")=session("dick_err")+1
            response.write " <b>���ȷ��ʧ��"&session("dick_err")&"�Σ���������"&htdlip&"�ν�����IP���޷����ʣ���<a href=login.asp><b><font color=#Ff0000>���µ�½</font></b></a></b> "
            
            response.end
	end if
Call CloseRs(rs)
Call CloseDB(conn)%>
 <LINK href="../css/css.css" rel=stylesheet></LINK>

