<!--#include file="../conn.asp"-->
<!--#include file="../SqlIn.Asp"-->
<!--#include file=cookies.asp-->
<!--#include file="../config.asp"-->
<!--#include file="../Include/head.asp"-->
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
<link rel="stylesheet" type="text/css" href="../css/css.css">
<BODY>
<%Call OpenConn
Dim biaoti,leixin,biaotixs,fbsj,scjiage,jiage,sdmba,username,usernameid,namea,dianhua,mobile,userqq,email,dizhi,youbian,memo,hfcs,userip,xxtp,xxmpname,wcolor,ncolor,tj,cksj
id=trim(request("id"))
tj=trim(request("tj"))
set rs=server.createobject("adodb.recordset")
sql="select tj,city_oneid,city_twoid,city_threeid,Readinfo from [DNJY_info] where id="&id
rs.open sql,conn,1,3
if rs.eof or rs.bof then
response.write "<li>��������"
cl
response.end

end if
rs("tj")=tj
rs.update
city_oneid=rs("city_oneid")
city_twoid=rs("city_twoid")
city_threeid=rs("city_threeid")
Readinfo=rs("Readinfo")
Call CloseRs(rs)
response.write "<li>�����Ƽ��ɹ���"


If Htmlhome=1 Then Call HomeHtml()'����������ҳ
call cl()
%>
<%sub cl()%>
<body onLoad="setTimeout(window.close, 2000)">
<%end Sub
Call CloseDB(conn)%>
