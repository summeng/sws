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
<%Call OpenConn
Dim username,biaoti,leixin,biaotixs,fbsj,scjiage,jiage,sdmba,usernameid,namea,dianhua,mobile,userqq,email,dizhi,youbian,memo,hfcs,userip,xxtp
dim str2
Dim ServerURL,aa,objfso,htmout,http
id=trim(request("selectedid"))
if trim(id)="" then
response.write "<script language=JavaScript>" & chr(13) & "alert('û��ѡ���¼��');" &"window.location='infolist.asp'" & "</script>"
response.end
end if
str2=split(id,",")
for i=0 to ubound(str2)
set rs=server.createobject("adodb.recordset")
sql="select tj,fsohtm,city_oneid,city_twoid,city_threeid from [DNJY_info] where id="&trim(str2(i))
rs.open sql,conn,1,3
if rs("tj")=1 or rs("tj")=2 then
rs("tj")=0
rs("fsohtm")=1
rs.update
city_oneid=rs("city_oneid")
city_twoid=rs("city_twoid")
city_threeid=rs("city_threeid")
end If



Next
If Htmlhome=1 Then Call HomeHtml()'����������ҳ
Call CloseRs(rs)
Call CloseDB(conn)
response.write "<script language=JavaScript>" & chr(13) & "alert('�ɹ�ȡ���Ƽ���')</script>"
response.write "<meta http-equiv=refresh content=""1;URL=infolist.asp"">"
%>