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
dim str2
Dim ServerURL,aa,objfso,htmout,http
Dim username,biaoti,leixin,biaotixs,fbsj,scjiage,jiage,sdmba,usernameid,namea,dianhua,mobile,userqq,email,dizhi,youbian,memo,hfcs,userip,xxtp,xxmpname,wcolor,ncolor,cksj
id=trim(request("selectedid"))
if trim(id)="" then
response.write "<script language=JavaScript>" & chr(13) & "alert('û��ѡ���¼��');" &"window.location='infolist.asp'" & "</script>"
response.end
end if
str2=split(id,",")
for i=0 to ubound(str2)
set rs=server.createobject("adodb.recordset")
sql="select tj,fbsj,fsohtm,city_oneid,city_twoid,city_threeid,Readinfo from [DNJY_info] where id="&trim(str2(i))
rs.open sql,conn,1,3
rs("tj")=request("tj")
rs("fbsj")=now()
rs("fsohtm")=1
rs.update
city_oneid=rs("city_oneid")
city_twoid=rs("city_twoid")
city_threeid=rs("city_threeid")
Readinfo=rs("Readinfo")

Next

Call CloseRs(rs)
Call CloseDB(conn)
response.write "<script language=JavaScript>" & chr(13) & "alert('��Ϣ�Ƽ��ɹ���')</script>"
response.write "<meta http-equiv=refresh content=""1;URL=infolist.asp"">"
%>