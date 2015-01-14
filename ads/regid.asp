<!--#include file="../conn.asp"-->
<!--#include file="../config.asp"-->
<!--#include file=cookies.asp-->
<%
'=====================================
'天天供求信息网站管理系统
'天天网络科技工作室出品
'客服中心http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>
<%Call OpenConn
dim adId,rsadId,city_oneid,city_twoid,city_threeid
adId=trim(request("ADID"))
city_oneid=TRIM(Request("city_one"))
city_twoid=TRIM(Request("city_two"))
city_threeid=TRIM(Request("city_three"))
If city_oneid="" Then city_oneid=0
If city_twoid="" Then city_twoid=0
If city_threeid="" Then city_threeid=0

set rsadId=server.CreateObject("adodb.recordset")
rsadId.open"select adId from DNJY_ad where  adId='"&adId&"' and city_oneid="&city_oneid&" and city_twoid="&city_twoid&" and city_threeid="&city_threeid&"",conn,1,1
if not rsadId.eof then
response.Write "<font color=#FF0000>此ID号已存在，请另选</font>"
else
response.Write "<font color=#0000ff>此ID号不存在，可以注册</font>"
rsadId.close
set rsadId=Nothing
Call CloseDB(conn)
end if%>