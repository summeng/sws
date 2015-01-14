<!--#include file="../conn.asp"-->
<!--#include file="../config.asp"-->
<%
'=====================================
'天天供求信息网站管理系统
'天天网络科技工作室出品
'客服中心http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>
<%Call OpenConn
	Dim ADID,AD_Link
	ADID=Request.QueryString("adid")
Dim ttccads,c,c1,c2,c3
If trim(request("c"))<>"" then
c=trim(request("c"))
c=split(c,",")
If c(0)="" Then c(0)=0
c1=strint(c(0))
If ubound(c)<1 Then
c2=0
else
c2=strint(c(1))
End if
If ubound(c)<2 Then
c3=0
else
c3=strint(c(2))
End If
If c1="" Then c1=0
If c2="" Then c2=0
If c3="" Then c3=0
IF c3>0 Then
ttccads=" and city_oneid="&c1&" and city_twoid="&c2&" and city_threeid="&c3
ElseIF c2>0 Then
ttccads=" and city_oneid="&c1&" and city_twoid="&c2
ElseIF c1>0 Then
ttccads=" and city_oneid="&c1
End If
End if
If shows8=0 Then ttccads=""
	Set rs=server.createobject("adodb.recordset")
		sql = "Select ADLink,ADHits from [DNJY_ad] where ADID='" & ADID & "'"&ttccads&""
		rs.open sql,conn,1,3
	If not (rs.bof or rs.eof) Then
		AD_Link=rs("ADLink")
		rs("ADHits")=rs("ADHits") + 1
		rs.Update
	else
	response.write "<font size=2 color=red>没有找到您要浏览的广告</font>"
	End If
	rs.Close
	set rs=nothing
    Call CloseDB(conn)
	If AD_Link="" Then AD_Link="http://"&Request.Servervariables("server_name")
	response.redirect AD_Link

'**************************************************
'函数名：strint
'作  用：ID号转为字数
'参  数：id ---要处理的字符串'             
'**************************************************
Function strint(id)
If IsNumeric(id) = 0 or id = "" Then id = 0
strint = Clng(id)
End Function%>