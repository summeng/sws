<!--#include file="../conn.asp"-->
<!--#include file="../SqlIn.Asp"-->
<!--#include file=cookies.asp-->
<!--#include file="../config.asp"-->
<!--#include file="../Include/head.asp"-->
<!--#include file="../Include/Function.asp"-->
<!--#include file="../Include/label.asp"-->

<%
'=====================================
'天天供求信息网站管理系统
'天天网络科技工作室出品
'客服中心http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>

<%call checkmanage("05")
dim str,fso,objStream
Server.ScriptTimeout=50000
str = "<?xml version='1.0' encoding='UTF-8'?>" & vbcrlf
str = str & "<rss version=""2.0"">" & vbcrlf
str = str & "<channel>" & vbcrlf
str = str & "<title>" & webname & " RSS feed</title>" & vbcrlf 
str = str & "<link>http://"&web&"/"&strInstallDir&"</link>" & vbcrlf  
str = str & "<description>"&description&"</description>" & vbcrlf

set rs=server.createobject("ADODB.recordset")
sql = "select top 100  * from DNJY_info where yz=1 and fsohtm=1"
if jiaoyi=0 Then sql=sql&" and zt="&jiaoyi
if overdue=0 Then sql=sql&" and dqsj>"&DN_Now&""
sql=sql&" order by fbsj "&DN_OrderType&""
rs.open sql,conn,1,1
if rs.Eof then
str = str & "<item>尚未有"&webname&"最新信息</item>" & vbcrlf
end if
Do while not rs.Eof 
str = str & "<item>"
If asphtm=1 then
str = str & "<link>http://"&web&"/"&strInstallDir&"html/"&rs("id")&".html</link>" & vbcrlf
Else
str = str & "<link>http://"&web&"/"&strInstallDir&"show.asp?id="&rs("id")&"</link>" & vbcrlf
end if
str = str & "<title>"&rs("biaoti")&"</title>" & vbcrlf
str = str & "<creator>"&rs("name")&"</creator>" & vbcrlf
str = str & "<PubDate>"&rs("fbsj")&"</PubDate>" & vbcrlf
If rs("Readinfo")=1 Then
str = str & "<description><![CDATA[会员权限阅读]]></description>" & vbcrlf
elseIf rs("Readinfo")=2 Then
str = str & "<description><![CDATA[VIP会员权限阅读]]></description>" & vbcrlf
Else
str = str & "<description><![CDATA["&replace(replace(left(rs("memo"),200),"<","&lt;"),">","&gt;")&"]]></description>" & vbcrlf
End If
str = str & "</item>" & vbcrlf
rs.MoveNext          
loop                   
          Call CloseRs(rs)
		  Call CloseDB(conn)

str = str & "</channel>" & vbcrlf
str = str & "</rss>" & vbcrlf
set fso = nothing

Set objStream = Server.CreateObject("ADODB.Stream")
With objStream
.Open
.Charset = "UTF-8"
.Position = objStream.Size
.WriteText=str
.SaveToFile server.mappath("/"&strInstallDir&"xml/Rss0_0_0.xml"),2 '生成的XML文件名,不建议修改,如果你有更好的方案,可以改
.Close
End With

Set objStream = Nothing
If Not Err Then
response.Write "<script language=javascript>alert('成功生成总站RSS聚合服务页!');location.replace('../xml/Rss0_0_0.xml');</script>"
Response.End
End If
%>
<style type="text/css">
<!--
body {
	background-color: #E3EBF9;
}
-->
</style>