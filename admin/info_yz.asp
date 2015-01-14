<!--#include file="../conn.asp"-->
<!--#include file="../SqlIn.Asp"-->
<!--#include file="../config.asp"-->
<!--#include file="../Include/head.asp"-->
<!--#include file="../Include/Function.asp"-->
<!--#include file="../Include/label.asp"-->
<!--#include file=cookies.asp-->
<%
'=====================================
'天天供求信息网站管理系统
'天天网络科技工作室出品
'客服中心http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>
<%Dim username,biaoti,leixin,biaotixs,fbsj,scjiage,jiage,sdmba,usernameid,namea,xxmpname,dianhua,mobile,userqq,email,dizhi,youbian,memo,hfcs,userip,xxtp
dim str2,qxyz,urlpage,wcolor,ncolor,cksj
dim rss,str,fso,objStream
Server.ScriptTimeout=50000
id=trim(request("selectedid"))
if trim(id)="" then
response.write "<script language=JavaScript>" & chr(13) & "alert('没有选择记录！');" &"window.location='infolist.asp'" & "</script>"
response.end
end if
qxyz=trim(request("qxyz"))
Dim ServerURL,aa,objfso,htmout,http
Call OpenConn
str2=split(id,",")
for i=0 to ubound(str2)
set rs=server.createobject("adodb.recordset")
sql="select yz,fbsj,city_oneid,city_twoid,city_threeid,fsohtm,Readinfo from [DNJY_info] where id="&trim(str2(i))
rs.open sql,conn,1,3

if qxyz="no" then
rs("yz")=0
rs("fbsj")=now()
rs("fsohtm")=0
city_oneid=rs("city_oneid")
city_twoid=rs("city_twoid")
city_threeid=rs("city_threeid")
rs.update
'================删除已生成的文件
Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
if objFSO.fileExists(Server.MapPath("../html/"&trim(str2(i))&".html")) then
objFSO.DeleteFile(Server.MapPath("../html/"&trim(str2(i))&".html"))
end if
set objfso=Nothing
'===============================
else
rs("yz")=1
rs("fbsj")=now()
rs("fsohtm")=1
rs.update
city_oneid=rs("city_oneid")
city_twoid=rs("city_twoid")
city_threeid=rs("city_threeid")
Readinfo=rs("Readinfo")


end If

'===============重新生成相关的rss开始=====================================
   If city_oneid="" Then city_oneid=0
   If city_twoid="" Then city_twoid=0
   If city_threeid="" Then city_threeid=0
   If city_oneid=0 Then city_one=webname
   IF city_oneid=0 then
	ttcc=""
	elseIF city_threeid>0 then
	ttcc=" and city_oneid="&city_oneid&" and city_twoid="&city_twoid&" and city_threeid="&city_threeid
	elseIF city_twoid>0 then
	ttcc=" and city_oneid="&city_oneid&" and city_twoid="&city_twoid
	Else
	ttcc=" and city_oneid="&city_oneid
	End If
str = "<?xml version='1.0' encoding='UTF-8'?>" & vbcrlf
str = str & "<rss version=""2.0"">" & vbcrlf
str = str & "<channel>" & vbcrlf
str = str & "<title />" & vbcrlf 
str = str & "<link>http://"&web&"</link>" & vbcrlf  
str = str & "<description>"&city_one&city_two&city_three&"供求信息</description>" & vbcrlf
Call OpenConn
set rss=server.createobject("ADODB.recordset")
sql = "select top 50 * from DNJY_info where yz=1 and fsohtm=1"&ttcc&""
if jiaoyi=0 Then sql=sql&" and zt="&jiaoyi
if overdue=0 Then sql=sql&" and dqsj>"&DN_Now&""
sql=sql&" order by fbsj "&DN_OrderType&""
rss.open sql,conn,1,1

if rss.Eof then
str = str & "<item>尚未有最新信息</item>" & vbcrlf
end if
Do while not rss.Eof 
str = str & "<item>"
If asphtm=1 then
str = str & "<link>http://"&web&"/"&strInstallDir&"html/"&rss("id")&".html</link>" & vbcrlf
Else
str = str & "<link>http://"&web&"/"&strInstallDir&"show.asp?id="&rss("id")&"</link>" & vbcrlf
end if
str = str & "<title>"&rss("biaoti")&"</title>" & vbcrlf
str = str & "<creator>"&rss("name")&"</creator>" & vbcrlf
str = str & "<PubDate>"&rss("fbsj")&"</PubDate>" & vbcrlf

If rss("Readinfo")=1 Then
str = str & "<description><![CDATA[会员权限阅读]]></description>" & vbcrlf
elseIf rss("Readinfo")=2 Then
str = str & "<description><![CDATA[VIP会员权限阅读]]></description>" & vbcrlf
Else
str = str & "<description><![CDATA["&replace(replace(left(rss("memo"),200),"<","&lt;"),">","&gt;")&"]]></description>" & vbcrlf
End If

str = str & "</item>" & vbcrlf
rss.MoveNext          
Loop
rss.close
set rss=Nothing

str = str & "</channel>" & vbcrlf
str = str & "</rss>" & vbcrlf
set fso = nothing

Set objStream = Server.CreateObject("ADODB.Stream")
With objStream
.Open
.Charset = "UTF-8"
.Position = objStream.Size
.WriteText=str
.SaveToFile server.mappath("/"&strInstallDir&"xml/rss"&city_oneid&"_"&city_twoid&"_"&city_threeid&".xml"),2 '生成的XML文件名,不建议修改,如果你有更好的方案,可以改
.Close
End With
'===============重新生成相关的rss结束=====================================

Next
Call CloseRs(rs)
Call CloseDB(conn)
If Htmlhome=1 Then Call HomeHtml()'重新生成首页
response.Write "<script language=javascript>location.replace('infolist.asp?page="&request("t1")&"&t1="&trim(request("t1"))&"&t2="&trim(request("t2"))&"&t3="&trim(request("t3"))&"&dick="&request("dick")&"&city_one="&request("city_one")&"&city_two="&request("city_two")&"&city_three="&request("city_three")&"');</script>"

%>