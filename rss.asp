<!--#include file="conn.asp"-->
<!--#include file="SqlIn.Asp"-->
<!--#include file="config.ASP"-->
<!--#include file="Include/head.asp"-->
<!--#include file="Include/Function.asp"-->
<!--#include file="Include/label.asp"-->

<%
'=====================================
'天天供求信息网站管理系统
'天天网络科技工作室出品
'客服中心http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>
<%Dim strURL,dqname,rsdq 
strURL = "http://" & request.servervariables("server_name") & _ 
left(request.servervariables("script_name"),len(request.servervariables("SCRIPT_NAME"))-len("/rss.asp")) '中的/rss.asp为你的该文件名 
If request("c")<>"" then
c=request("c")
c=split(c,",")
If c(0)="" Then c(0)=0
c1=strint(c(0))
If c(1)="" Then c(1)=0
c2=strint(c(1))
If c(2)="" Then c(2)=0
c3=strint(c(2))
End If

If request("t")<>"" then
t=request("t")
t=split(t,",")
If t(0)="" Then t(0)=0
t1=strint(t(0))
If t(1)="" Then t(1)=0
t2=strint(t(1))
If t(2)="" Then t(2)=0
t3=strint(t(2))
End If

If c1="" Then c1=0
If c2="" Then c2=0
If c3="" Then c3=0
If t1="" Then t1=0
If t2="" Then t2=0
If t3="" Then t3=0

dim a,b,d
Dim cc,tt
	IF c1=0 then
	cc=""
	elseIF c3>0 then
	cc=" and city_oneid="&c1&" and city_twoid="&c2&" and city_threeid="&c3
	elseIF c2>0 then
	cc=" and city_oneid="&c1&" and city_twoid="&c2
	Else
	cc=" and city_oneid="&c1
	End IF
	IF t1=0 then
	tt=""	
	elseIF t3>0 then
	tt=" and type_oneid="&t1&" and type_twoid="&t2&" and type_threeid="&t3
	elseIF t2>0 then
	tt=" and type_oneid="&t1&" and type_twoid="&t2
	Else
	tt=" and type_oneid="&t1
	End If
Call OpenConn
set rs=server.createobject("ADODB.recordset")
sql = "select top 20  * from DNJY_info where yz=1"
sql=sql&""&cc&tt&""
sql=sql&" order by fbsj "&DN_OrderType&""
rs.open sql,conn,1,1
if rs.Eof then
response.write "<item>本分类暂无信息</item>"
Call CloseRs(rs)
Call CloseDB(conn)
response.end
end If
set rsdq=server.createobject("ADODB.recordset")
sql = "select name from DNJY_type where id="&t1&" and twoid="&t2&" and threeid="&t3
rsdq.open sql,conn,1,1
if Not rsdq.Eof then
dqname=rsdq("name")
End if
response.contenttype="text/xml" 
response.write "<?xml version=""1.0"" encoding=""gb2312"" ?>" & vbcrlf 
response.write "<rss version=""2.0"">" & vbcrlf 
response.write "<channel>" & vbcrlf 
response.write "<title>" & webname&"_"&dqname& " RSS feed</title>" & vbcrlf 
response.write "<link>" & strURL & "</link>" & vbcrlf 
response.write "<language>zh-cn</language>" & vbcrlf 
response.write "<copyright>An RSS feed for " & web & "</copyright>" & vbcrlf 
while not rs.eof 
response.write "<item>" & vbcrlf 
response.write "<title><![CDATA["&rs("biaoti")&"]]></title>" & vbcrlf 
response.write "<link>"&strURL&"/"&strInstallDir&""&x_path("",rs("id"),"",rs("Readinfo"))&"</link>" & vbcrlf
response.write "<description><![CDATA[" & rs("biaoti") & "<br />" & rs("name") & "<br /><br />]]></description>" & vbcrlf 
response.write "<pubDate>" & return_RFC822_Date(rs("fbsj"),"GMT") & "</pubDate>" & vbcrlf 
response.write "</item>" & vbcrlf 
rs.movenext 
wend 
response.write "</channel>" & vbcrlf 
response.write "</rss>" & vbcrlf 
Call CloseRs(rs)
Call CloseDB(conn)

Function return_RFC822_Date(byVal myDate, byVal TimeZone) 
Dim myDay, myDays, myMonth, myYear 
Dim myHours, myMinutes, mySeconds 

myDate = CDate(myDate) 
myDay = EnWeekDayName(myDate) 
myDays = Right("00" & Day(myDate),2) 
myMonth = EnMonthName(myDate) 
myYear = Year(myDate) 
myHours = Right("00" & Hour(myDate),2) 
myMinutes = Right("00" & Minute(myDate),2) 
mySeconds = Right("00" & Second(myDate),2) 


return_RFC822_Date = myDay&","& _ 
myDays&" "& _ 
myMonth&" "& _ 
myYear&" "& _ 
myHours&":"& _ 
myMinutes&":"& _ 
mySeconds&" "& _ 
" " & TimeZone 
End Function 

Function EnWeekDayName(InputDate) 
Dim Result 
Select Case WeekDay(InputDate,1) 
Case 1:Result="Sun" 
Case 2:Result="Mon" 
Case 3:Result="Tue" 
Case 4:Result="Wed" 
Case 5:Result="Thu" 
Case 6:Result="Fri" 
Case 7:Result="Sat" 
End Select 
EnWeekDayName = Result 
End Function 

Function EnMonthName(InputDate) 
Dim Result 
Select Case Month(InputDate) 
Case 1:Result="Jan" 
Case 2:Result="Feb" 
Case 3:Result="Mar" 
Case 4:Result="Apr" 
Case 5:Result="May" 
Case 6:Result="Jun" 
Case 7:Result="Jul" 
Case 8:Result="Aug" 
Case 9:Result="Sep" 
Case 10:Result="Oct" 
Case 11:Result="Nov" 
Case 12:Result="Dec" 
End Select 
EnMonthName = Result 
End Function 
%>