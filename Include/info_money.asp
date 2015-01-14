<!--#include file="../conn.asp"-->
<!--#include file="../config.asp"-->
<!--#include file="head.asp"-->
<!--#include file="Function.asp"-->
<!--#include file="label.asp"-->
<%dbpath="../"
dim TempText,xxsx,gdxx,s,l,zs,zt,hg,d,dbpath,Fso,gd,dj
Dim ThisPage,Pagesize,Allrecord,Allpage,iii,xt,b,tj,bb,a,biaoti,sj2,yss,b2
xxsx=Strint(Request("xxsx"))
gdxx=Strint(Request("gdxx"))
id=Strint(Request("id"))
n=Strint(Request("n"))
s=Strint(Request("s"))
l=Strint(Request("l"))
zs=Strint(Request("zs"))
zt=Strint(Request("zt"))
hg=Strint(Request("hg"))
d=Strint(Request("d"))
gd=Strint(Request("gd"))
dj=Strint(Request("dj"))
IF s<=0 Then s=1
IF l<=0 Then l=1
IF zs<=0 Then zs=1
IF hg<=0 Then hg=22

Select Case xxsx
	Case 1
	b2="最新"
	Case 2
	b2="竞价"
	Case 3	
	b2="推荐"
	Case 4	
	b2="热门"
	Case 5	
	b2="回复"
	Case Else	
	b2=""
	End Select

Dim ServerURL,m
ServerURL=CStr(Request.ServerVariables("SCRIPT_NAME"))
m=InStrRev(ServerURL,"/") '从右边第一个字符起查找"_"的位置,n为返回值
ServerURL=left(ServerURL, m-1)'显示从左边数第"m-1"个字符前面的字符,
m=InStrRev(ServerURL,"/") '从右边第一个字符起查找"_"的位置,n为返回值
ServerURL=left(ServerURL, m)'显示从左边数第"m"个字符前面的字符,
ServerURL="http://"&Request.ServerVariables("SERVER_NAME")&""&ServerURL

Call OpenConn
set rs=server.createobject("ADODB.recordset")
sql = "select city_oneid,city_twoid,city_threeid from DNJY_info where id="&id&""
Rs.open Sql,conn,1,1
IF Not Rs.eof Then
city_oneid=rs("city_oneid")
city_twoid=rs("city_twoid")
city_threeid=rs("city_threeid")
Call CloseRs(rs)
End if
set rs=server.createobject("ADODB.recordset")
sql = "select * from DNJY_info where yz=1"
if jiaoyi=0 Then sql=sql&" and zt="&jiaoyi
if overdue=0 Then sql=sql&" and dqsj>"&DN_Now&""
sql=sql&" and city_oneid="&city_oneid&" and city_twoid="&city_twoid&" and city_threeid="&city_threeid

if xxsx=2 Then sql=sql&" and hfxx=1"
if xxsx=3 Then sql=sql&" and tj>0"

sql=sql&" order by "
If gd=1 Then sql=sql&"b "&DN_OrderType&","
if xxsx=2 Then
sql=sql&"hfje/fbts "&DN_OrderType&"," 
elseif xxsx=4 Then
sql=sql&"llcs "&DN_OrderType&","
end If
sql=sql&"fbsj "&DN_OrderType&",id "&DN_OrderType&""

Rs.open Sql,conn,1,1
IF Rs.eof Then
TempText="暂无"
IF city_oneid>0 Then TempText=TempText&conn.Execute("Select city From DNJY_city Where id="&city_oneid&" and twoid="&city_twoid&" and threeid="&city_threeid&"")(0)
TempText=TempText&b2&"信息"
Else
TempText=""
TempText="<table width=""100%""  border=""0"" cellspacing=""4"" cellpadding=""0"" >"
iii=0
i=0
Do While not Rs.eof
i=i+1
IF iii Mod l=0 then TempText=TempText&  "<tr height="""&hg&""">"
TempText=TempText&  "<td>"
IF n>0 Then
TempText=TempText&  "["
IF n=1 Then
TempText=TempText&  Rs("type_one")
Else
IF Rs("type_two")<>"" then
TempText=TempText&  Rs("type_two")
ElseIF Rs("type_two")<>"" then
TempText=TempText&  Rs("type_two")
Else
TempText=TempText&  Rs("type_one")
End IF
End IF
TempText=TempText&  "]"
End IF
TempText=TempText&  " <a title="""&Rs("biaoti")&"||"&rs("name")&"发布于"&FormatDateTime(rs("fbsj"),vbShortDate)&""" href="""&x_path("../",rs("id"),ServerURL,rs("Readinfo"))&""" target=""_blank""><SPAN style=""color:"&rs("a")&";font-size="&zt&"px"">"&CutStr(Rs("biaoti"),zs,"...")&"</SPAN></a> "
IF dj=1 Then TempText=TempText&  " <font color=""#008080"" >["&rs("llcs")&"]</font>" 
IF d>0 Then TempText=TempText&  " <font color=""#008080"" >"&FormatDate(rs("fbsj"),d)&"</font>" 
TempText=TempText&  "</TD>"
iii=iii+1
IF iii Mod l=0 then TempText=TempText&  "</tr>"
if i>=s then exit do
Rs.movenext
Loop
Set Fso=Nothing
IF iii Mod l<>0 then
Do While not iii Mod l=0
TempText=TempText&  "<td></td>"
iii=iii+1
Loop
TempText=TempText&  "</tr>"
End IF
TempText=TempText&  "</table>"
If gdxx=1 Then TempText=TempText&  "<table align=""right""><tr><td><A href=""../list.asp?"&CT_ID&"&xxsx="&xxsx&"""><SPAN style=""font-size="&zt&"px"">更多>>></SPAN></a></td></tr></table>"
End If
Call CloseRs(rs)
Call CloseDB(conn)
response.write "document.write('"&TempText&"');"%>



