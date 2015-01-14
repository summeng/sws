<!--#include file="../conn.asp"-->
<!--#include file="../config.asp"-->
<!--#include file="../Include/head.asp"-->
<!--#include file="../Include/Function.asp"-->
<!--#include file="../Include/label.asp"-->
<!--#include file="../SqlIn.Asp"-->
<title><%=webname%>_修改友情链接</title>
<link href="/<%=strInstallDir%>css/style.css" rel="stylesheet" type="text/css">
<%Call OpenConn
dim tp,hg,btw,w,j,k,LineNum,LineLogo,MaxLine,bb
tp=0
LineNum=8
LineLogo=0
MaxLine=0
hg=31
btw=6

Dim ServerURL,m
ServerURL=CStr(Request.ServerVariables("SCRIPT_NAME"))
m=InStrRev(ServerURL,"/") '从右边第一个字符起查找"_"的位置,n为返回值
ServerURL=left(ServerURL, m-1)'显示从左边数第"m-1"个字符前面的字符,
m=InStrRev(ServerURL,"/") '从右边第一个字符起查找"_"的位置,n为返回值
ServerURL=left(ServerURL, m)'显示从左边数第"m"个字符前面的字符,
ServerURL="http://"&Request.ServerVariables("SERVER_NAME")&""&ServerURL


Response.Write "<CENTER>修改友情链接<table border=""1"" cellspacing=""0"" width=""980"" bgcolor=""#CDE4CF"" bordercolorlight=""#008000"" bordercolordark=""#FFFFFF"" cellpadding=""1"" >" 
IF tp=1 then
w=" and logo<>'' and logo<>""0"""
hg=ph
Else
If tp=2 Then
w=" and (logo=""0"")"
else
w=""
End iF
End iF
set rs=server.createobject("adodb.recordset")
set rs=conn.execute("select  id,url,logo,wzname from DNJY_link where known=1 "&w&" "&ttcclink&" order by linktop "&DN_OrderType&",paixu "&DN_OrderType&",id")
IF Rs.eof Then
Rs.close:Set Rs=Nothing
set rs=conn.execute("select  id,url,logo,wzname from DNJY_link where  known=1 "&w&" "&ttcclink&" order by linktop "&DN_OrderType&",paixu "&DN_OrderType&",id")
End If
if rs.eof then
Response.Write "<p><br>暂无"
IF c1>0 Then
bb= conn.Execute("Select city From DNJY_city Where id="&c1&" and twoid="&c2&" and threeid="&c3&"")(0)
End IF
Response.Write ""&bb&"链接!"
Else
k=0
  i=0
  m=0
if MaxLine=0 Then MaxLine=10000
do while not rs.eof and m<MaxLine
Response.Write "<tr height="&hg&">"
for j=1 to LineNum
if rs.eof then
Response.Write "<td align=""center"" width=""980""> <center><a href=""addlink.asp?adminlink=2"" target=""_blank"">"
if k=<LineLogo*LineNum then
Response.Write "<img src=""../img/your_link.gif"" width=88 height=31 border=0 alt=""申请您的座位"">"
else
Response.Write "您的位置"
end if
Response.Write "</a></td>"
else
Response.Write "<td width=""980"" align=""center"">" 
if i<LineLogo*LineNum  then
Response.Write "<center><a href="""&rs("URL")&""" target=_blank>"
if rs("logo")="0"  then 
Response.Write "<table background=""../Img/nologo2.gif"" width=88 height=31 border=0><tr><td align=""center"" valign=""top""><a href="""&rs("URL")&""" target=_blank>"
if len(rs("wzname"))>btw then
Response.Write ""&Left(rs("wzname"),btw)&""
else
Response.Write ""&rs("wzname")&""
end if
Response.Write "</a></td></tr></table>"
else
if j = MaxLine * LineNum then
Response.Write "<a href=""link_gd.asp"" target=""_blank""><img src=""../"&ServerURL&"img/gdlj.gif"" width=88 height=31 border=0></a>"
else
Response.Write "<img src="""&rs("logo")&""" width=88 height=31 border=0 alt="&rs("wzname")&">"
end if
end if
Response.Write "<a href=""#"" onClick=""window.open('addlink.asp?adminlink=1&id="&rs("id")&"&"&C_ID&"&name="&rs("name")&"','Sample','toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=no,resizable=no,copyhistory=no,width=620,height=600,left=300,top=100')""><span style=""color:#999999"">修改</span></a>"
else
Response.Write "<table><tr><td align=""center"">"
if j = MaxLine * LineNum then
Response.Write "<a href=""link_gd.asp"" target=""_blank"" >更多链接>></a>"
else
Response.Write "<a href="&rs("URL")&" title="&rs("wzname")&"--"&rs("wzname")&" target=_blank>"
if len(rs("wzname"))>btw Then
Response.Write ""&Left(rs("wzname"),btw)&""
Else
Response.Write ""&rs("wzname")&""
end If
Response.Write "</a>"
end If
Response.Write " <a href=""#"" onClick=""window.open('addlink.asp?adminlink=1&id="&rs("id")&"&"&C_ID&"','Sample','toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=no,resizable=no,copyhistory=no,width=620,height=680,left=300,top=100')""><span style=""color:#999999"">修改</span></a></td></tr></table>" 
end if 
 i=i+1
 k=k+1
rs.movenext
end if
Next
Response.Write "</tr>"
m=m+1 
Loop
end if
Call CloseRs(rs)
Response.Write "</table>"
Call CloseDB(conn)%>
