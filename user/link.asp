<!--#include file="../conn.asp"-->
<!--#include file="../SqlIn.Asp"-->
<!--#include file="../Include/Function.asp"-->
<!--#include file="../Include/label.asp"-->

<%'友情链接调用说明
'调用代码<SCRIPT src="user/link.asp?tp=0&LineNum=10&LineLogo=1&MaxLine=0&hg=31&btw=6"></SCRIPT>
'参数说明：tp为1时，仅为LOGO，为2时，仅文字，为0时图文。LineNum每行显示个数，LineLogo图片显示行数，hg图片高度,标题字符数
Call OpenConn
dim ttcc,c,c1,c2,c3,TempText,tp,hg,btw,w,j,k,wzname,logo,LineNum,LineLogo,MaxLine,bb,i
c=Request("c")
c=split(c,",")
If c(0)="" Then c(0)=0
c1=strint(c(0))
If c(1)="" Then c(1)=0
c2=strint(c(1))
If c(2)="" Then c(2)=0
c3=strint(c(2))
If c1="" Then c1=0
If c2="" Then c2=0
If c3="" Then c3=0

tp=Strint(Request("tp"))
LineNum=Strint(Request("LineNum"))
LineLogo=Strint(Request("LineLogo"))
MaxLine=Strint(Request("MaxLine"))
hg=Strint(Request("hg"))
btw=Strint(Request("btw"))

IF c1=0 then
ttcc=""
elseIF c3>0 then
ttcc="and city_oneid="&c1&" and city_twoid="&c2&" and city_threeid="&c3
elseIF c2>0 then
ttcc="and city_oneid="&c1&" and city_twoid="&c2
Else
ttcc="and city_oneid="&c1
End If

Dim ServerURL,m
ServerURL=CStr(Request.ServerVariables("SCRIPT_NAME"))
m=InStrRev(ServerURL,"/") '从右边第一个字符起查找"_"的位置,n为返回值
ServerURL=left(ServerURL, m-1)'显示从左边数第"m-1"个字符前面的字符,
m=InStrRev(ServerURL,"/") '从右边第一个字符起查找"_"的位置,n为返回值
ServerURL=left(ServerURL, m)'显示从左边数第"m"个字符前面的字符,
ServerURL="http://"&Request.ServerVariables("SERVER_NAME")&""&ServerURL


TempText="<table border=""1"" cellspacing=""0"" width=""980"" bgcolor=""#CDE4CF"" bordercolorlight=""#008000"" bordercolordark=""#FFFFFF"" cellpadding=""1"" >" 
'图片参数：tp为1时，仅为LOGO，为2时，仅文字，为0时图文
IF tp=1 Then 
w=" and logo<>'' and logo<>""0"""
hg=ph
Else
If tp=2 Then
w=" and (logo=""0"")"
else
w=""
End iF
End If

set rs=server.createobject("adodb.recordset")
set rs=conn.execute("select  url,logo,wzname from DNJY_link where known=1 "&w&" "&ttcc&" order by paixu "&DN_OrderType&",id")
IF Rs.eof Then
Rs.close:Set Rs=Nothing
set rs=conn.execute("select  url,logo,wzname from DNJY_link where  known=1 "&w&" and city_oneid=0 and city_twoid=0 and city_threeid=0 order by paixu "&DN_OrderType&",id")
End If
if rs.eof then
TempText=TempText&"<p><br>暂无"
IF c1>0 Then
bb= conn.Execute("Select city From DNJY_city Where id="&c1&" and twoid="&c2&" and threeid="&c3&"")(0)
End IF
TempText=TempText&""&bb&"链接!"
else
  i=0
  m=0
if MaxLine=0 Then MaxLine=10000
do while not rs.eof and m<MaxLine
TempText=TempText&"<tr height="&hg&">"
for j=1 to LineNum
if rs.eof then
TempText=TempText&"<td align=""center"" width=""980""> <center><a href=""addlink.asp?adminlink=2"" target=""_blank"">"
if k=<LineLogo*LineNum then
TempText=TempText&"<img src=""../img/your_link.gif"" width=88 height=31 border=0 alt=""申请您的座位"">"
else
TempText=TempText&"您的位置"
end if
TempText=TempText&"</a></td>"
else
TempText=TempText&"<td width=""980"" align=""center"">" 
if i<LineLogo*LineNum  then
TempText=TempText&"<center><a href="""&rs("URL")&""" target=_blank>"
if rs("logo")="0"  then 
TempText=TempText&"<table width=88 height=31  background=""../Img/nologo2.gif""><tr><td  align=""center""><a href="""&rs("URL")&""" target=_blank>"
if len(rs("wzname"))>6 then
TempText=TempText&""&Left(rs("wzname"),6)&""
else
TempText=TempText&""&rs("wzname")&""
end if
TempText=TempText&"</a></td></tr></table>"
else
if j = MaxLine * LineNum then
TempText=TempText&"<a href=""link_gd.asp"" target=""_blank""><img src=""../"&ServerURL&"img/gdlj.gif"" width=88 height=31 border=0></a>"
else
TempText=TempText&"<img src="""&rs("logo")&""" width=88 height=31 border=0 alt="&rs("wzname")&">"
end if
end if
TempText=TempText&"</a>"
else
TempText=TempText&"<table><tr><td class=tdc1 align=""center"">"
if j = MaxLine * LineNum then
TempText=TempText&"<a href=""link_gd.asp"" target=""_blank"" >更多链接>></a>"
else
TempText=TempText&"<a href="&rs("URL")&" title="&rs("wzname")&"--"&rs("wzname")&" target=_blank>"
if len(rs("wzname"))>6 Then
TempText=TempText&""&Left(rs("wzname"),6)&""
Else
TempText=TempText&""&rs("wzname")&""
end If
TempText=TempText&"</a>"
end If
TempText=TempText&"</td></tr></table>" 
end if 
 i=i+1
 k=k+1
rs.movenext
end if
Next
TempText=TempText&"</tr>"
m=m+1 
Loop
end if
Call CloseRs(rs)
Response.Write "document.write('"&TempText&"</table>');"
Call CloseDB(conn)%>
