<!--#include file="../conn.asp"-->
<!--#include file="../config.asp"-->
<!--#include file="../Include/head.asp"-->
<!--#include file="../Include/Function.asp"-->
<!--#include file="../SqlIn.Asp"-->
<link href="/<%=strInstallDir%>css/style.css" rel="stylesheet" type="text/css">
<body topmargin="0" leftmargin="0" style="text-align:center;">
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
m=InStrRev(ServerURL,"/") '���ұߵ�һ���ַ������"_"��λ��,nΪ����ֵ
ServerURL=left(ServerURL, m-1)'��ʾ���������"m-1"���ַ�ǰ����ַ�,
m=InStrRev(ServerURL,"/") '���ұߵ�һ���ַ������"_"��λ��,nΪ����ֵ
ServerURL=left(ServerURL, m)'��ʾ���������"m"���ַ�ǰ����ַ�,
ServerURL="http://"&Request.ServerVariables("SERVER_NAME")&""&ServerURL


Response.Write "ȫ������&nbsp;&nbsp;&nbsp;&nbsp;<a href=""link_edit.asp?"&C_ID&""" target=""_blank"" >�޸�����</a><table border=""1"" cellspacing=""0"" width=""980"" bgcolor=""#CDE4CF"" bordercolorlight=""#008000"" bordercolordark=""#FFFFFF"" cellpadding=""1"" align=""center"">" 
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
Response.Write "<p><br>����"
IF c1>0 Then
bb= conn.Execute("Select city From DNJY_city Where id="&c1&" and twoid="&c2&" and threeid="&c3&"")(0)
End IF
Response.Write ""&bb&"����!"
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
Response.Write "<img src=""../img/your_link.gif"" width=88 height=31 border=0 alt=""����������λ"">"
else
Response.Write "����λ��"
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
Response.Write "</a>"
else
Response.Write "<table><tr><td align=""center"">"
if j = MaxLine * LineNum then
Response.Write "<a href=""link_gd.asp"" target=""_blank"" >��������>></a>"
else
Response.Write "<a href="&rs("URL")&" title="&rs("wzname")&"--"&rs("wzname")&" target=_blank>"
if len(rs("wzname"))>btw Then
Response.Write ""&Left(rs("wzname"),btw)&""
Else
Response.Write ""&rs("wzname")&""
end If
Response.Write "</a>"
end If
Response.Write "</td></tr></table>" 
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
<table border="1" cellspacing="0" width="980"  bgcolor="#CDE4CF" bordercolorlight="#008000" bordercolordark="#FFFFFF" cellpadding="1" align="center"><tr><td align="center" width="980"> <center><script src="../ads/JS/syzb730.js"></script></td></tr></table>