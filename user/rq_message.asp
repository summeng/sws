<!--#include file="../conn.asp"-->
<!--#include file="../SqlIn.Asp"-->
<!--#include file="../config.asp"-->
<!--#include file="../Include/Function.asp"-->
<!--#include file="../Include/label.asp"-->
<%dim id,ServerURL,n
id=strint(request.querystring("id"))

ServerURL=CStr(Request.ServerVariables("SCRIPT_NAME"))
n=InStrRev(ServerURL,"/") '从右边第一个字符起查找"_"的位置,n为返回值
ServerURL=left(ServerURL, n)'显示从左边数第"n"个字符前面的字符,
ServerURL="http://"&Request.ServerVariables("SERVER_NAME")&""&ServerURL

response.write "document.write('<table width=""100%"" border=""0"" cellspacing=""0"" cellpadding=""0""><tr><td>');" & vbCrLf
Call OpenConn            
              dim k
              k=1
			  Set rs= Server.CreateObject("ADODB.Recordset")
              sql = "select  * from DNJY_reply where xxid="&id&"  and hfyz=1 order by id "&DN_OrderType&""
              rs.open sql,conn,1,1
			  if rs.eof then
response.write "document.write('');" & vbCrLf
  else
         do while not rs.eof
response.write "document.write('<table width=""100%"" border=""0"" cellspacing=""0"" cellpadding=""0""><tr>');" & vbCrLf
response.write "document.write('<td width=""5%"" bgcolor=""#FAFAFA""><div align=""center""><b>"&k&"</b></div></td>');" & vbCrLf
response.write "document.write('<td width=""95%"" bgcolor=""#FAFAFA""><b>该信息由<font color=""#CC0000"">');" & vbCrLf
if rs("username")<>"" Then
response.write "document.write('"&trim(rs("username"))&"');" & vbCrLf
Else
response.write "document.write('匿名用户');" & vbCrLf
end If
response.write "document.write('</font></b>在<font color=""#CC0000"">"&trim(rs("hfsj"))&"</font>评论</b></td>');" & vbCrLf
response.write "document.write('</tr><tr><td height=""1"" colspan=""2"" bgcolor=""#CCCCCC""></td></tr><tr><td width=""5%"">&nbsp;</td><td width=""95%""><span style=""word-break:break-all"">"&HTMLEncode(rs("neirong"))&"</span></td></tr></table>');" & vbCrLf
              rs.movenext
              k=k+1
			  If k>5 Then
			  response.write "document.write('<A href=""../review.asp?ID="&id&""" target=_blank>查看全部评论>>></a>');" & vbCrLf
			  exit do
			  End if
              loop
              end If
         
response.write "document.write('</td></tr></table>');" & vbCrLf
rs.close:set rs=Nothing
Call CloseDB(conn)%>
