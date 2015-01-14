<!--#include file="../conn.asp"-->
<!--#include file="../SqlIn.Asp"-->
<!--#include file="../config.asp"-->
<!--#include file="../Include/Function.asp"-->

<%dim id,cs
id=strint(request.querystring("id"))
Call OpenConn            
cs=conn.Execute("Select hfcs From DNJY_info Where id="&id&"")(0)
response.write "document.write('此信息共有"&cs&"条评论，未审核的不可见。');" & vbCrLf
Call CloseDB(conn)%>
