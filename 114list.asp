<!--#include file="conn.asp"-->
<!--#include file="SqlIn.Asp"-->
<!--#include file=top.asp-->

<%
'=====================================
'天天供求信息网站管理系统
'天天网络科技工作室出品
'客服中心http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>

<link href="/<%=strInstallDir%>css/css.css" rel="stylesheet" type="text/css">
<style type="text/css">
<!--
.style1 {color: #FF0000}
-->
</style>
</head>

<body>
<table width="980" border="0" align="center" cellpadding="0" cellspacing="0">

  <tr>
    <td width="180" height="100%" valign="top">
<table width="180" border="0" align="left" cellPadding="0" cellSpacing="0" bordercolor="#111111" class="font_10_e_blue" style="border-collapse: collapse">

         <tr>
           <td vAlign="top"><div align="center"> 
             
  <!---分页左侧下广告开始---->  
  <table align="center" bgcolor="#ffffff">
  <tr>  
   <td class="ty1"><script language=javascript src="<%=adjs_path("ads/js","mtl",c1&"_"&c2&"_"&c3)%>"></script></td>
  </tr>
  </table>  
  </div></td>
         </tr>         
       </table>
	   </td>
 
<td width="800" valign="top" bgcolor="#FFFFFF">
	<table width="800" border="0" align="right" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF"> 
	<!---上部通栏广告开始----> 
  <table align="center" bgcolor="#ffffff">
  <tr>  
   <td class="ty1"><script language=javascript src="<%=adjs_path("ads/js","fygg",c1&"_"&c2&"_"&c3)%>"></script></td>
  </tr>
  </table>
  <!---上部通栏广告结束----> 

<table width="800" border="0" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF" class="ty1"> 
<tr><TD><div>
<%Dim rs114,bm
Call OpenConn
set rs114=server.CreateObject("adodb.recordset")
rs114.Open "select * from DNJY_114 where dqsj>"&DN_Now&""&ttcc114&" order by hfje/xsts "&DN_OrderType&"",conn,1,1
bm=1
if rs114.eof Then
Response.Write "<p><br>暂无"
IF c1>0 Then Response.Write conn.Execute("Select city From DNJY_city Where id="&c1&" and twoid="&c2&" and threeid="&c3&"")(0)
Response.Write "都市114服务!"
Else
do while not rs114.eof%>
 <div class="articleline114">        
 <%=rs114("d_name")%>:<%=rs114("d_tel")%>              
</div>
<%if (bm mod 2)=0 then%>
<br/>
<%bm=bm+1
end if
rs114.movenext
Loop
end if
Call CloseRs(rs114)
Call CloseDB(conn)%></div></TD></tr>
</table>
</td>
  <table width="980" cellpadding="0" cellspacing="0" align="center">
     <tr>
       <td><!--#include file=footer.asp--></td>
     </tr>
</table>
</body>
</html>

