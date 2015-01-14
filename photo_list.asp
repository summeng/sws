<!--#include file="conn.asp"-->
<!--#include file="SqlIn.Asp"-->
<!--#include file="Include/ContentAutoPage.asp"-->
<!--#include file="top.asp"-->
 <!--#include file="Include/err.asp"-->
<%
'=====================================
'天天供求信息网站管理系统
'天天网络科技工作室出品
'客服中心http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>
<script type="text/javascript" src="Include/onews.js"></script>
<script type="text/javascript" src="admin/inc/js_left.js"></script>
<link rel="stylesheet" href="css/news.css" type="text/css" />
</head>
<body>
<table width="980" border="0" align="center" cellpadding="0" cellspacing="0">
<TR>
<TD width="760" valign="top">
<table width="760" border="0" cellpadding="0" cellspacing="0" class="ty1">
<TR><TD width="760"><%=f_photo(c1,c2,c3,0,0,5,25,140,120,1,20,1,request("page"))%></TD></TR>
</TABLE>
<TABLE width="760">
<TR>
	<TD width="760"><script src="<%=adjs_path("ads/js","fygg",c1&"_"&c2&"_"&c3)%>"></script></TD>
</TR>
</TABLE>
</TD>

<TD width="3"></TD>

<TD width="215" valign="top">
  <table width="215" border="0" cellpadding="0" cellspacing="0" class="ty1">
  <tr>
    <td align="center" height="30" class="ty3"><b>资讯栏目</b></td>
  </tr>
  <tr>
    <td colspan="2" align="center">
<%Call OpenConn
Dim sqlt,rst,rc,rsSmallClass,bigcid,SmallCID
bigcid=Request("bigcid")
SmallCID=Request("SmallCID")
sqlt="select * from DNJY_bigClass order by BigClassID"
set rst=server.createobject("adodb.recordset")
rst.open sqlt,conn,1,1
if rst.eof then
response.write "<CENTER>不存在的大类别!</CENTER>"
rst.close
set rst=nothing
Call CloseDB(conn)
response.End
Else
i=0
rc=""
do while not rst.eof
If rst("ID")=bigcid Then rc="ff0000"
Response.Write "<div id=""body"">"
Response.Write "<div class=""a"" id=""AA"&i&""" onclick=""TAB("&i&")"" ><img src=""img/zj.gif"" width=""15"" height=""15""><a  href=""news.asp?classid="&rst("id")&"&bigcid="&rst("id")&"&"&L_ID&"""><strong><span style=""font-size:13px; font-weight:bold;color:"&rc&""";"" >"&rst("BigClassName")&"</strong></span></a></div>"
rc=""
If i=0 Then
Response.Write"<div class=""cc"" id=""BB"&i&""" style=""display:block;"">"
Else
Response.Write"<div class=""cc"" id=""BB"&i&""" style=""display:none;"">"
End If
set rsSmallClass=server.CreateObject("adodb.recordset")
rsSmallClass.open "Select * from DNJY_SmallClass Where bigid="&rst("id")&" order by sortingid",conn,1,1
if rsSmallClass.bof and rsSmallClass.eof Then
response.write "<p>&nbsp;&nbsp;&nbsp;无下级分类</p>"
else
do while not rsSmallClass.eof
If rsSmallClass("SmallClassID")=SmallCID Then rc="ff0000"
Response.Write "<p>&nbsp;&nbsp;<img src=""img/js.gif"" width=""15"" height=""15""><a  href=""news.asp?classid="&rst("id")&"&SmallCID="&rsSmallClass("SmallClassID")&"&bigcid="&rst("id")&"&i="&i&"&"&L_ID&""">"
Response.Write "<span style=""font-size:12px;color:"&rc&""";>"&rsSmallClass("SmallClassName")&"</span></a></p>"
rc=""
	rsSmallClass.movenext
	loop
	end If    
	i=i+1
Response.Write "</div>"
rst.movenext
loop	  
rsSmallClass.close
set rsSmallClass=Nothing
rst.close
set rst=Nothing
End if%>
</td>
  </tr>     
</table>
      <table width="215" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td  height="7"></td>
        </tr>
      </table>
  <table width="215" border="0" cellpadding="0" cellspacing="0" class="ty1">
    <tr>
    <td align="center" height="30" class="ty3"><b>最新推荐 new10</b></td>
  </tr>

  <tr>
    <td colspan="2" align="center"><%=f_news(c1,c2,c3,bigcid,1,0,0,0,0,0,25,10,1,13,10,0)%></td>
  </tr>     
</table>
      <table width="215" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td  height="7"></td>
        </tr>
      </table>
<table width="215" border="0" cellpadding="0" cellspacing="0"  class="ty1">
  <tr>
    <td align="center" height="30" class="ty3">热门点击   top10</td>
  </tr>
  <tr>
    <td height="22" align="center"><%=f_news(c1,c2,c3,bigcid,3,0,0,0,0,0,25,10,1,13,10,0)%></td>
  </tr>
</table>
</TD>

</TR>
</TABLE>

<TABLE cellSpacing=0 cellPadding=0 width=980 align=center border=0>
  <TBODY>
  <TR>
    <TD width="769">
    <p align="center"><!--#include file=footer.asp--></TD></TR>                         
  <TR>
<TD width="769"></TD></TR></TBODY></TABLE>
<p>
</body>
</html>