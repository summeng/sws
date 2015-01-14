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
<%Dim s1,s2
n=request("n")
If n="" Then n="0,0,0"
n=split(n,",")
If n(0)="" Then n(0)=0
n1=strint(n(0))
If ubound(n)<1 Then
n2=0
else
n2=strint(n(1))
End if
If ubound(n)<2 Then
n3=0
else
n3=strint(t(2))
End if
s1=trim(request.form("s1"))
s2=trim(request.form("s2"))
if s1="" then
s1=request.QueryString("s1")
end if
if s2="" then
s2=request.QueryString("s2")
end if
%>
<table width="980" border="0" align="center" cellpadding="0" cellspacing="0">
<TR>
<TD width="760" valign="top">
<table width="760" border="0" cellpadding="0" cellspacing="0" class="ty1">
<TR><TD width="760" ><%=news_photo(c1,c2,c3,n1,n2,n3,1,1,1,1,20,150,140,5,20,1,request("page"),s1,s2)%></TD></TR>
                     
</TABLE>
<table width="100%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="#6699cc">
  <tr> 
    <td align="center">搜索查询</td>
  </tr>
  <tr bgcolor="#FFFFFF"> 
    <form name="form2" method="post" action="news_list.asp">
      <td>
	  <table width="100%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="#CCCCCC">
          <tr bgcolor="#FFFFFF" align="center"> 
            <td > <input name="s2" type="text" id="s2" onFocus="this.value=''" value="请输入关健字" size="30"> 
             <select name="s1" id="select">			   
                <OPTION VALUE="1">按标题</OPTION>                
                <OPTION VALUE="2">按内容</OPTION>
              </select><input type="submit" name="Submit" value="搜 索" ></td>
          </tr>
        </table></td>
    </form>
  </tr>
</table>
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
    <td align="center" height="30" class="ty3"><b><A href="news_list.asp?n=0,0,0&<%=C_ID%>">文章栏目</a></b></td>
  </tr>
  <tr>
    <td colspan="2" align="center">
<%=news_class(1,1,1,0,0,0,1,12,11,9,15,13,10,"news_list.asp")%>
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
    <td align="center" height="30" class="ty3"><b>推荐图文</b></td>
  </tr>

  <tr>
    <td colspan="2" align="center"><%=news(c1,c2,c3,1,0,0,0,1,0,0,0,0,0,8,1,24,13,20)%></td>
  </tr>     
</table>
      <table width="215" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td  height="7"></td>
        </tr>
      </table>
<table width="215" border="0" cellpadding="0" cellspacing="0"  class="ty1">
  <tr>
    <td align="center" height="30" class="ty3"><b>热门点击</b></td>
  </tr>
  <tr>
    <td height="22" align="center"><%=news(c1,c2,c3,3,0,0,0,1,0,0,0,0,0,8,1,24,13,20)%></td>
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