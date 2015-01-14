<!--#include file="sdtop.asp"-->
<%Dim arrs,arsql,classname,classid,arttitle,Descrip
id=ReplaceUnsafe(request.QueryString("id"))
if not IsNumeric(id) then
response.write"<SCRIPT language=JavaScript>alert('id号错误！');"
response.write"javascript:history.go(-1)</SCRIPT>"
response.End
end if

set arrs=server.createobject("adodb.recordset")
	arsql="select * from DNJY_userNews where id="&id

	arrs.open arsql,conn,1,3
if arrs.eof then
response.write"<SCRIPT language=JavaScript>alert('没有此文章！');"
response.write"javascript:history.go(-1)</SCRIPT>"
response.End
end if
	classname=arrs("classname")
	classid=arrs("classid")
	arttitle=arrs("title")
	if arrs("keywords")<>"" then
	Keywords=arrs("keywords")
	else
	Keywords=arrs("title")&"_"&webname
	end if
	if arrs("Descrip")<>"" then
	Descrip=arrs("Descrip")
	else
	descrip=arrs("content")
	descrip=CutStr(RemoveHTML(descrip),150,"")
	end if


%>

<head id="Head1">
<title><%=arrs("title")%>_<%=webname%></title>
<meta http-equiv="X-UA-Compatible" content="IE=EmulateIE7" />
<meta name="Keywords" content="<%=keywords%>" />
<meta name="Description" content="<%=descrip%>" />
<link rel="stylesheet" href="../css/news.css" type="text/css" />
</head> 
<body>
<table width="960" align="center" border="0" cellspacing="0" cellpadding="0">
<TR>
	<td width="190" align="left" valign="top" class="top_m_txt01"><!--#include file="sdleft.asp"--></TD>
	<td width="5"></td>
	<TD width="765" valign="top">
	<table width="765" border="0" cellpadding="0" cellspacing="0" id="line">
<tr>
<td><div id="lt2"><span style="margin-left:30px;">详细内容</span></div></td>
</tr>
	<tr> <td>
	<!--右边开始-->
<br>
   <center>             
                       <p> <span id="title"><%=arttitle%></span></p>
               字体：<a href="javascript:doZoom(16)">大</a> <a href="javascript:doZoom(14)">中</a> <a href="javascript:doZoom(12)">小</a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;    
                 
				<%	Dim hits,origin,author				
				hits=arrs("hits")
				arrs("hits")=hits+1
				arrs.update
				response.Write " <span>阅读</span> " & hits+1 & " 次&nbsp;&nbsp;&nbsp;&nbsp;"	
			  if arrs("origin")<>"" then
			  origin=arrs("origin")
			  else
			  origin=webname
			  end if
			  response.Write "<span>来源：</span>"&origin
	          response.Write "&nbsp;&nbsp;&nbsp;&nbsp;"
			  if arrs("author")<>"" then
			  author=arrs("author")
			  else
			  author=com
			  end if
			  response.Write "<span>作者：</span>"&author			 
	response.Write "&nbsp;&nbsp;&nbsp;&nbsp;<span>日期：</span>"&arrs("addtime")%>
<br>
					 
					
                    
                       <hr />
     </center> <table width="95%" align="center">
	 <%If jf<adjf Or adjf=0 then%><div style="PADDING-RIGHT: 10px; FLOAT: left; widht: 280"><script language=javascript src="<%=adjs_path("ads/js","info3",c1&"_"&c2&"_"&c3)%>"></script></div><%End if%><div class="top_m_txt01" style="font-size:14px;line-height:25px;"><DIV><span id="zoom">&nbsp;<%=ContentPagination(arrs("content"))%></span></DIV>
	 <%If jf<adjf then%><script language=javascript src="<%=adjs_path("ads/js","xxfl1",c1&"_"&c2&"_"&c3)%>"></script><%End if%>
	 </table></TD>
</TR>
                    
                     <!--右边结束--></TABLE>
		 
					 
					 </TD>
</TR>
</TABLE>                
<div class="footer">
<!--#include file="sdend.asp"-->
</div>
</body>
</html>
