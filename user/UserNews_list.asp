<!--#include file="sdtop.asp"-->
<%Dim classid,dhrs,dhsql,classname,keyword
keyword=request.Form("keyword")
if keyword="" then keyword=request.QueryString("keyword")
keyword=ReplaceUnsafe(keyword)
if keyword="" And request("key")<>"" then
	response.write"<SCRIPT language=JavaScript>alert('�������ѯ�ؼ��֣�');"
	response.write"javascript:history.go(-1)</SCRIPT>"
	response.End
end if
classid=ReplaceUnsafe(request.QueryString("classid"))
if not IsNumeric(classid) then
response.write"<SCRIPT language=JavaScript>alert('���id����');"
response.write"javascript:history.go(-1)</SCRIPT>"
response.End
end if
set dhrs=server.createobject("adodb.recordset")
dhsql="select * From DNJY_userNewsClass where id="&classid
dhrs.open dhsql,conn,1,1
if dhrs.eof then
response.write"<SCRIPT language=JavaScript>alert('���id�����ڣ�');"
response.write"javascript:history.go(-1)</SCRIPT>"
response.End
end if
classname=dhrs("classname")
classid=dhrs("id")
dhrs.close
set dhrs=nothing

%>
<head id="Head1">
<title>��������</title>
<link rel='stylesheet' type='text/css' href='../images/sdimg/<%=sdfg%>/style.css' />
</head> 
<body>
<table width="960" align="center" border="0" cellspacing="0" cellpadding="0">
<tr>
<td width="190" align="left" valign="top" class="top_m_txt01">
<!--#include file="sdleft.asp"-->
</td>
<td width="5"></td>
<td width="765" valign="top">                  
    <!--�ұ߿�ʼ-->
<table width="765" cellspacing="0" cellpadding="0" id="line">
<tr>
<td><div id="lt2"><span style="margin-left:30px;"><%=classname%></span></div></td>
</tr>
<%
Set rsnews=Server.CreateObject("ADODB.RecordSet") 
If keyword<>"" Then
sql="select id,Title,hits,addtime,ClassName  from DNJY_userNews where username='"&com&"' and title like '%"&keyword&"%' and  classid="& classid &" order by id desc"
Else
sql="select id,Title,hits,addtime,ClassName  from DNJY_userNews where username='"&com&"' and classid="& classid &" order by id desc"
End if
rsnews.Open sql,conn,1,1
if rsnews.eof and rsnews.bof then
	%>
		<tr style="background-color:White;height:20px;">
			<td align="left" style="width:90%;"><div align="center">�����������</div></td>
		</tr>
<%Dim page,pages,rsnews,j,turl,hits,addtime
else
rsnews.PageSize=30
	page=request.QueryString("page")
	if page="" then page=request.Form("page")
	page=clng(page)
	if page=0 then page=1 
	pages=rsnews.pagecount
	if page > pages then page=pages
	rsnews.AbsolutePage=page 
	
	for j=1 to rsnews.PageSize 
	Title = rsnews("Title")
		turl="article.asp?id="&rsnews("id")
		hits= rsnews("hits")
		addtime=rsnews("addtime")	
%>		
		<tr style="background-color:White;height:20px;">
			<td align="left" style="width:90%;">
                                    <li><a title="<%=title%>" href="<%=turl%>&com=<%=com%>" target="_blank" style="color:#000000;"><%=title%></a>
                                    <span>[<%=hits%>]</span> <span id="AxGridView1_ctl02_ctime"><%=addtime%></span></li></td>
		</tr>	
<%
rsnews.movenext
if rsnews.eof then exit for
next
%>		
		<tr>
			<td class="pagebottom" colspan="1"><div style="text-align:right;">
			  <div align="center">
			  <form method=Post action="UserNews_list.asp?classid=<%=classid%>&com=<%=com%>">
                <%if Page<2 then      
    response.write "��ҳ ��һҳ&nbsp;"
  else
    response.write "<a href=UserNews_list.asp?classid="&classid&"&page=1&com="&com&">��ҳ</a>&nbsp;"
    response.write "<a href=UserNews_list.asp?classid="&classid&"&page=" & Page-1 & "&com="&com&">��һҳ</a>&nbsp;"
  end if
  if rsnews.pagecount-page<1 then
    response.write "��һҳ βҳ"
  else
    response.write "<a href=UserNews_list.asp?classid="&classid&"&page=" & (page+1) & "&com="&com&">"
    response.write "��һҳ</a> <a href=UserNews_list.asp?classid="&classid&"&page="&rsnews.pagecount&"&com="&com&">βҳ</a>"
  end if
   response.write "&nbsp;ҳ�Σ�<strong><font color=red>"&Page&"</font>/"&rsnews.pagecount&"</strong>ҳ "
    response.write "&nbsp;��<b><font color='#FF0000'>"&rsnews.recordcount&"</font></b>����¼ <b>"&rsnews.pagesize&"</b>����¼/ҳ"
   response.write " ת����<input type='text' name='page' size=4 maxlength=10  value="&page&">"
   response.write " <input class=button type='submit'  value=' Goto '  name='cndok'></span>"     
%>
</form>
</td>
</tr>
<%end If
rsnews.close
set rsnews=nothing%>
	</table>
<table width="765" cellspacing="0" cellpadding="2" id="line">
<TR>
	<TD><form action="UserNews_list.asp?com=<%=com%>&classid=<%=classid%>&key=ok" method="post">
����������<input name="keyword" type="text" id="keyword" style="height: 18px; color: Gray;" class="input_text" />&nbsp;&nbsp;
<input type="submit" name="btnsearch" value="����"  id="btnsearch" class="btnsearch" />						
</form>	
</ul></TD>
</TR>
</TABLE>


	<%If jf<adjf Or adjf=0 then%><table width="100%" border="0" cellspacing="0" cellpadding="0">
            <tr> 
              <td align="center"><script language=javascript src="<%=adjs_path("ads/js","xxfl1",c1&"_"&c2&"_"&c3)%>"></script></td>
            </tr>
          </table><%End if%>
<!--�ұ߽���--></TD>
</TR>
</TABLE>                 
<!--#include file="sdend.asp"-->
</body>
</html>
