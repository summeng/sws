<!--#include file="conn.asp"-->
<!--#include file="SqlIn.Asp"-->
<!--#include file="top.asp"-->
<%
'=====================================
'���칩����Ϣ��վ����ϵͳ
'��������Ƽ������ҳ�Ʒ
'�ͷ�����http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>
<%Call OpenConn
dim key,otype,idd,page,pages,j
t=clng(request.querystring("t"))
If t=0 Then
idd=""
else
idd=" and where ClassID="&id&""
End if
page=clng(request.querystring("page"))
key=trim(request.querystring("key"))
otype=request.querystring("otype")
if key="" then
   response.write "<script>alert('�����ַ�������Ϊ�գ�');history.back();</script>"
   response.end
end if
function boldword(strcontent,key) 
dim objregexp 
set objregexp=new regexp 
objregexp.ignorecase =true 
objregexp.global=true 
objregexp.pattern="(" & key & ")" 
strcontent=objregexp.replace(strcontent,"<font color=""#ff0000"">$1</font>" ) 
set objregexp=nothing 
boldword=strcontent 
end function
%>
<link rel="shortcut icon" href="favicon.ico">
</head>
<body topmargin="0">

  <table width="980" height="29" border="0" align="center" cellpadding="0" cellspacing="0" background="img/gaobei_top2.gif" bgcolor="#FFFFFF">
     <tbody>
         <tr>
            <td width="93"></td>
            <td valign="center" width="210"><div style="padding-bottom:6px; color:#FFFFFF"><div id='linkweb'></div><script>setInterval("linkweb.innerHTML=new Date().toLocaleString()+' ����'+'��һ����������'.charAt(new Date().getDay());",1000);</script></div></td>
             <td width="20"></td>
             <td width="577"><div style="padding-bottom:3px;"></td>
           </tr>
      </tbody>
    </table>
<table width="980" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#000000" class=table-zuoyou>
  <tr> 
    <td align="center" valign="top" bgcolor="#ffffff" width="180">
	<table width="100%" border="0" cellpadding="0" cellspacing="0"  class="ty1">
  <tr>
    <td height="30" align="center"class="ty3"><b>��Ѷ��Ŀ</b></td>
  </tr>
<%Call OpenConn
Dim sqlt,rst
id=clng(request.querystring("t"))
sqlt="select * from DNJY_userNewsClass order by showid"
set rst=server.createobject("adodb.recordset")
rst.open sqlt,conn,1,1
if rst.eof then
response.write "<CENTER>�����ڵ����!</CENTER>"
rst.close
set rst=nothing
Call CloseDB(conn)
response.End
Else
do while not rst.eof%>
        <tr> 
          <td height="30" align="center" class="dq2"><img src="img/zj.gif" width="15" height="15"><a  href="vipnews.asp?t=<%=rst("id")%>&<%=C_ID%>"><strong><span style="font-size:13px; font-weight:bold" ><%=rst("ClassName")%></strong></span></a></td>
        </tr>
<%
rst.movenext
loop 
rst.close
set rst=Nothing
End if%>
      </table>
	<table width="180" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td  height="7"></td>
        </tr>
      </table>
<table width="180" border="0" cellpadding="0" cellspacing="0" class="ty1">
  <tr>
    <td height="30" align="center"class="ty3">վ������</td>
  </tr>
  <tr>
    <td align="center"><table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#f7f7f7" class="tab_k">
      <form action="vipnewssearch.asp?t=<%=id%> method="get" name="form1" id="form1">
        <tr>
          <td height="33" align="center"class="dq2"><input name="key" type="text" class="input" id="key" size="19" /></td>
        </tr>
        <tr>
          <td height="35" align="center" class="dq2"><select name="otype" class="input" id="otype" style="line-height:30px;">
            <option value="title" selected="selected" class="input">����</option>
            <option value="msg" class="input">����</option></select>		  
		    <input name="t" type="hidden" class="input" id="id" value="<%=clng(request.querystring("id"))%>" size="19" />
            <input type="submit" name="submit2" value="����" class="input" /></td>
        </tr>
      </form>
    </table></td>
  </tr>
</table>
	  </td>  
    <td align="center" valign="top" bgcolor="#ffffff"><table width="800" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#000000">
<% Dim cs	
page=clng(request.querystring("page"))	
set rs= server.createobject("adodb.recordset")
if otype="title" then
sql="select * from DNJY_userNews where newsShared=0 and ifshow=1"&idd&" and title like '%"& key &"%' order by id "&DN_OrderType&""
elseif otype="msg" then
sql="select * from DNJY_userNews where newsShared=0 and ifshow=1"&idd&" and content like '%"& key &"%' order by id "&DN_OrderType&""
else
end if
rs.open sql,conn,1,1
if rs.eof and rs.bof Then
response.write "<tr bgcolor='#ffffff'><td colspan='4'><p align='center'>��Ǹ��δ�ҵ��������</p></td></tr>"
else
%>
        <tr bgcolor="#9999cc"> 
          <td width="9%" height="25" align="center" bgcolor="#BBE5A6">id</td>
          <td width="55%" align="center" bgcolor="#BBE5A6">��Դ����</td>
          <td width="15%" align="center" bgcolor="#BBE5A6">������</td>
          <td width="21%" align="center" bgcolor="#BBE5A6">��������</td>
        </tr>
<%
rs.pagesize=1
if page=0 then page=1 
pages=rs.pagecount
if page > pages then page=pages
rs.absolutepage=page  
for j=1 to rs.pagesize
%>
        <tr bgcolor="#ffffff"> 
          <td height="22" align="center"><%=rs("id")%></td>
          <td>��<% if rs("SavePathFileName")<>"" then response.write "<img src='img/pic.gif' border=0 alt='ͼƬ��Դ'>" end if %><a href="user/article.asp?id=<%=rs("id")%>&com=<%=rs("username")%>"  target="_blank"><%response.write(boldword(rs("title"),key))%></a><%if now()-rs("addtime")<5 then%><img src="img/new.gif" width="23" height="8" border=0 alt='������Դ'><%end if%></a> <%if rs("newstop")=1 then%><img src='img/top4.gif' border=0 alt='�ö���Դ'><%end if%>&nbsp;<%if rs("tj")=true then%><img src='img/jian.gif' border=0 alt='�Ƽ���Դ'><%end if%> <font color="#999999"> (�Ķ�<font color="#ff0000"><%= rs("hits") %></font>��)</td>
          <td align="center"><%=rs("username")%></td>
          <td align="center"><%if rs("addtime")=date() then%><font color="#ff0000"><%else%><font color="#999999"><%end if%><%=rs("addtime")%></font></td>
        </tr>
<%
rs.movenext
if rs.eof then exit for
next
end if
%>
      </table>
<br>
<div align="center">�ؼ���<font color="#ff0000"><strong><%=key%></strong></font>����Ϊ���ҵ�<font color="#ff0000"><%=rs.recordcount%></font>����Դ</div><table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr><form method=post action="?key=<%=key%>&t=<%=t%>&page=<%=page%>&otype=<%=otype%>"> 
    <td align="center"> 
              <%if page<2 then      
    response.write "��ҳ ��һҳ&nbsp;"
  else
    response.write "<a href=?key="&key&"&t="&t&"&otype="&otype&"&page=1>��ҳ</a>&nbsp;"
    response.write "<a href=?key="&key&"&t="&t&"&otype="&otype&"&page=" & page-1 & ">��һҳ</a>&nbsp;"
  end if
  if rs.pagecount-page<1 then
    response.write "��һҳ βҳ"
  else
    response.write "<a href=?key="&key&"&t="&t&"&otype="&otype&"&page=" & (page+1) & ">"
    response.write "��һҳ</a> <a href=?key="&key&"&t="&t&"&otype="&otype&"&page="&rs.pagecount&">βҳ</a>"
  end if
   response.write "&nbsp;ҳ�Σ�<strong><font color=red>"&page&"</font>/"&rs.pagecount&"</strong>ҳ "
    response.write "&nbsp;��<b><font color='#ff0000'>"&rs.recordcount&"</font></b>����¼ <b>"&rs.pagesize&"</b>����¼/ҳ"
   response.write " ת����<input type='text' name='page' size=4 maxlength=10 class=input value="&page&">"
   response.write " <input class=input type='submit'  value=' goto '  name='cndok'></span></p>"     
%>
          </td></form>
  </tr>
</table>
    </td>
  </tr>
</table><%
Call CloseRs(rs)
%> 
<!--#include file=footer.asp-->

</body>
</html>