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
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<script type="text/javascript" src="admin/inc/js_left.js"></script>
<link rel="stylesheet" href="css/news.css" type="text/css" />
<link rel="shortcut icon" href="favicon.ico">
</head>
<%Call OpenConn
dim key,otype,oyaya_title,oyaya_keywords,oyaya_description,author,owen1,owen2,page,pages,j,sqlt,rst,rc
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

<body topmargin="0">

  <table width="980" height="29" border="0" align="center" cellpadding="0" cellspacing="0" background="img/gaobei_top2.gif" bgcolor="#FFFFFF">
     <tbody>
         <tr>
            <td width="93"></td>
            <td valign="center" width="210"><div style="padding-bottom:6px; color:#FFFFFF"><div id='linkweb'></div><script>setInterval("linkweb.innerHTML=new Date().toLocaleString()+' ����'+'��һ����������'.charAt(new Date().getDay());",1000);</script></div></td>
             <td width="20"></td>
             <td width="577"><div style="padding-bottom:3px;"><font color="#008af7">���������ڵ�λ�ã�</font><a href="/<%=strInstallDir%>">��ҳ</a> &gt;&gt; վ������ </div></td>
           </tr>
      </tbody>
    </table>
<table width="980" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#000000" class=table-zuoyou>
  <tr> 
   <td align="center" valign="top" bgcolor="#ffffff" width="210">
	<table width="100%" border="0" cellpadding="0" cellspacing="0"  class="ty1">
  <tr>
    <td height="30" align="center"class="ty3"><b>��Ѷ��Ŀ</b></td>
  </tr></table>
<table width="100%" border="0" cellpadding="0" cellspacing="0"  class="ty1">
        <tr>
          <td><%=news_class(1,1,1,0,0,0,1,12,11,9,15,13,10,"news.asp")%></td>
        </tr>
 </table>
	<table width="210" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td  height="7"></td>
        </tr>
      </table>
<table width="210" border="0" cellpadding="0" cellspacing="0" class="ty1">
  <tr>
    <td height="30" align="center"class="ty3"><b>վ������</b></td>
  </tr>
  <tr>
    <td align="center"><table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#f7f7f7" class="tab_k">
      <form action="newssearch.asp" method="get" name="form1" id="form1">
        <tr>
          <td height="33" align="center"class="dq2"><input name="key" type="text" class="input" id="key" size="19" /></td>
        </tr>
        <tr>
          <td height="35" align="center" class="dq2"><select name="otype" class="input" id="otype" style="line-height:30px;">
            <option value="title" selected="selected" class="input">����</option>
            <option value="msg" class="input">����</option>
            </select>  
			<input name="c1" type="hidden" class="input" id="c1" value="<%=c1%>" size="19" />
			<input name="c2" type="hidden" class="input" id="c2" value="<%=c2%>" size="19" />
			<input name="c3" type="hidden" class="input" id="c3" value="<%=c3%>" size="19" />
            <input type="submit" name="submit2" value="����" class="input" /></td>
        </tr>
      </form>
    </table></td>
  </tr>
</table>
	  </td>
    <td align="center" valign="top" bgcolor="#ffffff"><table width="760" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#000000">
<% Dim cs
Call OpenConn
page=clng(request.querystring("page"))	
set rs= server.createobject("adodb.recordset")
if otype="title" then
sql="select * from DNJY_News where ifshow=1 and title like '%"& key &"%'"&ttccnews&" order by id "&DN_OrderType&""
elseif otype="msg" then
sql="select * from DNJY_News where ifshow=1 and content like '%"& key &"%'"&ttccnews&" order by id "&DN_OrderType&""
else
end if
rs.open sql,conn,1,1
if rs.eof and rs.bof Then
IF c1>0 Then cs=conn.Execute("Select city From DNJY_city Where id="&c1&" and twoid="&c2&" and threeid="&c3&"")(0)
response.write "<tr bgcolor='#ffffff'><td colspan='4'><p align='center'>�Բ���û���ҵ�"&cs&"���Ż��������</p></td></tr>"
else%>
        <tr bgcolor="#9999cc"> 
          <td width="9%" height="25" align="center" bgcolor="#BBE5A6">id</td>
          <td width="55%" align="center" bgcolor="#BBE5A6">���ű���</td>
          <td width="15%" align="center" bgcolor="#BBE5A6">������</td>
          <td width="21%" align="center" bgcolor="#BBE5A6">��������</td>
        </tr>
<%rs.pagesize=1
if page=0 then page=1 
pages=rs.pagecount
if page > pages then page=pages
rs.absolutepage=page  
for j=1 to rs.pagesize%>
        <tr bgcolor="#ffffff"> 
          <td height="22" align="center"><%=rs("id")%></td>
          <td><% if rs("img")=True then response.write "<img src='img/pic.gif' border=0 alt='�����'>" end if %><a href="news_show.asp?id=<%=rs("id")%>&<%=C_ID%>"  target="_blank"><%response.write(boldword(rs("title"),key))%></a><%if now()-rs("infotime")<10 then%><img src="img/new.gif" width="23" height="8" border=0 alt='��������'><%end if%></a> <%if rs("newstop")=1 then%><img src='img/top4.gif' border=0 alt='�ö�����'><%end if%>&nbsp;<%if rs("tuijian")=1 then%><img src='img/jian.gif' border=0 alt='�Ƽ�����'><%end if%> <font color="#999999"> (�Ķ�<font color="#ff0000"><%= rs("hits") %></font>��)</td>
          <td align="center"><%=left(rs("newsuser"),5)%></td>
          <td align="center"><%if rs("infotime")=date() then%><font color="#ff0000"><%else%><font color="#999999"><%end if%><%=rs("infotime")%></font></td>
        </tr>
<%rs.movenext
if rs.eof then exit for
next
end if%>
      </table>
<br>
<div align="center">�ؼ���<font color="#ff0000"><strong><%=key%></strong></font>����Ϊ���ҵ�<font color="#ff0000"><%=rs.recordcount%></font>������</div><table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr><form method=post action="?key=<%=key%>&page=<%=page%>&otype=<%=otype%>&<%=C_ID%>"> 
    <td align="center"> 
              <%if page<2 then      
    response.write "��ҳ ��һҳ&nbsp;"
  else
    response.write "<a href=?key="&key&"&otype="&otype&"&page=1&"&C_ID&">��ҳ</a>&nbsp;"
    response.write "<a href=?key="&key&"&otype="&otype&"&page=" & page-1 & "&"&C_ID&">��һҳ</a>&nbsp;"
  end if
  if rs.pagecount-page<1 then
    response.write "��һҳ βҳ"
  else
    response.write "<a href=?key="&key&"&otype="&otype&"&page=" & (page+1) & "&"&C_ID&">"
    response.write "��һҳ</a> <a href=?key="&key&"&otype="&otype&"&page="&rs.pagecount&"&"&C_ID&">βҳ</a>"
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
</table><%Call CloseRs(rs)%> 
<!--#include file=footer.asp-->
</body>
</html>