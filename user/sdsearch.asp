<!--#include file="sdtop.asp"-->
<%
'=====================================
'���칩����Ϣ��վ����ϵͳ
'��������Ƽ������ҳ�Ʒ
'�ͷ�����http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>
<%Dim rss,key,otype,page,pages,j
key=request.querystring("key")
otype=request.querystring("otype")
com=request.querystring("com")
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
             <table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr> 
                <td align="center" height="5"></td>
              </tr>			  
            </table>
<table width="960" border="0" align="center" cellpadding="0" cellspacing="0" >
  <tr> 
    <td width="190" height="200" align="center" valign="top" bgcolor="efefef"><!--#include file="sdleft.asp"-->
     </td>
	 <td width="5"> </td>
    <td align="center" valign="top" bgcolor="#ffffff"><table width="765" border="0" align="center" cellpadding="0" cellspacing="0" id="line">
<% 
page=clng(request.querystring("page"))	
set rss= server.createobject("adodb.recordset")
if otype="title" then
sql="select * from DNJY_info where biaoti like '%"& key &"%' and yz=1  and username='"&com&"' order by fbsj "&DN_OrderType&",ID "&DN_OrderType&"" 	
elseif otype="msg" then
sql="select * from DNJY_info where memo like '%"& key &"%' and yz=1  and username='"&com&"' order by fbsj "&DN_OrderType&",ID "&DN_OrderType&"" 	
else
end if
rss.open sql,conn,1,1
if rss.eof and rss.bof then
response.write "<tr bgcolor='#ffffff'><td colspan='4'><p align='center'>�Բ���û���ҵ������Ʒ</p></td></tr>"
else
%>
        <tr id="line5"> 
          <td width="9%" align="center">id</td>
          <td width="55%" align="center">��Ʒ����</td>
          <td width="15%" align="center">������</td>
          <td width="21%" align="center">��������</td>
        </tr>
<%
rss.pagesize=20
if page=0 then page=1 
pages=rss.pagecount
if page > pages then page=pages
rss.absolutepage=page  
for j=1 to rss.pagesize
%>
        <tr bgcolor="#ffffff"> 
          <td height="22" align="center"><%=rss("id")%></td>
          <td>��<a href="../<%=x_path("",rss("id"),"",rss("Readinfo"))%>"  target="_blank"><%response.write(boldword(rss("biaoti"),key))%></a></td>
          <td align="center"><%=left(rss("username"),5)%></td>
          <td align="center"><%if rss("fbsj")=date() then%><font color="#ff0000"><%else%><font color="#999999"><%end if%><%=rss("fbsj")%></font></td>
        </tr>
<%
rss.movenext
if rss.eof then exit for
next
end if
%>
      </table>
<br>
<div align="center">�ؼ���<font color="#ff0000"><strong><%=key%></strong></font>����Ϊ���ҵ�<font color="#ff0000"><%=rss.recordcount%></font>����Ʒ</div><table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr><form method=post action="?key=<%=key%>&otype=<%=otype%>&com=<%=com%>"> 
    <td align="center"> 
              <%if page<2 then      
    response.write "��ҳ ��һҳ&nbsp;"
  else
    response.write "<a href=?key="&key&"&otype="&otype&"&page=1&com="&com&">��ҳ</a>&nbsp;"
    response.write "<a href=?key="&key&"&otype="&otype&"&page=" & page-1 & "&com="&com&">��һҳ</a>&nbsp;"
  end if
  if rss.pagecount-page<1 then
    response.write "��һҳ βҳ"
  else
    response.write "<a href=?key="&key&"&otype="&otype&"&page=" & (page+1) & ">&com="&com&""
    response.write "��һҳ</a> <a href=?key="&key&"&otype="&otype&"&page="&rss.pagecount&"&com="&com&">βҳ</a>"
  end if
   response.write "&nbsp;ҳ�Σ�<strong><font color=red>"&page&"</font>/"&rss.pagecount&"</strong>ҳ "
    response.write "&nbsp;��<b><font color='#ff0000'>"&rss.recordcount&"</font></b>����¼ <b>"&rss.pagesize&"</b>����¼/ҳ"
   response.write " ת����<input type='text' name='page' size=4 maxlength=10 class=input value="&page&">"
   response.write " <input class=input type='submit'  value=' goto '  name='cndok'></span></p>"     
%>
          </form></td>
  </tr>
</table>
    </td>
  </tr>
</table>
<%rss.close
set rss=nothing%>
             <table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr> 
                <td align="center" height="5"></td>
              </tr>			  
            </table>
<!--#include file="sdend.asp"-->
