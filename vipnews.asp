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
<script type="text/javascript" src="Include/data.js"></script>
<body topmargin="0">
<table width="980" height="29" border="0" align="center" cellpadding="0" cellspacing="0" background="img/gaobei_top2.gif">
     <tbody>
     <tr>
      <td width="92"></td>
       <td valign="center" width="211"><div style="padding-bottom:6px; color:#ffffff">
<script src="Include/date.js" type=text/javascript></script> <span id="time"></span>	<script>setInteval("time.innerHTML=new Date().getHours()+':'+newDate().getMinutes()+':'+new Date().getSeconds()",1000);</script>
       </div></td>
     <td width="26"></td>
 <td width="571"></td>
       </tr>
  </tbody>
</table>
<table width="980" height="100%" border="0" align="center" cellpadding="0" cellspacing="0" class=table-zuoyou bgcolor="#FFFFFF">
  <tr> 
    <td width="218" height="300" align="center" valign="top">
	<table width="100%" border="0" cellpadding="0" cellspacing="0"  class="ty1">
  <tr>
    <td height="30" align="center"class="ty3"><b>��Ѷ��Ŀ</b></td>
  </tr>
<%Call OpenConn
Dim sqlt,rst,rc
id=clng(request.querystring("id"))
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
do while not rst.eof
If rst("ID")=ID Then rc="ff0000"%>
        <tr> 
          <td height="30" align="center" class="dq2"><img src="img/zj.gif" width="15" height="15"><a  href="vipnews.asp?id=<%=rst("id")%>&<%=C_ID%>"><strong><span style="font-size:13px; font-weight:bold;color:<%=rc%>" ><%=rst("ClassName")%></strong></span></a></td>
        </tr>
<%rc=""
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
	<!--#include file="vipnews_left.asp"-->
	</td>
    <td height="100%" align="center" valign="top">
	<table width="550" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td height="5"> </td>
        </tr>
      </table>
	<table width="750" border="0" cellpadding="0" cellspacing="0" bgcolor="#ffffff">
<%Dim page,pages,j
page=clng(request.querystring("page"))
If request("ppage")<>"" Then page=clng(request("ppage"))
Call OpenConn
set rs=server.createobject("adodb.recordset")
If id=0 Then
sql="select * from DNJY_userNews where newsShared=0 and ifshow=1 order by newstop "&DN_OrderType&",id "&DN_OrderType&""
rs.open sql,conn,1,1
else
sql="select * from DNJY_userNews where newsShared=0 and ifshow=1 and ClassID="&id&" order by newstop "&DN_OrderType&",id "&DN_OrderType&""
rs.open sql,conn,1,1
end if
if rs.eof and rs.bof then
Response.Write "<p><br>��������!"
else 
Response.Write"<tr><td width=""5%"" height=""30"" rowspan=""2""><img src=""img/dian.gif""></td><td width=""89%"" height=""29"" align=""left"">��<strong>"
Response.Write conn.Execute("Select ClassName From DNJY_userNewsClass Where ID="&rs("ClassID")&"")(0)
Response.Write "</strong></td></tr><tr><td height=""1"" bgcolor=""#BBE5A6""><img src=""img/spacer.gif"" width=""1"" height=""1""></td></tr>"
rs.pagesize=20
if page=0 then page=1 
pages=rs.pagecount
if page > pages then page=pages
rs.absolutepage=page  
for j=1 to rs.pagesize 
if conn.Execute("Select sdkg From DNJY_user Where username='"&rs("username")&"'")(0)=1 then%>
        <tr> 
          <td height="24" align="center" ><%if rs("newstop")=1 then%><img src="img/ding_ico.gif" width="16" height="16" alt="�ö�����"><%else%><img src="img/pu_ico.gif" width="16" height="16" alt="��ͨ����"><%end if%></td>
          <td height="24" colspan="2" align="left" style="border-bottom: #999999 1px dotted"> <% if rs("SavePathFileName")<>"" then response.write "<img src='img/pic.gif' border=0 alt='ͼƬ����'>" end if %>
            <a  href="user/article.asp?id=<%=rs("id")%>&com=<%=rs("username")%>" target="_blank"><%= rs("title") %></a><%if rs("newstop")=1 then%><img src='img/top4.gif' border=0 alt='�ö�����'><%end if%>&nbsp;<%if rs("tj")=true then%><img src='img/jian.gif' border=0 alt='�ö�����'><%end if%> <font color="#999999"> (�Ķ�<font color="#ff0000"><%= rs("hits") %></font>��)&nbsp; [<%if rs("addtime")=date() then%><font color="#ff0000"><%else%><font color="#999999"><%end if%><%= dicksj2(rs("addtime")) %></font>] <font color="#999999">[������:<a  href="user/com.asp?com=<%=rs("username")%>" target="_blank"><%= rs("username") %></a>]</font>    
          </td>
        </tr>
        <%
End if
rs.movenext
if rs.eof then exit for
next
%>
        <tr valign="bottom"> 
            <td height="100%" colspan="3" align="center" > 
          <form method=post action="vipnews.asp?id=<%= id %>">  
              <%if page<2 then      
    response.write "��ҳ ��һҳ&nbsp;"
  else
    response.write "<a href=vipnews.asp?id="&id&"&page=1>��ҳ</a>&nbsp;"
    response.write "<a href=vipnews.asp?id="&id&"&page=" & page-1 & ">��һҳ</a>&nbsp;"
  end if
  if rs.pagecount-page<1 then
    response.write "��һҳ βҳ"
  else
    response.write "<a href=vipnews.asp?id="&id&"&page=" & (page+1) & ">"
    response.write "��һҳ</a> <a href=vipnews.asp?id="&id&"&page="&rs.pagecount&">βҳ</a>"
  end if
   response.write "&nbsp;ҳ�Σ�<strong><font color=red>"&page&"</font>/"&rs.pagecount&"</strong>ҳ "
    response.write "&nbsp;��<b><font color='#ff0000'>"&rs.recordcount&"</font></b>����¼ <b>"&rs.pagesize&"</b>����¼/ҳ"
   response.write " ת����<input type='text' name='ppage' size=4 maxlength=10 class=input value="&page&">"
   response.write " <input class=input type='submit'  value=' goto '  name='cndok'></span></p>"     
%>
          </form>  </td> 
         
        </tr>
<% 
end if
Call CloseRs(rs)
%>
      </table>
 <!--JS��̬�������÷�ʽ����JS�ļ�����վ�����棬ֻ�ܶ�̬ҳ���ã����뿪ʼ-->
  <script src="<%=adjs_path("ads/js","tail",c1&"_"&c2&"_"&c3)%>"></script>
  <!--���������-->

    </td>
  </tr>
</table>
<!--#include file=footer.asp-->
</body>
</html>


                                                              
                                                              
                                                              