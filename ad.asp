<!--#include file="conn.asp"-->
<!--#include file="SqlIn.Asp"-->
<!--#include file=admin/cookies.asp-->
<!--#include file="config.asp"-->
<%
'=====================================
'���칩����Ϣ��վ����ϵͳ
'��������Ƽ������ҳ�Ʒ
'�ͷ�����http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>

<title>�������</title>
<link href="/<%=strInstallDir%>css/style.css" rel="stylesheet" type="text/css">
<BODY background="../images/back.gif">
<script language=javascript src=../Include/mouse_on_title.js></script>
<table width="100%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="#6699cc">
  <tr class=backs><td align="center" class=td height=18><font color="#FFFFFF">Ĭ�Ϲ����ʱ���أ�������Ҳ���Զ��رգ� ��������ر����˵���ֱ��</font></td>
  </tr>
  <tr> 
    <td height="200" valign="top" bgcolor="#FFFFFF"> 	
        <table width="100%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="#CCCCCC">
          <tr bgcolor="#FFFFFF"> 
            <td WIDTH="120">(�ɵ���޸�)<br>ͼƬ����*��</td>
			<td WIDTH="100" > λ��</td>
            <td WIDTH="400" > Ч����ʾ</td>
            <td WIDTH="50"> �Ƿ���ʾ</td>
          </tr>
          <tr> 
            <td colspan="7" ><strong>��վ��ҳ�������</strong>��</td>
          </tr>
		  <%Call OpenConn
		  Dim urlpa,url
		  set rs=server.CreateObject("adodb.recordset")
		  rs.open "select * from DNJY_ad",conn,1,1
		  do while not rs.eof%> 
           <tr bgcolor="#FFFFFF"> 
            <td >(<font color="#0000FF">id��<%=rs("ADID")%></font>) <%=rs("ADWidth")%>*<%=rs("ADHeight")%><br><a href="ads/edit.asp?id=<%=rs("id")%>&page=<%=urlpa%>">�޸�</a>  </td>
			<td><%=rs("ADNote")%></td>
              <td> <script src="ads/JS/<%=rs("ADID")%>_<%=rs("city_oneid")%>_<%=rs("city_twoid")%>_<%=rs("city_threeid")%>.js"></script></td>
             <td ><%if rs("adkg")=1 then%>��<%else%><font color="#FF0000">��</font><%end if%>&nbsp;&nbsp;&nbsp;&nbsp;<%if rs("adkg")=1 then%><a href="ads/createjs.asp?ad=zt&ADID=<%=rs("ADID")%>&city_oneid=<%=rs("city_oneid")%>&city_twoid=<%=rs("city_twoid")%>&city_threeid=<%=rs("city_threeid")%>&url=../ad.asp">��ͣ</a><%else%><a href="/<%=strInstallDir%>ads/createjs.asp?ad=hd&ADID=<%=rs("ADID")%>&city_oneid=<%=rs("city_oneid")%>&city_twoid=<%=rs("city_twoid")%>&city_threeid=<%=rs("city_threeid")%>&url=../ad.asp"><font color="#FF0000">�</font></a><%end if%></td> 
          </tr>
<%rs.movenext
Loop
Call CloseRs(rs)%>
                                          

        </table>
    </td>
  </tr>
</table>

