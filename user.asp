<!--#include file="conn.asp"-->
<!--#include file="SqlIn.Asp"-->
<!--#include file=top.asp-->
<!--#include file=usercookies.asp-->
<%
'=====================================
'���칩����Ϣ��վ����ϵͳ
'��������Ƽ������ҳ�Ʒ
'�ͷ�����http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>
<link rel="stylesheet" type="text/css" href="/<%=strInstallDir%>css/css.css">
<style type="text/css">
<!--
.style2 {color: #333333}
.style3 {color: #FF0000}
.style5 {color: #FF0000; font-weight: bold; }
-->
</style>
</head>

<body topmargin="0" leftmargin="0">

<div align="center">
  <center>
  <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="980" height="140">
    <tr>
      <td width="214" height="356" valign="top" bgcolor="#FFFFFF"><!--#include file=userleft.asp--></td>
	  <td width="5">&nbsp;</td>
      <td width="760" bgcolor="#FFFFFF"><table width="760" height="375" border="0" align="right" cellpadding="0" cellspacing="0" style="border-collapse: collapse" class="ty1">
        <tr bgcolor="#FAFAFA">
          <td height="25" colspan="7" ><span class="style5">������<%=rmb_mc%>���棺��վ��<%=xnpsj%> ��ÿ��VIP��Ա����<%=rmb_mc%><%=xnp%>����</span></td>
        </tr>
<%dim m,integrity
username=request.cookies("dick")("username")
set rs=conn.execute("select count(id) from [DNJY_info] where username='"&username&"'")
m=rs(0)
rs.close
set rs=server.createobject("adodb.recordset")
sql = "select * from [DNJY_user] where username='"&username&"'"
rs.open sql,conn,1,1
%>

        <tr>
          <td width="16%" height="28" align="right" background="img/22.gif"><img src="images/File_29.gif" width="16" height="16" border="0">�ʻ���</td>
          <td height="25" colspan="2" align="left" valign="middle" background="img/22.gif"><%=rs("Amount")%>Ԫ</td>
          <td width="15%" height="25" align="right" background="img/22.gif"><%=rmb_mc%>��</td>
          <td width="10%" height="25" align="left" valign="middle" background="img/22.gif"><%=rs("hb")%>��</td>
          <td width="16%" height="25" align="right" background="img/22.gif">����������</td>
          <td width="31%" height="25" align="left" valign="middle" background="img/22.gif"><%=rs("jf")%></td>
        </tr>
		
        <tr>
          <td width="16%" height="28" align="right" background="img/22.gif"><img src="images/File_29.gif" width="16" height="16" border="0">��½������</td>
          <td height="25" colspan="2" align="left" valign="middle" background="img/22.gif"><%=rs("dlcs")%>��</td>
          <td width="15%" height="25" align="right" background="img/22.gif">������Ϣ��</td>
          <td width="10%" height="25" align="left" valign="middle" background="img/22.gif"><%=m%>��</td>

          <td width="16%" height="25" align="right" background="img/22.gif">��Ա����</td>
          <td width="31%" height="25" align="left" valign="middle" background="img/22.gif"><%if vip>0 then
  response.write "<font color=#FF0000>VIP��Ա</font>"
if sj>0 then
response.Write "<font color=#FF0000>(�ʸ�ʣ��"&sj&"</b>��)</font>"
elseif sj=0 then
response.Write "<font color=""#414141"">(�ʸ���յ���)</font>"
elseif sj<0 then
response.Write "<font color=""#FF0000""> (�ʸ��Ѿ�����)</font>"
end If
else
response.write "<font color=#FF0000>��ͨ��Ա</font>"
response.write "(��������<a href='upgradevip.asp?"&C_ID&"'><font color=#0000ff>VIP��Ա</font></a>)"
end if%> <%if rs("sdname")<>"" then%><a  href="user/com.asp?com=<%=rs("username")%>&<%=C_ID%>" target="_blank"><img src="img/dian.gif" title="���ѿ�������" border="0"></a><%end if%> <%if rs("mpname")<>"" then%> <a  href="company.asp?username=<%=rs("username")%>&<%=C_ID%>" target="_blank"><img src="img/qy.gif" title="���ѷ�����ҵ��Ƭ" border="0"></a><%end if%></td>
        </tr>
        <tr>
          <td colspan="7" width="16%" height="28"  background="img/22.gif">&nbsp;&nbsp;&nbsp;<img src="images/File_29.gif" width="16" height="17" border="0">ע��ʱ�䣺<%=rs("zcdata")%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;��½ʱ�䣺<%=rs("dldata")%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;�����¼��<%=request.cookies("dick")("dldata")%></td>
        </tr>
        <tr>
          <td colspan="7" width="16%" height="28"  background="img/22.gif">&nbsp;&nbsp;&nbsp;<img src="images/File_29.gif" width="16" height="17" border="0">���ż�¼��<%=rs("goodfaith")%> �� <img src="img/hp.jpg">&nbsp;&nbsp;&nbsp;&nbsp; ʧ�ż�¼��<%=rs("bpromises")%> �� <img src="img/cp.jpg"></td>
        </tr>    		
        <tr>
          <td width="16%" height="28" align="right" background="img/22.gif" valign="top"><img src="images/File_29.gif" width="16" height="16" border="0">���Ѳ�����</td>
          <td height="25"  align="left" valign="middle" background="img/22.gif" valign="top"></td><%Dim df,cs,pj
df=rs("wevaluation")
cs=rs("participants")
If rs("wevaluation")=0 Then df=5
If rs("participants")=0 Then cs=1
pj=Formatnumber(df/(cs*5)*100,0,0,0,0)%>
          <td width="40%" align="left" background="img/22.gif" valign="top"> <%=cs%> �˴�&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;ƽ���÷� <%=Formatnumber(df/cs,1,0,0,0)%> ��</td>
          
          <td height="25" align="left" valign="middle" background="img/22.gif" colspan="6">	<style type="text/css"> 
body{ 
text-align:center; 
} 
.graph{ 
width:355px; 
border:1px solid #F8B3D0; 
height:15px; 
} 
#bar{ 
display:block; 
background:#FFE7F4; 
float:left; 
height:100%; 
text-align:center; 
} 
#barNum{ 
position:absolute; 
} 
</style> 
<script type="text/javascript"> 
function $(obj){ 
return document.getElementById(obj); 
} 
function go(){ 
$("bar").style.width = parseInt($("bar").style.width) + 1 + "%"; 
$("bar").innerHTML = $("bar").style.width; 
if($("bar").style.width =="<%=pj%>%"){ 
window.clearInterval(bar); 
} 

} 
var bar = window.setInterval("go()",40); 
window.onload = function(){ 
bar; 
} 
</script> 
<div class="graph"><strong id="bar" style="width:1%;"></strong></div>
<img src="img/pj.jpg" width="370"></td>
        </tr>
    
        <tr>
          <td height="10" colspan="7"></td>
        </tr>

        <tr>
          <td height="25" colspan="7"><table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="760" height="26">
              <tr>
                <td width="100%" height="26"><table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%">
                    <tr>
                      <td width="100%" height="10"></td>
                    </tr>
                    <tr>
                      <td height="5" bgcolor="#66CC00"></td>
                    </tr>
                    <tr>
                      <td height="8" bgcolor="#EEEEEE"></td>
                    </tr>
                    <tr>
                      <td height="10"></td>
                    </tr>
                </table></td>
              </tr>
          </table></td>
        </tr>
        <tr>
          <td height="25" colspan="7"><table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="99%" height="17">
              <tr>
                <td height="20"><div align="center">ע������������(�����������Ϸ����Ժ󷢲���Ϣ) [ <a href="user_data.asp?<%=C_ID%>">�޸�����</a>]</div></td>
              </tr>
              <tr>
                <td height="10">
				���뱣��<%If rs("problem")<>"" Then
                response.Write "<font color=#66CC00><b>��</b></font>"
				integrity=integrity+1
				Else
				response.Write "��"
				End if%>
				��������<%If rs("city_oneid")>0 Then
                response.Write "<font color=#66CC00><b>��</b></font>"
				integrity=integrity+1
				Else
				response.Write "��"
				End if%>
				��Ϣ����<%If rs("type_oneid")>0 Then
                response.Write "<font color=#66CC00><b>��</b></font>"
				integrity=integrity+1
				Else
				response.Write "��"
				End if%>
				��ʵ����<%If rs("name")<>"" Then
                response.Write "<font color=#66CC00><b>��</b></font>"
				integrity=integrity+1
				Else
				response.Write "��"
				End if%>
				��&nbsp;&nbsp;&nbsp;&nbsp;��
				<%response.Write "<font color=#66CC00><b>��</b></font>"
				integrity=integrity+1%>
				�� �� ֤<%If rs("idcard")<>"" Then
                response.Write "<font color=#66CC00><b>��</b></font>"
				integrity=integrity+1
				Else
				response.Write "��"
				End if%>
				�̶��绰<%If rs("dianhua")<>"" Then
                response.Write "<font color=#66CC00><b>��</b></font>"
				integrity=integrity+1
				Else
				response.Write "��"
				End if%>
				�ֻ�����<%If rs("CompPhone")<>"" Then
                response.Write "<font color=#66CC00><b>��</b></font>"
				integrity=integrity+1
				Else
				response.Write "��"
				End if%>
				�� �� ��<%If rs("fax")<>"" Then
                response.Write "<font color=#66CC00><b>��</b></font>"
				integrity=integrity+1
				Else
				response.Write "��"
				End if%>
				��������<%If rs("email")<>"" Then
                response.Write "<font color=#66CC00><b>��</b></font>"
				integrity=integrity+1
				Else
				response.Write "��"
				End if%>
				��ϵ QQ<%If rs("qq")<>"" Then
                response.Write "<font color=#66CC00><b>��</b></font>"
				integrity=integrity+1
				Else
				response.Write "��"
				End if%>
				<br>΢�ź���<%If rs("weixin")<>"" Then
                response.Write "<font color=#66CC00><b>��</b></font>"
				integrity=integrity+1
				Else
				response.Write "��"
				End if%>
				��˾����<%If rs("mpname")<>"" Then
                response.Write "<font color=#66CC00><b>��</b></font>"
				integrity=integrity+1
				Else
				response.Write "��"
				End if%>
				��˾��ַ<%If rs("http")<>"" Then
                response.Write "<font color=#66CC00><b>��</b></font>"
				integrity=integrity+1
				Else
				response.Write "��"
				End if%>
				ͨ�ŵ�ַ<%If rs("dizhi")<>"" Then
                response.Write "<font color=#66CC00><b>��</b></font>"
				integrity=integrity+1
				Else
				response.Write "��"
				End if%>
				��������<%If rs("youbian")<>"" Then
                response.Write "<font color=#66CC00><b>��</b></font>"
				integrity=integrity+1
				Else
				response.Write "��"
				End if%>
				��ͼ��ע<%If rs("Emap")<>"" Then
                response.Write "<font color=#66CC00><b>��</b></font>"
				integrity=integrity+1
				Else
				response.Write "��"
				End if%>
				������Ѷ<%response.Write "<font color=#66CC00><b>��</b></font>"
				integrity=integrity+1%>				
				</td>
              </tr>
              <tr>
                <td height="20">
<%Dim wzd
wzd=Formatnumber(integrity/18*100,0,0,0,0)%>
<style type="text/css"> 
body{ 
text-align:center; 
} 
.graph2{ 
width:650px; 
border:1px solid #F8B3D0; 
height:15px; 
} 
#wzd{ 
display:block; 
background:#FFE7F4; 
float:left; 
height:100%; 
text-align:center; 
} 
#wzdNum{ 
position:absolute; 
} 
</style> 
<script type="text/javascript"> 
function $(objwzd){ 
return document.getElementById(objwzd); 
} 
function go2(){ 
$("wzd").style.width = parseInt($("wzd").style.width) + 1 + "%"; 
$("wzd").innerHTML = $("wzd").style.width; 
if($("wzd").style.width =="<%=wzd%>%"){ 
window.clearInterval(wzd); 
} 

} 
var wzd = window.setInterval("go2()",40); 
window.onload = function(){ 
wzd; 
} 
</script> 
<div class="graph2"><strong id="wzd" style="width:1%;"></strong></div>
<img src="img/wzd.jpg" width="675"></td>
              </tr>
          </table></td>
        </tr>

        <tr>
          <td height="31" colspan="7"><table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%">
              <tr>
                <td width="100%" height="5" bgcolor="#FFCC33"></td>
              </tr>
              <tr>
                <td height="8" bgcolor="#EEEEEE"></td>
              </tr>
              <tr>
                <td height="10"></td>
              </tr>
          </table></td>
        </tr>
        <tr>
          <td width="16%" height="25" valign="top"><div align="right">����<font color="#CC5200">��ɫ</font>���ߣ�</div></td>
          <td height="25" colspan="2" align="right" valign="top"><div align="center"><%=rs("a")%>&nbsp;��</div></td>
          <td height="25" colspan="4"><span class="style2"> <img src="images/form2_r3_c3.gif"> ����ӵ�д˵���ʱ�����趨��������Ϣ������ɫ���Դ�����ȡ����û������ߣ�</span></td>
        </tr>
        <tr>
          <td width="16%" height="25" valign="top"><div align="right">��Ϣ<font color="#CC5200">�ö�</font>���ߣ�</div></td>
          <td height="25" colspan="2" align="right" valign="top"><div align="center"><%=rs("b")%>&nbsp;��</div></td>
          <td height="25" colspan="4"><span class="style2"> <img src="images/form2_r3_c3.gif"> ʹ��Խ�������Ϣ������ҳ������ҳ����ǰ����ʾ������Ч��ֻ��һ�죡</span></td>
        </tr>
        <tr>
          <td width="16%" height="25" valign="top"><div align="right">����<font color="#CC5200">��ͼ</font>���ߣ�</div></td>
          <td height="25" colspan="2" align="right" valign="top"><div align="center"><%=rs("c")%>&nbsp;��</div></td>
          <td height="25" colspan="4"><span class="style2"> <img src="images/form2_r3_c3.gif"> ʹ�ú��������Ϣ���������ͼƬ���Ա��û��˽������Ʒ��</span></td>
        </tr>
        <tr>
          <td width="16%" height="25" valign="top"><div align="right">ͨ��<font color="#CC5200">��֤</font>���ߣ�</div></td>
          <td height="25" colspan="2" align="right" valign="top"><div align="center"><%=rs("d")%>&nbsp;��</div></td>
          <td height="25" colspan="4"><span class="style2"> <img src="images/form2_r3_c3.gif"> ʹ�ú󲻾�������Ա��֤��ֱ����ʾ��������վ���������������ʱ�۸�ϸߣ�</span></td>
        </tr>
        <tr>
          <td height="25" colspan="7"><table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="760" height="26">
              <tr>
                <td width="100%" height="26"><table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%">
                    <tr>
                      <td width="100%" height="10"></td>
                    </tr>
                    <tr>
                      <td height="5" bgcolor="#66CC00"></td>
                    </tr>
                    <tr>
                      <td height="8" bgcolor="#EEEEEE"></td>
                    </tr>
                    <tr>
                      <td height="10"></td>
                    </tr>
                </table></td>
              </tr>
          </table></td>
        </tr>
        <tr>
          <td height="25" colspan="7"><table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%" height="17">
              <tr>
                <td width="185" height="48" background="img/alipaybutton.gif" ><div align="center"><a href="props.asp?<%=C_ID%>"><br>��Ҫ������Ҷһ�<%=rmb_mc%></a></div></td>
                <td width="75%" height="13" valign="top"><span class="style2"> ʹ�õ���������ø���Ȩ�ޣ�����Ҫʹ��<%=rmb_mc%>�һ����������<%=rmb_mc%>���㣬����������Ҷһ�<%=rmb_mc%>��һԪ����ҿɶһ�<font color=#ff0000><%=rmb_hb%></font>��<%=rmb_mc%>��������߰�ť�һ���</span> </td>
              </tr>
          </table></td>
        </tr>
        <tr>
          <td height="25" colspan="7"><table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="760" height="1">
              <tr>
                <td height="26" colspan="2"><table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%">
                    <tr>
                      <td width="100%" height="10"></td>
                    </tr>
                    <tr>
                      <td height="5" bgcolor="#CC99FF"></td>
                    </tr>
                    <tr>
                      <td height="8" bgcolor="#EEEEEE"></td>
                    </tr>
                    <tr>
                      <td height="10"><script src="<%=adjs_path("ads/js","info1",c1&"_"&c2&"_"&c3)%>"></script></td>
                    </tr>
                </table></td>
              </tr>

          </table></td>
        </tr>
      </table></td>
      <td width="10" valign="top" bgcolor="#e6e6e6"></td>
    </tr>
    <tr><%Call CloseRs(rs)%>
      <td height="26" colspan="3"><!--#include file=footer.asp--></td>
    </tr>
  </table>
  </center>
</div>
</body>
</html>

