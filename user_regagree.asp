<!--#include file="conn.asp"-->
<!--#include file="SqlIn.Asp"-->
<!--#include file=top.asp-->
<!--#include file="Include/err.asp"-->
<%
'=====================================
'���칩����Ϣ��վ����ϵͳ
'��������Ƽ������ҳ�Ʒ
'�ͷ�����http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>
<%if lmkg2="1" then
call errdick(50)
Call CloseDB(conn)
response.end
end If
Dim tjname
tjname=trim(request("tjname"))%>
<link rel="stylesheet" href="/<%=strInstallDir%>css/css.css" type="text/css">

<center>
<TABLE width=980 border=0 cellpadding=3 cellspacing=1 align=center  bgcolor=#FFFFFF    class="grayline"  background="/img/hd_normal_bg3.gif">
<tr><td height=30>Ŀǰλ�ã�<a href=index.asp?<%=C_ID%>>��ҳ</a> > ��Աע���һ��</td></tr>
</TABLE>
<TABLE width=980 border=0 cellpadding=3 cellspacing=1 align=center  bgcolor=#FFFFFF height=300    class="grayline">

<tr><td>
<table border=0 width=547 align=center cellpadding=0 cellspacing=0  background=images/regbg2.gif>
<tr><td width=550 height=120><img src=img/reg.jpg border=0></td></tr>
<tr><td width=547 height=60 valign=middle bgcolor=#FFFFFF>
<br>�� ���ڱ�վע�ᣬ��Ϊ�Ա�վ�����������Э�顢������Լ������������������ܺ��Ͽ�<br><br>
</td></tr>
<tr><td width=547 height=12><img src=images/regbg1.gif border=0></td></tr>
<tr><td width=547 height=250  background=images/regbg2.gif>
����ӭ��ע��Ϊ��վ��Ա�����ڷ��ʱ�վ������̳��ʱ�������Ծ������������<br>

һ���������ñ�վΣ�����Ұ�ȫ��й¶�������ܣ������ַ�������Ἧ��ĺ͹���ĺϷ�Ȩ�棬�������ñ�վ���������ƺʹ���������Ϣ��<br> 
������һ��ɿ�����ܡ��ƻ��ܷ��ͷ��ɡ���������ʵʩ�ģ�<br>
����������ɿ���߸�������Ȩ���Ʒ���������ƶȵģ�<br>
����������ɿ�����ѹ��ҡ��ƻ�����ͳһ�ģ�<br>
�������ģ�ɿ�������ޡ��������ӣ��ƻ������Ž�ģ�<br>
�������壩�������������ʵ��ɢ��ҥ�ԣ������������ģ�<br>
��������������⽨���š����ࡢɫ�顢�Ĳ�����������ɱ���ֲ�����������ģ�<br>
�������ߣ���Ȼ�������˻���������ʵ�̰����˵ģ����߽����������⹥���ģ�<br>
�������ˣ��𺦹��һ��������ģ�<br>
�������ţ�����Υ���ܷ��ͷ�����������ģ�<br>
������ʮ��������١�α����Ʒ������������Ϣ��<br>

�����������أ����Լ������ۺ���Ϊ����<br>
����ͬ�Ⲣ�������¸����е�<%=title%>�û�Э�鼰����ص������������<br>
<br>������<A href="help.asp?id=2&<%=C_ID%>" target=_blank><%=title%>�û�Э��</A>������ص������������</td></tr>

<tr><td width=547 height=12><img src=images/regbg3.gif border=0></td></tr></table>
<%Response.Cookies("reg")("regkey")="ok" %>
<table border=0 width=547 align=center cellpadding=0 cellspacing=0>
<tr><td height="100" align="center">
<a href=<%=Regto%>?tjname=<%=tjname%>&<%=C_ID%>><img border=0 src="images/agree.gif" title="��ͬ��"></a>
&nbsp;&nbsp;&nbsp;&nbsp;
<a href=index.asp?<%=C_ID%>><img border=0 src="images/agree_no.gif" title="��ͬ�⣬�ٿ���һ��"></a>
</form>
</td>
</tr>
 
</table>
</td></tr>
</table>
</center>
</BODY>
</HTML>

<%Call OpenConn%>