<!--#include file="conn.asp"-->
<!--#include file="top.asp"-->
<!--#include file="SqlIn.Asp"-->
<%
'=====================================
'���칩����Ϣ��վ����ϵͳ
'��������Ƽ������ҳ�Ʒ
'�ͷ�����http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>
<%Dim Data
Call OpenConn
ID=Clng(Strint(Request.QueryString("ID")))
Set Rs=Conn.Execute("Select Text From DNJY_help Where OneID="&city_oneid&" and TwoID="&city_twoid&" and ThreeID="&city_threeid&"")
If Not Rs.Eof Then
	Data=Split(Rs(0),"�ӡӡ�")
Else
Data=Split(Application(CacheName&"_help"),"�ӡӡ�")
End If
Call CloseRs(rs)
Call CloseDB(conn)%>
<link href="Css/help.css" rel="stylesheet" type="text/css" />
<div class="content xw">
	<div class="l leftBar">
		<ul class="listMenu">
			<li <%If ID=0 Then%>class="current"<%End If%>><a href="/<%=strInstallDir%>help.asp?ID=0&<%=C_ID%>">���ڱ�վ</a></li>
			<li <%If ID=1 Then%>class="current"<%End If%>><a href="/<%=strInstallDir%>help.asp?id=1&<%=C_ID%>">��Ա����</a></li>
			<li <%If ID=2 Then%>class="current"<%End If%>><a href="/<%=strInstallDir%>help.asp?id=2&<%=C_ID%>">��������</a></li>
			<li <%If ID=3 Then%>class="current"<%End If%>><a href="/<%=strInstallDir%>help.asp?id=3&<%=C_ID%>">��������</a></li>
			<li <%If ID=4 Then%>class="current"<%End If%>><a href="/<%=strInstallDir%>help.asp?id=4&<%=C_ID%>">������</a></li>
			<li <%If ID=5 Then%>class="current"<%End If%>><a href="/<%=strInstallDir%>help.asp?id=5&<%=C_ID%>">��ϵ����</a></li>

		</ul>
	</div>
	<div class="r rightTxt">
	<%Select Case ID
	Case 0%>
		<div class="view">
			<h2 style="line-height:170%">���ڱ�վ</h2>
			<div class="txt"><%=Data(0)%></div>
		</div>
		<%CAse 1%>
		<div class="view">
			<h2 style="line-height:170%">��Ա����</h2>
			<div class="txt"><%=Data(1)%></div>
		</div>
<TABLE>
<TR>
	<TD height=10></TD>
</TR>
</TABLE>
<TABLE style="BORDER-COLLAPSE: collapse" border=0 cellSpacing=0 borderColor=#111111 cellPadding=0 width=700 height=240>
<TBODY>
<TR>
<TD bgColor=#ffcc33 height=5 width="100%">
<DIV class=bd align=center><IMG border=0 src="img/vip1.jpg" width=760 height=240></DIV></TD></TR></TBODY></TABLE>
<TABLE style="BORDER-BOTTOM: #cc9900 1px dotted; BORDER-LEFT: #cc9900 1px dotted; BORDER-TOP: #cc9900 1px dotted; BORDER-RIGHT: #cc9900 1px dotted" id=table1 border=0 cellSpacing=1 width="100%">
<TBODY>
<TR>
<TD>
<P style="LINE-HEIGHT: 150%; MARGIN: 0px 15px"><B>�� <U>��վ����</U></B></P>
<P style="LINE-HEIGHT: 150%; MARGIN: 0px 15px">�ڱ�վ���ɷ�����ѯ������Ϣ������������̡�������ҵ��Ƭ��</P>
<P style="LINE-HEIGHT: 150%; MARGIN: 0px 15px"><B>�� <U>ע���û�</U></B></P>
<P style="LINE-HEIGHT: 150%; MARGIN: 0px 15px">��վ�û���Ϊ<FONT color=#ff0000><B>�ο�</B></FONT>��<FONT color=#ff0000><B>��ͨ��Ա</B></FONT>��<FONT color=#ff0000><B>VIP��Ա</B></FONT>���֡��ο���ָ��ע�ᣬֱ�ӷ����Ͳ�ѯ��Ϣ�����ѡ����ע�ἴ�ɳ�Ϊ��վ��ͨ��Ա����ʹ�õ��߹��ܣ�ӵ��ǿ��ĺ�̨����ƽ̨�����ɻ�Ա�Ѽ��ɳ�Ϊvip��Ա����ӵ����ͨ��ԱȨ���⣬���öȸ��ߣ��ɿ����̡���ҵ��Ƭ����Ϣ����ˡ���ñ�վ����
<SCRIPT src="user/helppz.asp?action=rmb_mc"></SCRIPT>
�ȡ�����������Ȩ�޼��±�</P></TD></TR></TBODY></TABLE>
<DIV align=center>
<TABLE style="BORDER-COLLAPSE: collapse" border=1 cellSpacing=0 borderColor=#fff2d7 cellPadding=0 width="98%">
<TBODY>
<TR>
<TD width="8%">
<DIV align=center><FONT color=#b548e1>��Ա����</DIV></FONT></TD>
<TD width="8%">
<DIV align=center><FONT color=#b548e1>������� (Ԫ/��)</DIV></FONT></TD>
<TD width="5%">
<DIV align=center><FONT color=#b548e1>����ʹ��</DIV></FONT></TD>
<TD width="5%">
<DIV align=center><FONT color=#b548e1>�ϴ�ͼƬ</DIV></FONT></TD>
<TD width="8%">
<DIV align=center><FONT color=#b548e1>��������Ƭ</DIV></FONT></TD>
<TD width="8%">
<DIV align=center><FONT color=#b548e1>������Ʒչʾ(��)</DIV></FONT></TD>
<TD width="8%">
<DIV align=center><FONT color=#b548e1>�����Ƽ�</DIV></FONT></TD>
<TD width="8%">
<DIV align=center><FONT color=#b548e1>������Ϣ��ʾ(��)</DIV></FONT></TD>
<TD width="6%">
<DIV align=center><FONT color=#b548e1>HTM�������</DIV></FONT></TD>
<TD width="6%">
<DIV align=center><FONT color=#b548e1>��Ϣ����</DIV></FONT></TD>
<TD width="6%">
<DIV align=center><FONT color=#b548e1>��Ϣ�Ƽ�</DIV></FONT></TD>
<TD width="6%">
<DIV align=center><FONT color=#b548e1>������ҵ��Ƭ</DIV></FONT></TD>
<TD width="8%">
<DIV align=center><FONT color=#b548e1>����������Ϣ</DIV></FONT></TD>
<TD width="4%">
<DIV align=center><FONT color=#b548e1>�޸���Ϣ</DIV></FONT></TD>
<TD width="6%">
<DIV align=center><FONT color=#b548e1>������
<SCRIPT src="user/helppz.asp?action=rmb_mc"></SCRIPT>
 </DIV></FONT></TD></TR>
<TR>
<TD width="8%">
<DIV align=center><FONT color=#008080>�� ��</FONT></DIV>
<TD width="8%">
<DIV align=center>��</DIV></TD>
<TD width="5%">
<DIV align=center>��</DIV></TD>
<TD width="5%">
<DIV align=center>��</DIV></TD>
<TD width="8%">
<DIV align=center>��</DIV></TD>
<TD width="8%">
<DIV align=center>��</DIV></TD>
<TD width="8%">
<DIV align=center>��</DIV></TD>
<TD width="8%">
<DIV align=center>����
<SCRIPT src="user/helppz.asp?action=huiyuanxx"></SCRIPT>
</DIV></TD>
<TD width="6%">
<DIV align=center>��</DIV></TD>
<TD width="8%">
<DIV align=center>��</DIV></TD>
<TD width="6%">
<DIV align=center>��</DIV></TD>
<TD width="6%">
<DIV align=center>��</DIV></TD>
<TD width="8%">
<DIV align=center>��</DIV></TD>
<TD width="4%">
<DIV align=center>��</DIV></TD>
<TD width="6%">
<DIV align=center>��</DIV></TD></TR>
<TR>
<TD width="8%">
<DIV align=center><FONT color=#008080>��ͨ��Ա</FONT></DIV></TD>
<TD width="8%">
<DIV align=center>��</DIV></TD>
<TD width="5%">
<DIV align=center>��</DIV></TD>
<TD width="5%">
<DIV align=center>��</DIV></TD>
<TD width="8%">
<DIV align=center>��</DIV></TD>
<TD width="8%">
<DIV align=center>����
<SCRIPT src="user/helppz.asp?action=huiyuansp"></SCRIPT>
</DIV></TD>
<TD width="8%">
<DIV align=center>��</DIV></TD>
<TD width="8%">
<DIV align=center>����
<SCRIPT src="user/helppz.asp?action=huiyuanxx"></SCRIPT>
</DIV></TD>
<TD width="6%">
<DIV align=center>��</DIV></TD>
<TD width="6%">
<DIV align=center>��</DIV></TD>
<TD width="6%">
<DIV align=center>��</DIV></TD>
<TD width="6%">
<DIV align=center>��</DIV></TD>
<TD width="8%">
<DIV align=center>��</DIV></TD>
<TD width="4%">
<DIV align=center>��</DIV></TD>
<TD width="6%">
<DIV align=center>��</DIV></TD></TR>
<TR>
<TD width="8%">
<DIV align=center><FONT color=#008080>VIP��Ա</FONT></DIV></TD>
<TD width="8%">
<DIV align=center>
<SCRIPT src="user/helppz.asp?action=vipje"></SCRIPT>
</DIV></TD>
<TD width="5%">
<DIV align=center>��</DIV></TD>
<TD width="5%">
<DIV align=center>��</DIV></TD>
<TD width="8%">
<DIV align=center>��רҳ</DIV></TD>
<TD width="8%">
<DIV align=center>����</DIV></TD>
<TD width="8%">
<DIV align=center>��</DIV></TD>
<TD width="8%">
<DIV align=center>����</DIV></TD>
<TD width="6%">
<DIV align=center>��</DIV></TD>
<TD width="6%">
<DIV align=center>��</DIV></TD>
<TD width="6%">
<DIV align=center>��</DIV></TD>
<TD width="6%">
<DIV align=center>��</DIV></TD>
<TD width="8%">
<DIV align=center>��</DIV></TD>
<TD width="4%">
<DIV align=center>��</DIV></TD>
<TD width="6%">
<DIV align=center>��</DIV></TD></TR></TBODY></TABLE></DIV>ע�� a������Ϊ��ϵͳ��������ƶ�ʱҪ��ˡ�����ƶȷּ����С� b�������̻���ҵ��Ƭ�������֤��������վ�ṩ�������֤��ӡ�������й̶��绰��ϵ��ʽ���ɱ�վ��֤�󱸰���������վ����ʾΪ��֤��Ա�� c����ҳ������Ϣ�������ö�&gt;����ʱ�䡱����˳����ҳ�в�������Ϣ�����ö�&gt;���۸ߵ�&gt;����ʱ�䡱����˳������ֻ�о�����Ϣ��VIP��Ա��Ϣ������ʾͼƬ��
<P>
<TABLE style="BORDER-COLLAPSE: collapse" border=0 cellSpacing=0 borderColor=#111111 cellPadding=0 width="100%" height=17>
<TBODY>
<TR>
<TD height=13 vAlign=top width="31%">
<P style="LINE-HEIGHT: 150%; MARGIN: 0px 15px"><B>�� <U>������Ϣ</U></B></P>
<P style="LINE-HEIGHT: 150%; MARGIN: 0px 15px">���[������Ϣ]�����ѵ�½��������½��������Ϣ���ɣ�Ҳ����ѡ����Ҫ������ѷ���������Ϣ��������Ϣ��Ҫͨ�����,�Ҳ����ϴ�ͼƬ��</P>
<P style="LINE-HEIGHT: 150%; MARGIN: 0px 15px"><B>�� <U>ȡ����Ϣ</U></B></P>
<P style="LINE-HEIGHT: 150%; MARGIN: 0px 15px">�������Ϣ��������Ч���ɵ�¼������ɾ��������ɽ��׵ģ���¼�󣬽�����Ϣ���������������޸ģ���������Ϣ����Ϊ������ɽ��ס����ѳɽ�����Ϣ��������ʾ��������Ϣ�����ܸ��š�</P>
<P style="LINE-HEIGHT: 150%; MARGIN: 0px 15px"><B>�� <U>������Ϣ</U></B></P>
<P style="LINE-HEIGHT: 150%; MARGIN: 0px 15px">������Ϣ��������λ����ʾ������������Ϣʱ���������ʻ����㹻����ǰ���£�ѡ�񷢲���������д�����ܽ�ϵͳ�Զ��۳����á�ÿ�����Խ�ߣ�����������ϢԽ��ǰ��</P>
<P style="LINE-HEIGHT: 150%; MARGIN: 0px 15px"><B>�� <U>���߹��ܽ���</U></B></P>
<P style="LINE-HEIGHT: 150%; MARGIN: 0px 15px">��վʹ���ĸ����ߣ��������ɫ���ߣ��ı�����������ɫ���������򣩣���Ϣ�ö����ߣ�<FONT face=����>�ܷ�����Ϣ�ö�����Чʱ��Ϊ24Сʱ��ʹ�ø���Խ�࣬λ��Խ��</FONT>��������ͬ��ʱ��Խ�磬λ��Խ�ߣ���������ͼ���ߣ�<FONT face=����>���Է�����Ϣ��ص�ͼƬ</FONT>����ͨ����֤���ߣ�<FONT face=����>�ɲ�ͨ��������ˣ�ֱ�ӷ���</FONT>�������ֵ�������
<SCRIPT src="user/helppz.asp?action=rmb_mc"></SCRIPT>
��������ѡ�շ���Ŀ������±�</FONT></P></TD>
<TD height=13 vAlign=top width="69%"><A href="upgradevip.asp"><IMG border=0 src="images/vip3.gif"></A></TD></TR></TBODY></TABLE>
<DIV align=center>
<TABLE style="BORDER-COLLAPSE: collapse" border=1 cellSpacing=0 borderColor=#fff2d7 cellPadding=0 width="98%">
<TBODY>
<TR>
<TD>
<DIV align=center>
<P style="LINE-HEIGHT: 150%; MARGIN: 0px 15px">������</P></DIV></TD>
<TD>
<DIV align=center>
<P style="LINE-HEIGHT: 150%; MARGIN: 0px 15px">ʵ�ֹ���</P></DIV></TD>
<TD>
<DIV align=center>
<P style="LINE-HEIGHT: 150%; MARGIN: 0px 15px">��
<SCRIPT src="user/helppz.asp?action=rmb_mc"></SCRIPT>
</P></DIV></TD></TR>
<TR>
<TD width="24%">
<DIV align=center>
<P style="LINE-HEIGHT: 150%; MARGIN: 0px 15px"><FONT color=#008080>�����ɫ����</FONT></FONT></P></DIV></TD>
<TD width="60%">
<P style="LINE-HEIGHT: 150%; MARGIN: 0px 15px"><FONT face=����>�ı������ɫ</FONT></P></TD>
<TD width="15%">
<DIV align=center>
<SCRIPT src="user/helppz.asp?action=g_a"></SCRIPT>
</DIV></TD></TR>
<TR>
<TD>
<DIV align=center>
<P style="LINE-HEIGHT: 150%; MARGIN: 0px 15px"><FONT color=#008080>��Ϣ�ö�����</FONT></FONT></P></DIV></TD>
<TD>
<P style="LINE-HEIGHT: 150%; MARGIN: 0px 15px"><FONT face=����>��ʹ������Ϣ�ö�����Чʱ��Ϊ1���ö�/�죻ʹ�ø���Խ�࣬λ��Խ�ߡ�������ͬ��ʱ��Խ�磬λ��Խ�ߡ�</FONT></P></TD>
<TD>
<DIV align=center>
<SCRIPT src="user/helppz.asp?action=g_b"></SCRIPT>
</DIV></TD></TR>
<TR>
<TD>
<DIV align=center>
<P style="LINE-HEIGHT: 150%; MARGIN: 0px 15px"><FONT color=#008080>������ͼ����</FONT></FONT></P></DIV></TD>
<TD>
<P style="LINE-HEIGHT: 150%; MARGIN: 0px 15px"><FONT face=����>���Է�������Ϣ��ص�ͼƬ</FONT></P></TD>
<TD>
<DIV align=center>
<SCRIPT src="user/helppz.asp?action=g_c"></SCRIPT>
</DIV></TD></TR>
<TR>
<TD>
<DIV align=center>
<P style="LINE-HEIGHT: 150%; MARGIN: 0px 15px"><FONT color=#008080>ͨ����֤����</FONT></FONT></P></DIV></TD>
<TD>
<P style="LINE-HEIGHT: 150%; MARGIN: 0px 15px"><FONT face=����>�ɲ�ͨ��������ˣ�ֱ�ӷ���</FONT></P></TD>
<TD>
<DIV align=center>
<SCRIPT src="user/helppz.asp?action=g_d"></SCRIPT>
</DIV></TD></TR></TBODY></TABLE></DIV>
<TABLE style="BORDER-COLLAPSE: collapse" border=1 cellSpacing=0 borderColor=#fff2d7 cellPadding=0 width="98%">
<TBODY>
<TR>
<TD>
<P style="LINE-HEIGHT: 150%; MARGIN: 0px 15px"><B>�� <U>���֡�
<SCRIPT src="user/helppz.asp?action=rmb_mc"></SCRIPT>
��������ô����</U></B></P>
<P style="LINE-HEIGHT: 150%; MARGIN: 0px 15px"><br>�������������û��������û���¼��վ�Ĵ����������������ڱ�վ���ѵ�
<SCRIPT src="user/helppz.asp?action=rmb_mc"></SCRIPT>
�������Ƽ�ע�᱾վ��Ա�Ĺ��׻��֣���ϵͳ�Զ��Ƿ֣����壺 ע�����
<SCRIPT src="user/helppz.asp?action=jf_1"></SCRIPT>
����� ����һ����Ϣ��
<SCRIPT src="user/helppz.asp?action=jf_2"></SCRIPT>
����֣�ɾ��һ����Ϣ��
<SCRIPT src="user/helppz.asp?action=jf_2"></SCRIPT>
����֣�����Աɾ����Ϣ�ӱ��۷֡�
������ѶͶ��һƪ��
<SCRIPT src="user/helppz.asp?action=tgjf"></SCRIPT>
����֣�ɾ��������Ѷ��
<SCRIPT src="user/helppz.asp?action=tgjf"></SCRIPT>����֡�
 ��½һ�ε�
<SCRIPT src="user/helppz.asp?action=jf_3"></SCRIPT>
����� ����
<SCRIPT src="user/helppz.asp?action=jf_4"></SCRIPT>
��
<SCRIPT src="user/helppz.asp?action=rmb_mc"></SCRIPT>
��1����� ���Ƽ�ע�᱾վ��Աÿ�����<SCRIPT src="user/helppz.asp?action=jf_5"></SCRIPT>����֡�<br>����
<SCRIPT src="user/helppz.asp?action=rmb_mc"></SCRIPT>
�ɹ�����ߡ�
<SCRIPT src="user/helppz.asp?action=rmb_mc"></SCRIPT>
��ͨ��ע���Ա����������ת��������ҹ��򡣾��壺 ���ע�᱾վ��Ա�ɻ���
<SCRIPT src="user/helppz.asp?action=rmb_mc"></SCRIPT>

<SCRIPT src="user/helppz.asp?action=z_hb"></SCRIPT>
�� �������
<SCRIPT src="user/helppz.asp?action=jf_hb"></SCRIPT>
�����ת��Ϊһ��
<SCRIPT src="user/helppz.asp?action=rmb_mc"></SCRIPT>
 ���ɻ�1Ԫ����ҹ���
<SCRIPT src="user/helppz.asp?action=rmb_hb"></SCRIPT>
��
<SCRIPT src="user/helppz.asp?action=rmb_mc"></SCRIPT>
 ��<br>���õ����ܸı������ɫ����Ϣ�ö���������ͼ���������֤��������ͨ��ע���Ա��������
<SCRIPT src="user/helppz.asp?action=rmb_mc"></SCRIPT>
���򡣾��壺 ע����������ɫ����
<SCRIPT src="user/helppz.asp?action=z_a"></SCRIPT>
�� ע�������Ϣ�ö�����
<SCRIPT src="user/helppz.asp?action=z_b"></SCRIPT>
�� ע�����������ͼ����
<SCRIPT src="user/helppz.asp?action=z_c"></SCRIPT>
�� ע�����ͨ����֤����
<SCRIPT src="user/helppz.asp?action=z_d"></SCRIPT>
�� һ�������ɫ���߼۸�
<SCRIPT src="user/helppz.asp?action=g_a"></SCRIPT>
��
<SCRIPT src="user/helppz.asp?action=rmb_mc"></SCRIPT>
 һ����Ϣ�ö����߼۸�
<SCRIPT src="user/helppz.asp?action=g_b"></SCRIPT>
��
<SCRIPT src="user/helppz.asp?action=rmb_mc"></SCRIPT>
 һ��������ͼ���߼۸�
<SCRIPT src="user/helppz.asp?action=g_c"></SCRIPT>
��
<SCRIPT src="user/helppz.asp?action=rmb_mc"></SCRIPT>
 һ��ͨ����֤���߼۸�
<SCRIPT src="user/helppz.asp?action=g_d"></SCRIPT>
��
<SCRIPT src="user/helppz.asp?action=rmb_mc"></SCRIPT>��

<P></P>
<P style="LINE-HEIGHT: 150%; MARGIN: 0px 15px">������ʹ�ã�������Ϣʱ��ֱ��ʹ�ã�������߲��㣬�ɵ�½��վ�����û����ġ�
<SCRIPT src="user/helppz.asp?action=rmb_mc"></SCRIPT>
�һ����ߣ�
<SCRIPT src="user/helppz.asp?action=rmb_mc"></SCRIPT>
����ɻ���ת����������ҹ��򣩡��õ��߷����͹������ĸ�����Ϣ�� 
<P style="LINE-HEIGHT: 150%; MARGIN: 0px 15px"><B>�� <U>�����ʽ��ʻ�����ֵ</U></B></P>
<P style="LINE-HEIGHT: 150%; MARGIN: 0px 15px">��ÿ����Ա����һ��������ʽ��ʻ����ɶ��ʻ����г�ֵ���������ʻ��ʽ��ڱ�վ�������ѡ��ʻ��ʽ�ֻ���ڱ�վ���ѣ�����ȡ�ֽ���˿</P></TD></TR></TBODY></TABLE>

		<%CAse 2%>
		<div class="view">
			<h2 style="line-height:170%">��������</h2>
			<div class="txt"><%=Data(2)%></div>
		</div>
		<%Case 3%>
		<div class="view">
			<h2 style="line-height:170%">��������</h2>
			<div class="txt"><%=Data(3)%></div>
		</div>
		<%Case 4%>
		<div class="view">
			<h2 style="line-height:170%">������</h2>
			<div class="txt"><%=Data(4)%></div>
		</div>
		<%Case Else %>
		<div class="view">
			<h2 style="line-height:170%">��ϵ����</h2>
			<div class="txt"><%=Data(5)%></div>
		</div>
		<%End Select%>
	</div>
	<span class="cf"></span>
</div>

<!--#include file="footer.asp"-->
</body>
</html>
