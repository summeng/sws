<!--#include file=cookies.asp-->
<!--#include file="../config.asp"-->
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

<title>����˵�</title>
<script type="text/javascript" src="inc/js_left.js"></script>

<base target="mainFrame" />
<link rel="stylesheet" href="inc/style_left.css" type="text/css" />
</head>
<body>
<div id="body">

<div class="d">�������</div>
<div class="e"><a href="right.asp" TARGET=right>������ҳ</a> �� <a HREF="exit.asp" TARGET="_parent">��ȫ�˳�</a></div>

<div class="a" id="AA0" onclick="TAB(0)" >������Ϣ����</div>
<div class="c" id="BB0" style="display:block;">
<p><A href="infolist.asp" TARGET=right>��Ϣ����</A> �� <A href="info_add.asp" TARGET=right>��������</A></p>
<p><A href="xxhflist.asp" TARGET=right>�ظ���Ϣ</A> �� <A href="Bad_info_list.asp" TARGET=right>�ٱ���Ϣ</A></p>
<p><a href="siterss.asp" TARGET=right>��վRSS</a> �� <a href="Substationrss.asp" TARGET=right>��վRSS</a></p>
<p><a href="Substationrss_Batches.asp" TARGET=right>����RSS</a> �� <a href="info_html.asp" TARGET=right>html����</a></p>
<p><a href="siteinfo.asp" TARGET=right>�˳̵���</a> �� <a href="sitenew.asp" TARGET=right>JS����</a></p>
</div>

<div class="b" id="AA1" onclick="TAB(1)">��Ա���Ϲ���</div>
<div class="c" id="BB1" style="display:none;">
<p><A href="users_List.asp" TARGET=right>��Ա����</A> �� <A href="add_user.asp" TARGET=right>��ӻ�Ա</A></p>
<p><A href="admin_order.asp" TARGET=right>��������</A> �� <A href=" admin_Amount.asp" TARGET=right>������־</A></p>
<p><A href="xnp.asp" TARGET=right>��<%=rmb_mc%></A> �� <a href="sitehy.asp" TARGET=right>�������</a></p>
<p><A href="UserNews_ModSelClass.asp" TARGET=right>��Ա��Ѷ</A> �� <A href="usernews_ClassManage.asp" TARGET=right>��Ŀ����</A></p>
<p><A href="sms_loglist.asp?isse=0" TARGET=right><font color="#4144C1">��֤��¼</font></A> <%If mailqr=1 then%>�� <A href="user_Temporary.asp" TARGET=right>��ʱ��Ա</A><%End if%></p>
</div>

<div class="b" id="AA2" onclick="TAB(2)">������Ϣ����</div>
<div class="c" id="BB2" style="display:none;">
<p><A href="book.asp" TARGET=right>���Թ���</A> �� </p>
<!--p><a href="#" ONCLICK="window.open('e_mail.asp','Sample','toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=yes,resizable=no,copyhistory=no,width=730,height=700,left=300,top=20')" target=_self>������־</font></A> �� <a href="#" ONCLICK="window.open('sendmail_user.asp','Sample','toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=yes,resizable=no,copyhistory=no,width=730,height=700,left=300,top=20')" target=_self>��־����</font></A></p-->
<p><a href="#" ONCLICK="window.open('AdminSendEmail.asp','Sample','toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=yes,resizable=no,copyhistory=no,width=730,height=700,left=300,top=20')" target=_self>�ʼ�Ⱥ��</font></A> �� <A href="mail_loglist.asp?isse=0" TARGET=right><font color="#4144C1">�ʼ���־</font></A></p>
<p><A href="Message1.asp" TARGET=right><font color="#1DA2DB">���ն���</font></A> �� <A href="Message2.asp" TARGET=right><font color="#1DA2DB">׫д����</font></A></p>
<p><A href="Message3.asp" TARGET=right><font color="#1DA2DB">���Ź���</font></A> �� <A href="Message4.asp" TARGET=right><font color="#1DA2DB">ɾ������</font></A></p>
</div>

<div class="b" id="AA3" onclick="TAB(3)">������Ѷ����</div>
<div class="c" id="BB3" style="display:none;">
<p><A href="admin_article_add.asp"  TARGET=right>���������Ѷ</A></p>
<p><A href="admin_article_list.asp"  TARGET=right>����������Ѷ</A></p>
<p><A href="admin_article_comment.asp"  TARGET=right>�������۹���</A></p>
<p><A href="admin_article_Class1.asp"  TARGET=right>������Ŀ����</A></p>
</div>

<div class="b" id="AA4" onclick="TAB(4)">������־����</div>
<div class="c" id="BB4" style="display:none;">
<p><A href="../magazine/magazine_issue.asp"  TARGET=right>��־�ںŹ���</A></p>
<p><A href="../magazine/magazine_edition.asp"  TARGET=right>��־�������</A></p>
</div>

<div class="b" id="AA5" onclick="TAB(5)">��Ϣ�ɼ�����</div>
<div class="c" id="BB5" style="display:none;">
<p><A href="collect_management.asp"  TARGET=right>��Ŀ����</A></p>
<p><A href="collect_start.asp"  TARGET=right>���ݲɼ�</A></p>
<p><A href="collect_Settings.asp"  TARGET=right>��������</A></p>
<p><A href="collect_record.asp"  TARGET=right>��ʷ��¼</A></p>
<p><A href="collect_laoy.asp?Action=import"  TARGET=right>�������</A></p>
<p><A href="collect_laoy.asp"  TARGET=right>��������</A></p>
</div>

<div class="b" id="AA6" onclick="TAB(6)">���Ųɼ�����</div>
<div class="c" id="BB6" style="display:none;">
<p><A href="cj_management.asp"  TARGET=right>��Ŀ����</A></p>
<p><A href="cj_start.asp"  TARGET=right>���ݲɼ�</A></p>
<p><A href="cj_Settings.asp"  TARGET=right>��������</A></p>
<p><A href="cj_record.asp"  TARGET=right>��ʷ��¼</A></p>
<p><A href="cj_laoy.asp?Action=import"  TARGET=right>�������</A></p>
<p><A href="cj_laoy.asp"  TARGET=right>��������</A></p>
</div>

<div class="b" id="AA7" onclick="TAB(7)">��վϵͳ����</div>
<div class="c" id="BB7" style="display:none;">
<p><A href="admin_parameters.asp?id=1" TARGET=right>��������</A> �� <A href="add_administrator.asp" TARGET=right>����ԱȨ��</A></p>
<p><a href="pay1.asp" TARGET=right>֧��ƽ̨</a> �� <a href="pay2.asp" TARGET=right>�����ʺ�</a></p>
<p><a href="api.asp" TARGET=right>API����</a> �� <A href="Mail_system.asp?id=1" TARGET=right>�ʼ�ϵͳ</A> </p>
<p><a href="type_Level1.asp"  TARGET=right>��Ϣ����</a> �� <a href="city_Level1.asp"  TARGET=right>��������</a></p>
<p><a href="store_Level1.asp"  TARGET=right>�������</a> �� <a href="admin_help.asp"  TARGET=right>��������</A></p>
<p><a href="admin_announce.asp"  TARGET=right>��վ����</A> �� <a href="#" ONCLICK="window.open('notice.asp','Sample','toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=yes,resizable=no,copyhistory=no,width=730,height=500,left=300,top=20')" target=_self>֪ͨ��վ</a></p>
</div>

<div class="b" id="AA8" onclick="TAB(8)">��վģ�����</div>
<div class="c" id="BB8" style="display:none;">
<p><a href="admin_Template_Management.asp" TARGET=right>ģ������</a> �� <a href="Admin_Template_Main.asp?Action=Main" TARGET=right>ģ�����վ</a></p>
<p><a href="Admin_Template.asp?Action=Import" TARGET=right>ģ�嵼��</a> �� <a href="Admin_Template.asp?Action=Export" TARGET=right>ģ�嵼��</a></p>
<!--<p><a href="sitemap.asp" TARGET=right>���� sitemap</a></p>-->
</div>

<div class="b" id="AA9" onclick="TAB(9)">�������ܹ���</div>
<div class="c" id="BB9" style="display:none;">
<p><a href="bmcs.asp" TARGET=right>������Ϣ</a> �� <a href="admin_114.asp" TARGET=right>����114</a></p>
<p><a href="admin_link.asp" TARGET=right>��������</a> �� </p>
</div>

<div class="b" id="AA10" onclick="TAB(10)">��׼������</div>
<div class="c" id="BB10" style="display:none;">
<p><a href="../ads/add.asp" TARGET=right>��ӹ��</a> �� <a href="../ads/list.asp?all=ok" TARGET=right>����б�</a></p>
<p><a href="../ads/list.asp?action=stop&all=ok" TARGET=right>���ڹ��</a> �� <a href="../ad.asp" TARGET=right>Ĭ�Ϲ��</a></p>
<p><a href="../ads/code.asp" TARGET=right>����ʾ��</a> �� <a href="../ads/Leading_in_out.asp" TARGET=right>���뵼��</a></p>
</div>

<div class="b" id="AA11" onclick="TAB(11)">���������</div>
<div class="c" id="BB11" style="display:none;">
<p><a href="adv_slide.asp" TARGET=right><font color="#ff0000">�õƹ��A</font></a> �� <a href="adv.asp" TARGET=right><font color="#ff0000">�õƹ��B</font></a></p>
<p> <a href="adfp.asp" TARGET=right>���ƹ��</a>&nbsp;&nbsp;&nbsp;�� <a href="admin_flv.asp" TARGET=right>FLV��Ƶ</a></p>
</div>

<div class="b" id="AA12" onclick="TAB(12)">������������</div>
<div class="c" id="BB12" style="display:none;">
<p><A href="admin_ip.asp" TARGET=right>�ַ�IP���ƣ�<A href="admin_check.asp" TARGET=right>�ʴ���֤</A></p>
<!--p><a href="admin_backupdata.asp" TARGET=right>��������</a>��<a href="admin_Compressdata.asp" TARGET=right>ѹ������</a></p-->
<p><a href="delcache.asp" TARGET=right>������</a>��<a href="../SpaceSize.asp" TARGET=right>�ռ�ռ��</a></p>
<p><a href="delpic.asp" TARGET=right>���������ϴ�ͼƬ���ļ�</a></p>
<p><a href="admin_server.asp" TARGET=right>������̽��</a>��<a href="admin_data.asp" TARGET=right>���ݿ����</a></p>
<!--<p><a href="fileUpdate.asp" TARGET=right>�����������ļ���</a></p>-->
<%If sys=True then%><p><a href="admin_server.asp" TARGET=right>������̽��</a>��<a href="sqladmin.asp" TARGET=right>ע���¼</a></p><%End if%>
</div>

<div class="f">ϵͳ��Ȩ��Ϣ</div>
<div class="c">
<p><a href="http://www.ip126.com/" target="_blank">�ͷ�����</a> �� <a href="http://www.ip126.com/bbs/index.asp?boardid=49" target="_blank">�ͷ���̳</a></p>
<p><a href="http://www.ip126.com" target="_blank">��������Ƽ�������</a></p>
</div>

</div>
</body>
</html>
