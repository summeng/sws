<!--#include file="../../conn.asp"-->
<!--#include file="../../config.asp"-->
<%dim alipay_id,alipay_securityCode,partner,key,seller_email,notify_url,return_url,show_url,mainname,input_charset,sign_type,ddhm,hkuse,alimoney
Call OpenConn
set rs=server.CreateObject("adodb.recordset")
rs.open "select * from DNJY_pay",conn,1,1
alipay_id=rs("alipay_id")
alipay_securityCode=rs("alipay_securityCode")
partner=rs("partner")    
Call CloseRs(rs)
 %>
<%
	'���ܣ������ʻ��й���Ϣ������·������������ҳ�棩
	'�汾��3.1
	'���ڣ�2010-10-29
	'˵����
	'���´���ֻ��Ϊ�˷����̻����Զ��ṩ���������룬�̻����Ը����Լ���վ����Ҫ�����ռ����ĵ���д,����һ��Ҫʹ�øô��롣
	'�ô������ѧϰ���о�֧�����ӿ�ʹ�ã�ֻ���ṩһ���ο���

'��ʾ����λ�ȡ��ȫУ����ͺ��������ID
'1.����֧�����̻���������(b.alipay.com)��Ȼ��������ǩԼ֧�����˺ŵ�½.
'2.���ʡ��������񡱡������ؼ��������ĵ�����https://b.alipay.com/support/helperApply.htm?action=selfIntegration��
'3.�ڡ��������ɰ������У���������������(Partner ID)��ѯ��������ȫУ����(Key)��ѯ��

'��ȫУ����鿴ʱ������֧�������ҳ��ʻ�ɫ��������ô�죿
'���������
'1�������������ã������������������������
'2���������������ԣ����µ�¼��ѯ��

'�����������������������������������Ļ�����Ϣ������������������������������
'���������ID����2088��ͷ��16λ������
partner         = ""&partner&""

'��ȫ�����룬�����ֺ���ĸ��ɵ�32λ�ַ�
key   			= ""&alipay_securityCode&""

'ǩԼ֧�����˺Ż�����֧�����ʻ�
seller_email    = ""&alipay_id&""

'���׹����з�����֪ͨ��ҳ�� Ҫ�� http://��ʽ������·�����������?id=123�����Զ������
notify_url      = "http://"&web&"/"&strInstallDir&"alipay/js_asp_gb/notify_url.asp"

'��������ת��ҳ�� Ҫ�� http://��ʽ������·�����������?id=123�����Զ������
return_url      = "http://"&web&"/"&strInstallDir&"alipay/js_asp_gb/return_url.asp"

'��վ��Ʒ��չʾ��ַ���������?id=123�����Զ������
show_url        = "http://"&web&"/"&strInstallDir&""

'�տ���ƣ��磺��˾���ơ���վ���ơ��տ���������
mainname		= ""&webname&""

'�����������������������������������Ļ�����Ϣ������������������������������



'�ַ������ʽ Ŀǰ֧�� gbk �� utf-8
input_charset	= "gbk"

'ǩ����ʽ �����޸�
sign_type       = "MD5"


%>