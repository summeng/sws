<!--#include file="alipay_config.asp"-->
<!--#include file="class/alipay_service.asp"-->
<%
	'���ܣ�������Ʒ�й���Ϣ��ȷ�϶���֧�������߹������ҳ��
	'��ϸ����ҳ���ǽӿ����ҳ�棬����֧��ʱ��URL
	'�汾��3.1
	'���ڣ�2010-10-29
	'˵����
	'���´���ֻ��Ϊ�˷����̻����Զ��ṩ���������룬�̻����Ը����Լ���վ����Ҫ�����ռ����ĵ���д,����һ��Ҫʹ�øô��롣
	'�ô������ѧϰ���о�֧�����ӿ�ʹ�ã�ֻ���ṩһ���ο���
	
'''''''''''''''''ע��'''''''''''''''''''''''''
'������ڽӿڼ��ɹ������������⣬
'�����Ե��̻��������ģ�https://b.alipay.com/support/helperApply.htm?action=consultationApply�����ύ���뼯��Э�������ǻ���רҵ�ļ�������ʦ������ϵ��Э�������
'��Ҳ���Ե�֧������̳��http://club.alipay.com/read-htm-tid-8681712.html��Ѱ����ؽ������
'�������ʹ����չ���������չ���ܲ�������ֵ��
''''''''''''''''''''''''''''''''''''''''''''''
%>



<%Dim out_trade_no,subject,body,total_fee,pay_mode,paymethod,defaultbank,exter_invoke_ip,anti_phishing_key,extra_common_param,buyer_email,royalty_type,royalty_parameters,para,sHtmlText
'''���²�������Ҫͨ���µ�ʱ�Ķ������ݴ���������'''
'�������
sTime=now()
'out_trade_no =year(sTime)&month(sTime)&day(sTime)&hour(sTime)&minute(sTime)&second(sTime)'�������վ����ϵͳ�е�Ψһ������ƥ��
out_trade_no =request.Form("aliorder")'�̼���վ����ϵͳ�е�Ψһ������
subject      = request.Form("aliorder")		'�������ƣ���ʾ��֧��������̨��ġ���Ʒ���ơ����ʾ��֧�����Ľ��׹���ġ���Ʒ���ơ����б��
body         = request.Form("alibody")		'����������������ϸ��������ע����ʾ��֧��������̨��ġ���Ʒ��������
total_fee    = request.Form("alimoney")		'�����ܽ���ʾ��֧��������̨��ġ�Ӧ���ܶ��

'��չ���ܲ�������Ĭ��֧����ʽ
pay_mode	 = request.Form("pay_bank")
if pay_mode = "directPay" then
	paymethod    = "directPay"	'Ĭ��֧����ʽ���ĸ�ֵ��ѡ��bankPay(����); cartoon(��ͨ); directPay(���); CASH(����֧��)
	defaultbank	 = ""
else
	paymethod    = "bankPay"	'Ĭ��֧����ʽ���ĸ�ֵ��ѡ��bankPay(����); cartoon(��ͨ); directPay(���); CASH(����֧��)
	defaultbank  = pay_mode		'Ĭ���������ţ������б��http://club.alipay.com/read.php?tid=8681379
end if

'��չ���ܲ�������������
'������ѡ���Ƿ��������㹦��
'exter_invoke_ip��anti_phishing_keyһ�������ù�����ô���Ǿͻ��Ϊ�������
'��Ҫʹ�÷����㹦�ܣ�����ʹ��POST��ʽ��������
exter_invoke_ip   = ""			'��ȡ�ͻ��˵�IP��ַ�����飺��д��ȡ�ͻ���IP��ַ�ĳ���
anti_phishing_key = ""			'������ʱ���
'�磺
'exter_invoke_ip = "202.1.1.1"
'anti_phishing_key = query_timestamp(partner)		'��ȡ������ʱ�������

'��չ���ܲ�����������
extra_common_param = ""			'�Զ���������ɴ���κ����ݣ���=��&�������ַ��⣩��������ʾ��ҳ����
buyer_email		   = ""			'Ĭ�����֧�����˺�

'��չ���ܲ�����������(��Ҫʹ�ã��밴��ע��Ҫ��ĸ�ʽ��ֵ)
royalty_type		= ""		'������ͣ���ֵΪ�̶�ֵ��10������Ҫ�޸�
royalty_parameters	= ""
'�����Ϣ��������Ҫ����̻���վ���������̬��ȡÿ�ʽ��׵ĸ������տ��˺š��������������˵�������ֻ������10��
'����������ܺ���С�ڵ���total_fee
'�����Ϣ����ʽΪ���տEmail_1^���1^��ע1|�տEmail_2^���2^��ע2
'�磺
'royalty_type = "10"
'royalty_parameters	= "111@126.com^0.01^����עһ|222@126.com^0.01^����ע��"

''''''''''''''''''''''''''''''''''''''''''''''''''''
'����Ҫ����Ĳ������飬����Ķ�
para = Array("service=create_direct_pay_by_user","payment_type=1","partner="&partner,"seller_email="&seller_email,"return_url="&return_url,"notify_url="&notify_url,"_input_charset="&input_charset,"show_url="&show_url,"out_trade_no="&out_trade_no,"subject="&subject,"body="&body,"total_fee="&total_fee,"paymethod="&paymethod,"defaultbank="&defaultbank,"anti_phishing_key="&anti_phishing_key,"exter_invoke_ip="&exter_invoke_ip,"extra_common_param="&extra_common_param,"buyer_email="&buyer_email,"royalty_type="&royalty_type,"royalty_parameters="&royalty_parameters)

'����������
alipay_service(para)
sHtmlText = build_form()
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>֧������ʱ֧��</title>
<style type="text/css">
.font_content{
	font-family:"����";
	font-size:14px;
	color:#FF6600;
}
.font_title{
	font-family:"����";
	font-size:16px;
	color:#FF0000;
	font-weight:bold;
}
table{
	border: 1px solid #CCCCCC;
}
</style>
</head>
<body>
        <table align="center" width="350" cellpadding="5" cellspacing="0">
            <tr>
                <td align="center" class="font_title" colspan="2">����ȷ��</td>
            </tr>
            <tr>
                <td class="font_content" align="right">�����ţ�</td>
                <td class="font_content" align="left"><%=out_trade_no%></td>
            </tr>
            <tr>
                <td class="font_content" align="right">�����ܽ�</td>
                <td class="font_content" align="left"><%=total_fee%></td>
            </tr>
            <tr>
                <td align="center" colspan="2"><%=sHtmlText%></td>
            </tr>
        </table>
</body>
</html>
<%Call CloseDB(conn)%>