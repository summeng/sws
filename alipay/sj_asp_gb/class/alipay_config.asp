<!--#include file="../../../conn.asp"-->
<!--#include file="../../../config.asp"-->
<%dim alipay_id,alipay_securityCode,partner,key,seller_email,notify_url,return_url,show_url,mainname,input_charset,sign_type,ddhm,hkuse,alimoney
Dim payment_type,out_trade_no,subject,price,quantity,logistics_fee,logistics_type,logistics_payment,body,receive_name,receive_zip,receive_phone,receive_mobile,sParaTemp,sHtml,receive_address,objSubmit,i,pos,nLen,itemName,itemValue,sPara,minmax,minmaxSlot,j,mark,temp,sParaSort,nCount,prestr,iPos,sItemName,sItemValue
Dim sVerifyResult,trade_no,trade_status,HTTPS_VERIFY_URL,objNotify,varItem,order_no,total_fee

Call OpenConn
set rs=server.CreateObject("adodb.recordset")
rs.open "select * from DNJY_pay",conn,1,1
alipay_id=rs("alipay_id")
alipay_securityCode=rs("alipay_securityCode")
partner=rs("partner")    
rs.close
set rs=Nothing
 %>
<%
' �����ļ�
' ���ܣ������ʻ��й���Ϣ������·��
' �汾��3.3
' ���ڣ�2012-07-13
' ˵����
' ���´���ֻ��Ϊ�˷����̻����Զ��ṩ���������룬�̻����Ը����Լ���վ����Ҫ�����ռ����ĵ���д,����һ��Ҫʹ�øô��롣
' �ô������ѧϰ���о�֧�����ӿ�ʹ�ã�ֻ���ṩһ���ο���

' ��ʾ����λ�ȡ��ȫУ����ͺ��������ID
' 1.������ǩԼ֧�����˺ŵ�¼֧������վ(www.alipay.com)
' 2.������̼ҷ���(https://b.alipay.com/order/myOrder.htm)
' 3.�������ѯ���������(PID)��������ѯ��ȫУ����(Key)��

' ��ȫУ����鿴ʱ������֧�������ҳ��ʻ�ɫ��������ô�죿
' ���������
' 1�������������ã������������������������
' 2���������������ԣ����µ�¼��ѯ��

'�����������������������������������Ļ�����Ϣ������������������������������

'���������ID����2088��ͷ��16λ������
partner= ""&partner&""

'��ȫ�����룬�����ֺ���ĸ��ɵ�32λ�ַ�
key= ""&alipay_securityCode&""


'�����������������������������������Ļ�����Ϣ������������������������������



'�ַ������ʽ Ŀǰ֧�� gbk �� utf-8
input_charset= "gbk"

'ǩ����ʽ �����޸�
sign_type= "MD5"
%>