<!--#include file="../../conn.asp"-->
<!--#include file="../../config.asp"--> 
<%
dim tenpay_id,tenpay_key,tenpay_name,ddhm,hkuse,alimoney,spname,partner,key,return_url,notify_url
Dim strNow,i,j,tmp,md5str,fs,randNumber,dingdan,product_name,order_price,currTime,strTime,strReq,total_fee,remarkexplain,out_trade_no,trade_mode,bank_type_value,ts

Call OpenConn
set rs=server.CreateObject("adodb.recordset")
rs.open "select * from DNJY_pay",conn,1,1
tenpay_id=rs("tenpay_id")
tenpay_key=rs("tenpay_key")
tenpay_name=rs("tenpay_name")    
Call CloseRs(rs)
 spname=tenpay_name
 partner=tenpay_id                                              '�Ƹ�ͨ�̻���
 key=tenpay_key                    								 '�Ƹ�ͨ��Կ
 return_url="http://"&web&"/"&strInstallDir&"tenpay/b2c/return_url.asp"         	         '��ʾ֧�����ҳ��,*�滻��return_url.asp����·��
 notify_url="http://"&web&"/"&strInstallDir&"tenpay/b2c/notify_url.asp"              		 '֧����ɺ�Ļص�����ҳ��,*�滻��notify_url.asp����·��
%>