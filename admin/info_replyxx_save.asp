<!--#include file="../conn.asp"-->
<!--#include file="../SqlIn.Asp"-->
<!--#include file="../Include/err.asp"-->
<!--#include file=cookies.asp-->
<!--#include file="../config.asp"-->
<!--#include file="../Include/head.asp"-->
<!--#include file="../Include/Function.asp"-->
<!--#include file="../Include/label.asp"-->
<%
'=====================================
'���칩����Ϣ��վ����ϵͳ
'��������Ƽ������ҳ�Ʒ
'�ͷ�����http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>
<%dim sdays,hfje,urlpage,hfxxzg
hfxxzg=request("hfxxzg")
If hfxxzg=1 Then
Call OpenConn
conn.execute "UPDATE DNJY_info SET hfje=0 WHERE  hfxx=1 and "&DN_Now&">=dqsj" 'ͬʱ���û������ʻ����
conn.execute "UPDATE DNJY_info SET hfxx=0 WHERE  hfxx=1 and "&DN_Now&">=dqsj" 'ͬʱ���û������ʻ����
End If
Call CloseDB(conn)
response.write "<p align=""center"">�����ɹ���</p>"
response.end
%>
