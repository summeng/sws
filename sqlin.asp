<%
'=====================================
'���칩����Ϣ��վ����ϵͳ
'��������Ƽ������ҳ�Ʒ
'�ͷ�����http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>
<%'SQLͨ�÷�ע��ϵͳV3.0
'--------���岿��------------------
Dim Fy_Post,Fy_Get,Fy_In,Fy_Inf,Fy_Xh,Fy_db,Fy_dbstr,Kill_IP,WriteSql
'�Զ�����Ҫ���˵��ִ�,�� "|" �ָ�
Fy_In = "'|;|and|(|)|exec|insert|select|delete|update|count|*|%|chr|mid|master|truncate|char|declare|��|��|��|��"
Kill_IP=False
WriteSql=True			
'----------------------------------


Fy_Inf = split(Fy_In,"|")
'--------POST����------------------

If Request.Form<>"" Then
	For Each Fy_Post In Request.Form
		For Fy_Xh=0 To Ubound(Fy_Inf)
			If Instr(LCase(Request.Form(Fy_Post)),Fy_Inf(Fy_Xh))<>0 Then
				If WriteSql=True And sys=True Then
					conn.Execute("insert into DNJY_sqlin(Sqlin_IP,SqlIn_Web,SqlIn_TIME,SqlIn_FS,SqlIn_CS,SqlIn_SJ,Kill_ip) values('"&Request.ServerVariables("REMOTE_ADDR")&"','"&Request.ServerVariables("URL")&"','"&Now()&"','POST','"&Fy_Post&"','"&replace(Request.Form(Fy_Post),"'","''")&"',"&Kill_IP&")")
					conn.close
					Set conn = Nothing
				End If
				session("sql_err")=session("sql_err")+1'ip126
				Response.Write "<Script Language=JavaScript>alert('��վSQL��ע��ϵͳ��ʾ���\n\n�벻Ҫ�ڲ����а����Ƿ��ַ�����ע�룡\n\n����Ƿ���;%=*or/@""--$���ַ���\n\n���Ѿ�Υ�����"&session("sql_err")&"�Σ�');</Script>"'ip126
				Response.Write "�Ƿ�������ϵͳ�������¼�¼��<br>"
				Response.Write "�����ɣУ�"&Request.ServerVariables("REMOTE_ADDR")&"<br>"
				Response.Write "����ʱ�䣺"&Now&"<br>"
				Response.Write "����ҳ�棺"&Request.ServerVariables("URL")&"<br>"
				Response.Write "�ύ��ʽ���Уϣӣ�<br>"
				Response.Write "�ύ������"&Fy_Post&"<br>"
				Response.Write "�ύ���ݣ�"&Request.Form(Fy_Post)
				Response.End
			End If
		Next
	Next
End If

'----------------------------------

'--------GET����-------------------
If Request.QueryString<>"" Then
	For Each Fy_Get In Request.QueryString
		For Fy_Xh=0 To Ubound(Fy_Inf)
			If Instr(LCase(Request.QueryString(Fy_Get)),Fy_Inf(Fy_Xh))<>0 Then
				If WriteSql=True And sys=True Then
					conn.Execute("insert into DNJY_sqlin(Sqlin_IP,SqlIn_Web,SqlIn_TIME,SqlIn_FS,SqlIn_CS,SqlIn_SJ,Kill_ip) values('"&Request.ServerVariables("REMOTE_ADDR")&"','"&Request.ServerVariables("URL")&"','"&Now()&"','GET','"&Fy_Get&"','"&replace(Request.QueryString(Fy_Get),"'","''")&"',"&Kill_IP&")")
				conn.close
				Set conn = Nothing
				End If
				session("sql_err")=session("sql_err")+1'ip126
				Response.Write "<Script Language=JavaScript>alert('��վSQL��ע��ϵͳ��ʾ���\n\n�벻Ҫ�ڲ����а����Ƿ��ַ�����ע�룡\n\n����Ƿ���;%=*or/@""--$���ַ���\n\n���Ѿ�Υ�����"&session("sql_err")&"�Σ�');</Script>"'ip126
				Response.Write "�Ƿ�������ϵͳ�������¼�¼��<br>"
				Response.Write "�����ɣУ�"&Request.ServerVariables("REMOTE_ADDR")&"<br>"
				Response.Write "����ʱ�䣺"&Now&"<br>"
				Response.Write "����ҳ�棺"&Request.ServerVariables("URL")&"<br>"
				Response.Write "�ύ��ʽ���ǣţ�<br>"
				Response.Write "�ύ������"&Fy_Get&"<br>"
				Response.Write "�ύ���ݣ�"&Request.QueryString(Fy_Get)
				
				Response.End
			End If
		Next
	Next
End If

If Kill_IP=True Then
	Dim Sqlin_IP,rsKill_IP
	Sqlin_IP=Request.ServerVariables("REMOTE_ADDR")
	Sql="select Sqlin_IP from DNJY_sqlin where Kill_ip="&DN_true&" and Sqlin_IP='"&Sqlin_IP&"'"
	Set rsKill_IP=conn.execute(sql)
	If Not(rsKill_IP.eof or rsKill_IP.bof) Then
	Response.write "<Script Language=JavaScript>alert('SQLͨ�÷�ע��ϵͳ��ʾ���\n\n���Ip�Ѿ�����ϵͳ�Զ�������\n\n������ʱ�վ��͹���Ա��ϵ��');</Script>"
  Sqlin_IP.Close:Set Sqlin_IP=Nothing
  rsKill_IP.Close: Set rsKill_IP = Nothing
  Response.End
	End If
    Sqlin_IP.Close:Set Sqlin_IP=Nothing
	rsKill_IP.Close: Set rsKill_IP = Nothing
End If
%>