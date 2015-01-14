<%
'=====================================
'天天供求信息网站管理系统
'天天网络科技工作室出品
'客服中心http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>
<%'SQL通用防注入系统V3.0
'--------定义部份------------------
Dim Fy_Post,Fy_Get,Fy_In,Fy_Inf,Fy_Xh,Fy_db,Fy_dbstr,Kill_IP,WriteSql
'自定义需要过滤的字串,用 "|" 分隔
Fy_In = "'|;|and|(|)|exec|insert|select|delete|update|count|*|%|chr|mid|master|truncate|char|declare|┼|∨|┩|≡"
Kill_IP=False
WriteSql=True			
'----------------------------------


Fy_Inf = split(Fy_In,"|")
'--------POST部份------------------

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
				Response.Write "<Script Language=JavaScript>alert('本站SQL防注入系统提示你↓\n\n请不要在参数中包含非法字符尝试注入！\n\n检查是否有;%=*or/@""--$等字符！\n\n你已经违规操作"&session("sql_err")&"次！');</Script>"'ip126
				Response.Write "非法操作！系统做了如下记录↓<br>"
				Response.Write "操作ＩＰ："&Request.ServerVariables("REMOTE_ADDR")&"<br>"
				Response.Write "操作时间："&Now&"<br>"
				Response.Write "操作页面："&Request.ServerVariables("URL")&"<br>"
				Response.Write "提交方式：ＰＯＳＴ<br>"
				Response.Write "提交参数："&Fy_Post&"<br>"
				Response.Write "提交数据："&Request.Form(Fy_Post)
				Response.End
			End If
		Next
	Next
End If

'----------------------------------

'--------GET部份-------------------
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
				Response.Write "<Script Language=JavaScript>alert('本站SQL防注入系统提示你↓\n\n请不要在参数中包含非法字符尝试注入！\n\n检查是否有;%=*or/@""--$等字符！\n\n你已经违规操作"&session("sql_err")&"次！');</Script>"'ip126
				Response.Write "非法操作！系统做了如下记录↓<br>"
				Response.Write "操作ＩＰ："&Request.ServerVariables("REMOTE_ADDR")&"<br>"
				Response.Write "操作时间："&Now&"<br>"
				Response.Write "操作页面："&Request.ServerVariables("URL")&"<br>"
				Response.Write "提交方式：ＧＥＴ<br>"
				Response.Write "提交参数："&Fy_Get&"<br>"
				Response.Write "提交数据："&Request.QueryString(Fy_Get)
				
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
	Response.write "<Script Language=JavaScript>alert('SQL通用防注入系统提示你↓\n\n你的Ip已经被本系统自动锁定！\n\n如想访问本站请和管理员联系！');</Script>"
  Sqlin_IP.Close:Set Sqlin_IP=Nothing
  rsKill_IP.Close: Set rsKill_IP = Nothing
  Response.End
	End If
    Sqlin_IP.Close:Set Sqlin_IP=Nothing
	rsKill_IP.Close: Set rsKill_IP = Nothing
End If
%>