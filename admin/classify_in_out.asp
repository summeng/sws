<!--#include file="../conn.asp"-->
<!--#include file="../config.asp"-->
<!--#include file="../SqlIn.Asp"-->
<!--#include file=cookies.asp-->

<%
'=====================================
'���칩����Ϣ��վ����ϵͳ
'��������Ƽ������ҳ�Ʒ
'�ͷ�����http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>
<%call checkmanage("02")%>
<HTML>
<HEAD>
<TITLE>���󣭵�������Ϣ���������</TITLE>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312"><style type="text/css">
<!--
body {
	background-image: url(../images/dj_bg.gif);
}
-->
</style>
<link href="../css/style.css" rel="stylesheet" type="text/css">
<style type="text/css">
<!--
.style4 {color: #FFFFFF;
	font-weight: bold;
}
-->
</style>
<script>
function loadThreadFollow(t_id,b_id){
	var targetImg =eval("document.all.followImg" + t_id);
	var targetDiv =eval("document.all.follow" + t_id);
	
	if ("object"==typeof(targetImg)){
		if (targetDiv.style.display!='block'){
			targetDiv.style.display="block";
			
		}else{
			targetDiv.style.display="none";
		}
	}
}
function CheckAll(form)
  {
  for (var i=0;i<form.elements.length;i++)
    {
    var e = form.elements[i];
    if (e.name != 'chkall' )
       e.checked = form.chkall.checked;
    }
  }
//-->
</script>
</head>
<BODY>
<BODY>
<table align="center" width="100%" border="1" cellspacing="0" cellpadding="4" bordercolor="#ADAED6" style="border-collapse: collapse">
<TR> 
    <TD height="5" colspan="12" bgcolor="#799AE1" align="center"><FONT color="#FFFFFF" style="font-size:14px">Ĭ�Ϸ��ർ�����</FONT></TD>
  </TR>
<table>
<%Dim DBmb,connstrmb,connmb,action,id,Importpath,Importto,str2,rsi,rsmb,k
action=request.form("action")
Call OpenConn
DBmb="../temp/Classify.mdb"  '���޸�
connstrmb="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(""&DBmb&"")
  On Error Resume Next
	Set connmb = Server.CreateObject("ADODB.Connection")
	connmb.open connstrmb
	If Err Then
		err.Clear
		Set Connmb = Nothing
		Response.Write "���ݿ����ӳ������������ִ���"'ע�ͣ���Ҫ���⼸���ַ����Ӣ�ġ�
		Response.End
End If

'=============================����������====================================
if request("action")="dqfl" Then%>
<div align="center">
  <center>
<table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#ADAED6" width="100%" >
<TR> 
    <TD  bgcolor="#799AE1" align="center">  <form action="?Action=dqflto" name="form1" method="post">
    <tr valign="middle"> 
      <td bgcolor="#FFFFFF"><br>
	 	  &nbsp;&nbsp;Ĭ�ϵ�����ർ����Դ���ݿ⣺
              <input type=text size=50 name=Importpath value="../temp/Classify.mdb" class="smallInput" readonly>
              &nbsp;&nbsp;<BR>
  					&nbsp;&nbsp;���ݿ����Ƽ�·��ӦΪ��temp/Classify.mdb<BR>
						&nbsp;&nbsp;
              <input type=submit value="��һ��" class="button">
                </td>
	</tr>
	<input type="hidden" name="action" value="Restore">
  </form></TD>
  </TR>
<table>
</div>

<%End if
if request("action")="dqflto" Then
set rsmb=server.createobject("adodb.recordset")
sql = "select * from DNJY_hy_type"
rsmb.open sql,connmb,1,1
if rsmb.eof then
response.write "<table>��û��Ĭ�ϵ�����࣬�޷����룡</table>"
rsmb.close
set rsmb=Nothing
response.end
end If
Set Rsi=server.createobject("adodb.recordset")
	Sql="Select * From DNJY_hy_type"
	Rsi.open Sql,conn,1,3
	IF Not Rsi.Eof Then
	response.write "<script language='javascript'>	alert('��ǰϵͳ�Ѵ��ڷ���,Ҫȫ��ɾ������ܵ��룡');location.href='javascript:history.go(-1)';</script>"		
	rsi.close:Set rsi=Nothing
    rsmb.close:set rsmb=Nothing
	Response.End
	Else	
do while not rsmb.eof	

	Rsi.Addnew
	Rsi("ID")=Rsmb("ID")
	Rsi("twoid")=Rsmb("twoid")
	Rsi("threeid")=Rsmb("threeid")
	Rsi("indexid")=Rsmb("indexid")
	Rsi("name")=Rsmb("name")
	Rsi("titlecolor")=Rsmb("titlecolor")
	Rsi("indexshow")=Rsmb("indexshow")
	Rsi.update	

rsmb.movenext
Loop
End if
Call CloseRs(rsi)
rsmb.close
set rsmb=Nothing
response.write "<script language=JavaScript>" & chr(13) & "alert('����ٷ�Ĭ�ϵ������ɹ���');" &"window.location='store_Level1.asp'" & "</script>"
End if
'=============================����������end====================================


'=============================������Ϣ����====================================
if request("action")="xxfl" Then%>
<div align="center">
  <center>
<table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#ADAED6" width="100%" >
<TR> 
    <TD  bgcolor="#799AE1" align="center">  <form action="?Action=xxflto" name="form1" method="post">
    <tr valign="middle"> 
      <td bgcolor="#FFFFFF"><br>
	 	  &nbsp;&nbsp;Ĭ����Ϣ���ർ����Դ���ݿ⣺
              <input type=text size=50 name=Importpath value="../temp/Classify.mdb" class="smallInput" readonly>
              &nbsp;&nbsp;<BR>
  					&nbsp;&nbsp;���ݿ����Ƽ�·��ӦΪ��temp/Classify.mdb<BR>
						&nbsp;&nbsp;
              <input type=submit value="��һ��" class="button">
                </td>
	</tr>
	<input type="hidden" name="action" value="Restore">
  </form></TD>
  </TR>
<table>
</div>

<%End if
if request("action")="xxflto" Then
set rsmb=server.createobject("adodb.recordset")
sql = "select * from DNJY_type"
rsmb.open sql,connmb,1,1
if rsmb.eof then
response.write "<table>��û��Ĭ����Ϣ���࣬�޷����룡</table>"
rsmb.close
set rsmb=Nothing
response.end
end If
Set Rsi=server.createobject("adodb.recordset")
	Sql="Select * From DNJY_type"
	Rsi.open Sql,conn,1,3
	IF Not Rsi.Eof Then
	response.write "<script language='javascript'>	alert('��ǰϵͳ�Ѵ��ڷ���,Ҫȫ��ɾ������ܵ��룡');location.href='javascript:history.go(-1)';</script>"		
	rsi.close:Set rsi=Nothing
    rsmb.close:set rsmb=Nothing
	Response.End
	Else	
do while not rsmb.eof	

	Rsi.Addnew
	Rsi("ID")=Rsmb("ID")
	Rsi("twoid")=Rsmb("twoid")
	Rsi("threeid")=Rsmb("threeid")
	Rsi("indexid")=Rsmb("indexid")
	Rsi("name")=Rsmb("name")
	Rsi("titlecolor")=Rsmb("titlecolor")
	Rsi("indexshow")=Rsmb("indexshow")
	Rsi.update	

rsmb.movenext
Loop
End if
Call CloseRs(rsi)
rsmb.close
set rsmb=Nothing
response.write "<script language=JavaScript>" & chr(13) & "alert('����ٷ�Ĭ����Ϣ����ɹ���');" &"window.location='type_Level1.asp'" & "</script>"
end If
'=============================������Ϣ����end====================================


'=============================�����������====================================
if request("action")="cityfl" Then%>
<div align="center">
  <center>
<table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#ADAED6" width="100%" >
<TR> 
    <TD  bgcolor="#799AE1" align="center">  <form action="?Action=cityflto" name="form1" method="post">
    <tr valign="middle"> 
      <td bgcolor="#FFFFFF"><br>
	 	  &nbsp;&nbsp;Ĭ�ϵ������ർ����Դ���ݿ⣺
              <input type=text size=50 name=Importpath value="../temp/Classify.mdb" class="smallInput" readonly>
              &nbsp;&nbsp;<BR>
  					&nbsp;&nbsp;���ݿ����Ƽ�·��ӦΪ��temp/Classify.mdb<BR>
						&nbsp;&nbsp;
              <input type=submit value="��һ��" class="button">
                </td>
	</tr>
	<input type="hidden" name="action" value="Restore">
  </form></TD>
  </TR>
<table>
</div>

<%End if
if request("action")="cityflto" Then
set rsmb=server.createobject("adodb.recordset")
sql = "select * from DNJY_city"
rsmb.open sql,connmb,1,1
if rsmb.eof then
response.write "<table>��û��Ĭ�ϵ������࣬�޷����룡</table>"
rsmb.close
set rsmb=Nothing
response.end
end If
Set Rsi=server.createobject("adodb.recordset")
	Sql="Select * From DNJY_city"
	Rsi.open Sql,conn,1,3
	IF Not Rsi.Eof Then
	response.write "<script language='javascript'>	alert('��ǰϵͳ�Ѵ��ڷ���,Ҫȫ��ɾ������ܵ��룡');location.href='javascript:history.go(-1)';</script>"		
	rsi.close:Set rsi=Nothing
    rsmb.close:set rsmb=Nothing
	Response.End
	Else	
do while not rsmb.eof	

	Rsi.Addnew
	Rsi("ID")=Rsmb("ID")
	Rsi("twoid")=Rsmb("twoid")
	Rsi("threeid")=Rsmb("threeid")
	Rsi("indexid")=Rsmb("indexid")
	Rsi("city")=Rsmb("city")
	Rsi("cityadmin")=Rsmb("cityadmin")
	Rsi("citypass")=Rsmb("citypass")
	Rsi("indexshow")=Rsmb("indexshow")
	Rsi("dname")=Rsmb("dname")
	Rsi("WebSetting")=Rsmb("WebSetting")
	Rsi.update
rsmb.movenext
Loop
End if
Call CloseRs(rsi)
rsmb.close
set rsmb=Nothing
response.write "<script language=JavaScript>" & chr(13) & "alert('����ٷ�Ĭ�ϵ�������ɹ���');" &"window.location='city_Level1.asp'" & "</script>"
end If
'=============================�����������end====================================

Call CloseDB(conn)%>