<!--#include file="../conn.asp"-->
<!--#include file="../config.asp"-->
<!--#include file="../SqlIn.Asp"-->
<!--#include file=cookies.asp-->

<%
'=====================================
'天天供求信息网站管理系统
'天天网络科技工作室出品
'客服中心http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>
<%call checkmanage("02")%>
<HTML>
<HEAD>
<TITLE>店企－地区－信息－分类管理</TITLE>
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
    <TD height="5" colspan="12" bgcolor="#799AE1" align="center"><FONT color="#FFFFFF" style="font-size:14px">默认分类导入管理</FONT></TD>
  </TR>
<table>
<%Dim DBmb,connstrmb,connmb,action,id,Importpath,Importto,str2,rsi,rsmb,k
action=request.form("action")
Call OpenConn
DBmb="../temp/Classify.mdb"  '可修改
connstrmb="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(""&DBmb&"")
  On Error Resume Next
	Set connmb = Server.CreateObject("ADODB.Connection")
	connmb.open connstrmb
	If Err Then
		err.Clear
		Set Connmb = Nothing
		Response.Write "数据库连接出错，请检查连接字串。"'注释，需要把这几个字翻译成英文。
		Response.End
End If

'=============================导入店企分类====================================
if request("action")="dqfl" Then%>
<div align="center">
  <center>
<table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#ADAED6" width="100%" >
<TR> 
    <TD  bgcolor="#799AE1" align="center">  <form action="?Action=dqflto" name="form1" method="post">
    <tr valign="middle"> 
      <td bgcolor="#FFFFFF"><br>
	 	  &nbsp;&nbsp;默认店企分类导入来源数据库：
              <input type=text size=50 name=Importpath value="../temp/Classify.mdb" class="smallInput" readonly>
              &nbsp;&nbsp;<BR>
  					&nbsp;&nbsp;数据库名称及路径应为：temp/Classify.mdb<BR>
						&nbsp;&nbsp;
              <input type=submit value="下一步" class="button">
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
response.write "<table>还没有默认店企分类，无法导入！</table>"
rsmb.close
set rsmb=Nothing
response.end
end If
Set Rsi=server.createobject("adodb.recordset")
	Sql="Select * From DNJY_hy_type"
	Rsi.open Sql,conn,1,3
	IF Not Rsi.Eof Then
	response.write "<script language='javascript'>	alert('当前系统已存在分类,要全部删除后才能导入！');location.href='javascript:history.go(-1)';</script>"		
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
response.write "<script language=JavaScript>" & chr(13) & "alert('导入官方默认店企分类成功！');" &"window.location='store_Level1.asp'" & "</script>"
End if
'=============================导入店企分类end====================================


'=============================导入信息分类====================================
if request("action")="xxfl" Then%>
<div align="center">
  <center>
<table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#ADAED6" width="100%" >
<TR> 
    <TD  bgcolor="#799AE1" align="center">  <form action="?Action=xxflto" name="form1" method="post">
    <tr valign="middle"> 
      <td bgcolor="#FFFFFF"><br>
	 	  &nbsp;&nbsp;默认信息分类导入来源数据库：
              <input type=text size=50 name=Importpath value="../temp/Classify.mdb" class="smallInput" readonly>
              &nbsp;&nbsp;<BR>
  					&nbsp;&nbsp;数据库名称及路径应为：temp/Classify.mdb<BR>
						&nbsp;&nbsp;
              <input type=submit value="下一步" class="button">
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
response.write "<table>还没有默认信息分类，无法导入！</table>"
rsmb.close
set rsmb=Nothing
response.end
end If
Set Rsi=server.createobject("adodb.recordset")
	Sql="Select * From DNJY_type"
	Rsi.open Sql,conn,1,3
	IF Not Rsi.Eof Then
	response.write "<script language='javascript'>	alert('当前系统已存在分类,要全部删除后才能导入！');location.href='javascript:history.go(-1)';</script>"		
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
response.write "<script language=JavaScript>" & chr(13) & "alert('导入官方默认信息分类成功！');" &"window.location='type_Level1.asp'" & "</script>"
end If
'=============================导入信息分类end====================================


'=============================导入地区分类====================================
if request("action")="cityfl" Then%>
<div align="center">
  <center>
<table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#ADAED6" width="100%" >
<TR> 
    <TD  bgcolor="#799AE1" align="center">  <form action="?Action=cityflto" name="form1" method="post">
    <tr valign="middle"> 
      <td bgcolor="#FFFFFF"><br>
	 	  &nbsp;&nbsp;默认地区分类导入来源数据库：
              <input type=text size=50 name=Importpath value="../temp/Classify.mdb" class="smallInput" readonly>
              &nbsp;&nbsp;<BR>
  					&nbsp;&nbsp;数据库名称及路径应为：temp/Classify.mdb<BR>
						&nbsp;&nbsp;
              <input type=submit value="下一步" class="button">
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
response.write "<table>还没有默认地区分类，无法导入！</table>"
rsmb.close
set rsmb=Nothing
response.end
end If
Set Rsi=server.createobject("adodb.recordset")
	Sql="Select * From DNJY_city"
	Rsi.open Sql,conn,1,3
	IF Not Rsi.Eof Then
	response.write "<script language='javascript'>	alert('当前系统已存在分类,要全部删除后才能导入！');location.href='javascript:history.go(-1)';</script>"		
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
response.write "<script language=JavaScript>" & chr(13) & "alert('导入官方默认地区分类成功！');" &"window.location='city_Level1.asp'" & "</script>"
end If
'=============================导入地区分类end====================================

Call CloseDB(conn)%>