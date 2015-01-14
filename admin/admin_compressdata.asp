<!--#include file="../conn.asp"-->
<!--#include file="../config.asp"-->
<!--#include file=cookies.asp-->
<!--#include file="error.asp"-->

<%
'=====================================
'天天供求信息网站管理系统
'天天网络科技工作室出品
'客服中心http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>
<%call checkmanage("14")
dim founderr,errmsg
founderr=false
errmsg=""

dim tmprs,allarticle,Maxid,topic,username,dateandtime,body,dbpath,boolIs97
dbpath = request("dbpath")
boolIs97 = request("boolIs97")
If dbpath <> "" Then
dbpath = server.mappath(dbpath)
	response.write(CompactDB(dbpath,boolIs97))
End If

Const JET_3X = 4

Function CompactDB(dbPath, boolIs97)
Dim fso, Engine, strDBPath
strDBPath = left(dbPath,instrrev(DBPath,"\"))
Set fso = CreateObject("Scripting.FileSystemObject")
On Error Resume Next 
If fso.FileExists(dbPath) Then
Set Engine = CreateObject("JRO.JetEngine")

	If boolIs97 = "True" Then
		Engine.CompactDatabase "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & dbpath, _
		"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strDBPath & "temp.mdb;" _
		& "Jet OLEDB:Engine Type=" & JET_3X
	Else
		Engine.CompactDatabase "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & dbpath, _
		"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strDBPath & "temp.mdb"
	End If

fso.CopyFile strDBPath & "temp.mdb",dbpath
fso.DeleteFile(strDBPath & "temp.mdb")

if err.number="0" then
	errmsg = "<li>你的数据库已经压缩成功!<br>"
	errmsg = errmsg & dbpath
	call disok()
	response.end

else
	errmsg = "<li>压缩过程中出现错误，具体出错如下：<br>"
	errmsg = errmsg & Err.Description
	err.clear
	call diserror()
	response.end
end if
Set fso = nothing
Set Engine = nothing


Else
	response.write "<li>数据库名称或路径不正确. 请重试!"
	call diserror()
	response.end
End If

End Function
%>
<HTML>
<HEAD>
<TITLE> New Document </TITLE>
<META NAME="Generator" CONTENT="EditPlus">
<META NAME="Author" CONTENT="">
<META NAME="Keywords" CONTENT="">
<META NAME="Description" CONTENT="">
<LINK href="../Include/djcss" type=text/css rel=StyleSheet>
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
</head>
<BODY>
<BODY>
<table align="center" width="100%" border="1" cellspacing="0" cellpadding="4" bordercolor="#ADAED6" style="border-collapse: collapse">
  <form action="" name="form1" method="post">
    <tr> 
      <td height="28" colspan="2" align="center" bgcolor="#BDBEDE"><div align="left"><span class="style4"><span style="font-size: 9pt"><font size="3"><font size="2">&nbsp;</font></font>压缩数据库（对access数据库有效）</span></span></div></td>
    </tr>
    <tr valign="middle"> 
	  <td bgcolor="#FFFFFF"><br>
	  <b>注意：</b>( 需要FSO支持，FSO相关帮助请看微软网站 ) <br>输入数据库所在相对路径,并且输入数据库名称（<font color="red">正在使用中数据库不能压缩，请选择备份数据库进行压缩操作</font>）

      <br>
      压缩数据库：
        <input name="dbpath" type="text" class="smallInput" value=../DataBackup/#databack#.asp size="50"  readonly>
        &nbsp;<br><br>
        <input type="submit" value="开始压缩" class="button">
<BR><br>
<input type="checkbox" name="boolIs97" value="True">如果使用 Access 97 数据库请选择
(默认为 Access 2000 数据库)<br><br>	  </td>
	</tr>
  </form>
</table>