<!--#include file="../conn.asp"-->
<!--#include file="../config.asp"-->
<!--#include file=cookies.asp-->
<!--#include file="error.asp"-->

<%
'=====================================
'���칩����Ϣ��վ����ϵͳ
'��������Ƽ������ҳ�Ʒ
'�ͷ�����http://www.ip126.com
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
	errmsg = "<li>������ݿ��Ѿ�ѹ���ɹ�!<br>"
	errmsg = errmsg & dbpath
	call disok()
	response.end

else
	errmsg = "<li>ѹ�������г��ִ��󣬾���������£�<br>"
	errmsg = errmsg & Err.Description
	err.clear
	call diserror()
	response.end
end if
Set fso = nothing
Set Engine = nothing


Else
	response.write "<li>���ݿ����ƻ�·������ȷ. ������!"
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
      <td height="28" colspan="2" align="center" bgcolor="#BDBEDE"><div align="left"><span class="style4"><span style="font-size: 9pt"><font size="3"><font size="2">&nbsp;</font></font>ѹ�����ݿ⣨��access���ݿ���Ч��</span></span></div></td>
    </tr>
    <tr valign="middle"> 
	  <td bgcolor="#FFFFFF"><br>
	  <b>ע�⣺</b>( ��ҪFSO֧�֣�FSO��ذ����뿴΢����վ ) <br>�������ݿ��������·��,�����������ݿ����ƣ�<font color="red">����ʹ�������ݿⲻ��ѹ������ѡ�񱸷����ݿ����ѹ������</font>��

      <br>
      ѹ�����ݿ⣺
        <input name="dbpath" type="text" class="smallInput" value=../DataBackup/#databack#.asp size="50"  readonly>
        &nbsp;<br><br>
        <input type="submit" value="��ʼѹ��" class="button">
<BR><br>
<input type="checkbox" name="boolIs97" value="True">���ʹ�� Access 97 ���ݿ���ѡ��
(Ĭ��Ϊ Access 2000 ���ݿ�)<br><br>	  </td>
	</tr>
  </form>
</table>