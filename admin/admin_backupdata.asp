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
if request("action")="Backup" then
call backupdata()
else
%>
<HTML>
<HEAD>
<TITLE> New Document </TITLE>
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
      <td height="28" colspan="2" align="center" bgcolor="#BDBEDE"><div align="left"><span class="style4"><span style="font-size: 9pt"><font size="3"><font size="2">&nbsp;</font></font>�������ݿ⣨��access���ݿ���Ч��</span></span> </div></td>
    </tr>
    <tr valign="middle"> 
      <td bgcolor="#FFFFFF"><br>
      &nbsp;&nbsp;&nbsp;<B>��������</B>( ��ҪFSO֧�֣�FSO��ذ����뿴΢����վ ) 
  <form method="post" action="admin_BackupData.asp">
        <br>&nbsp;&nbsp; ��ǰ���ݿ�·��(���·��)�� 
        <input type=text size=40 name=DBpath value=<%=DBFileName%> class="smallInput" readonly>
        &nbsp;&nbsp; ���·�����Կ��ڸ�Ŀ¼Database.asp�ļ����޸����ݿ��·��������<BR>
        &nbsp;&nbsp; �������ݿ�Ŀ¼(���·��)�� 
        <input type=text size=25 name=bkfolder value="../DataBackup" class="smallInput" readonly>
        &nbsp;��Ŀ¼�����ڣ������Զ�����<BR>
        &nbsp;&nbsp; �������ݿ�����(��д����)�� 
        <input type=text size=20 name=bkDBname value="#databack#.asp" class="smallInput" readonly>
        &nbsp;�籸��Ŀ¼�и��ļ��������ǣ���û�У����Զ�����<BR>
        &nbsp;&nbsp; 
        <input type=submit value="ȷ������" class="button">
        <br>
        -----------------------------------------------------------------------------------------<br>
        &nbsp;&nbsp;��������д����������ݿ�·��ȫ����<br>
        &nbsp;&nbsp;������������������������ķ������ݣ��Ա�֤�������ݰ�ȫ��<br>
        &nbsp;&nbsp;ע�⣺����·��������������ռ�ADMINĿ¼�����·�� </font> 
		<input type="hidden" name="action" value="Backup">
  </form>
<%end if%>
<%Dim Dbpath,bkfolder,bkdbname,Fso,folderpath,fso1,f

sub backupdata()
		Dbpath=request.form("Dbpath")
		Dbpath=server.mappath(Dbpath)
		bkfolder=request.form("bkfolder")
		bkdbname=request.form("bkdbname")
		Set Fso=server.createobject("scripting.filesystemobject")
		if fso.fileexists(dbpath) then
		If CheckDir(bkfolder) = True Then
		fso.copyfile dbpath,bkfolder& "\"& bkdbname
		else
		MakeNewsDir bkfolder
		fso.copyfile dbpath,bkfolder& "\"& bkdbname
		end if
		errmsg="<li>�ɹ��������ݣ�"
		call disok()
 		Else
                  errmsg="�������ݿ�ʧ�ܣ���Ҫ���ݵ����ݿⲻ����"
		call diserror()
		End if
end sub
'------------------���ĳһĿ¼�Ƿ����-------------------
Function CheckDir(FolderPath)
	folderpath=Server.MapPath(".")&"\"&folderpath
    Set fso1 = CreateObject("Scripting.FileSystemObject")
    If fso1.FolderExists(FolderPath) then
       '����
       CheckDir = True
    Else
       '������
       CheckDir = False
    End if
    Set fso1 = nothing
End Function
'-------------����ָ����������Ŀ¼---------
Function MakeNewsDir(foldername)
    Set fso1 = CreateObject("Scripting.FileSystemObject")
        Set f = fso1.CreateFolder(foldername)
        MakeNewsDir = True
    Set fso1 = nothing
End Function
%>
	  </td>
    </tr>
</table>
</body></html>