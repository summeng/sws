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
      <td height="28" colspan="2" align="center" bgcolor="#BDBEDE"><div align="left"><span class="style4"><span style="font-size: 9pt"><font size="3"><font size="2">&nbsp;</font></font>备份数据库（对access数据库有效）</span></span> </div></td>
    </tr>
    <tr valign="middle"> 
      <td bgcolor="#FFFFFF"><br>
      &nbsp;&nbsp;&nbsp;<B>备份数据</B>( 需要FSO支持，FSO相关帮助请看微软网站 ) 
  <form method="post" action="admin_BackupData.asp">
        <br>&nbsp;&nbsp; 当前数据库路径(相对路径)： 
        <input type=text size=40 name=DBpath value=<%=DBFileName%> class="smallInput" readonly>
        &nbsp;&nbsp; 如果路径不对可在根目录Database.asp文件中修改数据库的路径和名称<BR>
        &nbsp;&nbsp; 备份数据库目录(相对路径)： 
        <input type=text size=25 name=bkfolder value="../DataBackup" class="smallInput" readonly>
        &nbsp;如目录不存在，程序将自动创建<BR>
        &nbsp;&nbsp; 备份数据库名称(填写名称)： 
        <input type=text size=20 name=bkDBname value="#databack#.asp" class="smallInput" readonly>
        &nbsp;如备份目录有该文件，将覆盖，如没有，将自动创建<BR>
        &nbsp;&nbsp; 
        <input type=submit value="确定备份" class="button">
        <br>
        -----------------------------------------------------------------------------------------<br>
        &nbsp;&nbsp;在上面填写本程序的数据库路径全名。<br>
        &nbsp;&nbsp;您可以用这个功能来备份您的法规数据，以保证您的数据安全！<br>
        &nbsp;&nbsp;注意：所有路径都是相对与程序空间ADMIN目录的相对路径 </font> 
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
		errmsg="<li>成功备份数据！"
		call disok()
 		Else
                  errmsg="备份数据库失败，您要备份的数据库不存在"
		call diserror()
		End if
end sub
'------------------检查某一目录是否存在-------------------
Function CheckDir(FolderPath)
	folderpath=Server.MapPath(".")&"\"&folderpath
    Set fso1 = CreateObject("Scripting.FileSystemObject")
    If fso1.FolderExists(FolderPath) then
       '存在
       CheckDir = True
    Else
       '不存在
       CheckDir = False
    End if
    Set fso1 = nothing
End Function
'-------------根据指定名称生成目录---------
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