<%
'=====================================
'���칩����Ϣ��վ����ϵͳ
'��������Ƽ������ҳ�Ʒ
'�ͷ�����http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
if request.cookies("administrator")="" and Session("DNJY_cityadmin")="" then 
Response.Write "��Ǹ����û�д��������Ȩ�ޣ�"
response.end
end if
session("ttup")=request("ttup")
session("dform")=request("dform")
session("fupname")=request("fupname")		'�����ļ���
session("frmname")=request("frmname")		'���ݱ���
Server.ScriptTimeOut=99999                      '�ű�����ʱ��

%>
<html>
<head>
<title>�ļ��ϴ�</title>
<meta name="Description" Content="">
<link href="../css/style.css" rel="stylesheet" type="text/css">

</head>
<body> 
<form name="form1" method="post" action="DNJY_upsave_ad.asp" enctype="multipart/form-data">
<b>��ѡ��Ҫ�ϴ����ļ���</b><br>
<input type=file name="file1">
<input type=submit name="submit" value="�ϴ�"><br><br>
����jpg;jpeg;gif;bmp;png;swf��ʽ��500K��<br>
��������ϴ����������ĵȴ�����Ҫ�ظ�������ϴ��������ϴ�ʱ�����ļ���С������״������<br>
�����ʹ��ļ�ʱ�����ܵ��·������������߲��ȶ�������ʹ��FTP�ϴ����ļ���
</form>
</body>  
</html>
