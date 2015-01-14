<%
'=====================================
'天天供求信息网站管理系统
'天天网络科技工作室出品
'客服中心http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
if request.cookies("administrator")="" and Session("DNJY_cityadmin")="" then 
Response.Write "抱歉：你没有此项操作的权限！"
response.end
end if
session("ttup")=request("ttup")
session("dform")=request("dform")
session("fupname")=request("fupname")		'传递文件名
session("frmname")=request("frmname")		'传递表单名
Server.ScriptTimeOut=99999                      '脚本运行时间

%>
<html>
<head>
<title>文件上传</title>
<meta name="Description" Content="">
<link href="../css/style.css" rel="stylesheet" type="text/css">

</head>
<body> 
<form name="form1" method="post" action="DNJY_upsave_ad.asp" enctype="multipart/form-data">
<b>请选择要上传的文件：</b><br>
<input type=file name="file1">
<input type=submit name="submit" value="上传"><br><br>
·限jpg;jpeg;gif;bmp;png;swf格式，500K内<br>
·点击“上传”后，请耐心等待（不要重复点击“上传”），上传时间视文件大小和网络状况而定<br>
·传送大文件时，可能导致服务器变慢或者不稳定。建议使用FTP上传大文件。
</form>
</body>  
</html>
