<!--#include file="../conn.asp"-->
<!--#include file="../config.asp"-->
<!--#include file="../Include/Function.asp"-->
<!--#include file=cookies.asp-->
<%
'=====================================
'���칩����Ϣ��վ����ϵͳ
'��������Ƽ������ҳ�Ʒ
'�ͷ�����http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
call checkmanage("06")%>	
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<title>ͼƬ�б�</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<meta name="description" content="ͼƬ�б�" />
<meta name="keywords" content="ͼƬ�б�" />
<link rel="stylesheet" type="text/css" href="../css/photo.css" media="screen" />

<script type="text/javascript"> 
    function jsCopy(){ 
        var e=document.getElementById("content");//������content 
        e.select(); //ѡ����� 
        document.execCommand("Copy"); //ִ���������������

       alert("�Ѹ��ƺã�����ճ��"); 
    } 
</script> 
</head>
<body>
<div id="wrap">
<div class="header">	
	<div class="menu">
		<ul>
			<li><a href="admin_article_upload.asp?photoid=<%=request.querystring("id")%>" target="_self" class="current">����ϴ�</a></li>			
			<li><a href="admin_article_piclist.asp?id=<%=request.querystring("id")%>" target="_self" class="current">����б�</a></li>						
			<li><a href="admin_article_edit.asp?id=<%=request.querystring("id")%>" target="_self" class="current">������</a></li>			
		</ul>
	</div>
</div>
<%If request.querystring("action")="del" Then 
	If request.querystring("id")<>"" Then 
	If IsNumeric(request.querystring("photoid"))=False Then response.redirect "admin_article_edit.asp?id="&request.querystring("id")&""
	set rs=server.createobject("adodb.recordset")
	sql="select * from DNJY_news where id = " & request.querystring("id")
	rs.open sql,conn,1,3	
	If Not rs.eof Then
	DoDelFile("/"&strInstallDir&Replace(request("file"),"_s",""))'ɾ��ԭͼ
    DoDelFile("/"&strInstallDir&"UploadFiles/upnews/"&Replace(request("file"),"UploadFiles/upnews/",""))'ɾ������ͼ
	rs("photo")=Replace(rs("photo"),"|"&Replace(request("file"),"_s",""),"")'����ԭͼ��¼
	rs("s_photo")=Replace(rs("s_photo"),"|"&request("file"),"")'��������ͼ��¼
	If rs("photo")="" Then rs("img")=false
	rs.update	
	End If 
	rs.close
	Set rs=Nothing 
	Set sql=Nothing 
	End If 
	response.redirect "admin_article_edit.asp?id="&request.querystring("id")&""
End If 
%>
<div class="main">
	<div class="index_left" style="width:1000px;">
	<%set rs=server.createobject("adodb.recordset")
	sql="select * from DNJY_news where id = " & request.querystring("id")
	rs.open sql,conn,1,1%>
		<h3>���������� �������£�<a href="../news_show.asp?id=<%=request.querystring("id")%>" target="_blank" style="font-weight:900;"><%=rs("title")%></a></h3>
<%If rs("s_photo")<>"" Then
Dim Bimg,Strsimg
Bimg=split(Mid(rs("s_photo"),2),"|")
For Each Strsimg In Bimg
		%>
		<div class="picture"><a href="../<%=Replace(Strsimg,"_s","")%>" target="_blank"><img src="../UploadFiles/upnews/<%=mid(Strsimg,19)%>" border="0" alt="" width="170" height="140"/><br/><%=mid(Replace(Strsimg,"_s",""),19)%></a><span class="del"><a href="?action=del&id=<%=request.querystring("id")%>&file=<%=Strsimg%>" onClick="{if(confirm('ɾ����ͼƬ���޷��ָ���ȷ������ɾ��ô��')){return true;}return false;}">ɾ��</a></span></div>
		<%Next
		Else
		response.write"<div class=""picture"">��������ͼƬ���</div>"
		End if%>
		<div class="clear"></div>
			<div class="clear"></div>
		</div>
	</div>
</div>
	<div class="clear"></div>
</div>
</body>
</html>

<%
Function ShowSpacesize(drvpath)'�����ļ�
	dim fso,d,size,showsize
	set fso=server.createobject("scripting.filesystemobject") 		
	drvpath=server.mappath(drvpath) 		 		
	set d=fso.getfolder(drvpath) 		
	size=d.size
	ShowSpacesize=size
 End Function	
Function spacesize(size)'�����ļ���С
	spacesize=size &" Byte"
	if size>1024 then
		 size=(size\1024)
		 spacesize=size & "&nbsp;KB"
	end if
	if size>1024 then
		 size=(size/1024)
		 spacesize=formatnumber(size,2) & "&nbsp;MB"		
	end if
	if size>1024 then
		 size=(size/1024)
		 spacesize=formatnumber(size,2) & "&nbsp;GB"	   
	end if   
End Function 

sub delfile(filename)'ɾ��
	dim fso
	set fso=server.createobject("scripting.filesystemobject")
	delefilepath=Server.MapPath(filename)
	if fso.FileExists(delefilepath) then
			fso.DeleteFile(delefilepath)
	end if
	set fso = nothing
end Sub
%>