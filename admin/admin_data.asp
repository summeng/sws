<!--#include file="../conn.asp"-->
<!--#include file="../config.asp"-->
<!--#include file=cookies.asp-->
<%
'=====================================
'天天供求信息网站管理系统
'天天网络科技工作室出品
'客服中心http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>
<%call checkmanage("14")%>
<link rel="stylesheet" type="text/css" href="../css/css.css">
<style type="text/css">
　　<!--
　　a { color: #0A0101; text-decoration: none}
　　a:hover { color: #0A0101; text-decoration: underline}
　　-->
</style>
<style type="text/css">
<!--
.style1 {font-size: 12px}
body {
	margin-top: 5px;
	margin-left: 5px;
	margin-right: 5px;
	background-color: #f5f9fd;
}
.style2 {color: #CC3300}
.style3 {color: #FF0000}
-->
</style>   
<html>   
<head>   
<meta   http-equiv="Content-Type"   content="text/html;   charset=gb2312">   
<title>数据库管理</title>   
 </head>       
<body class="ss">
<table border="0" width="98%" align="center" cellspacing="0" cellpadding="0">
        <tr>
          <td  height="2"></td>
        </tr>
</table>
<table border="0" width="100%" style="font-size: 12px; border-bottom:#D6DDE9 2px solid; border-left:#D6DDE9 1px solid; border-top:#D6DDE9 1px solid; border-right:#D6DDE9 2px solid"  align="center" cellpadding="4" cellspacing="0"  bordercolorlight="#D6DDE9" bordercolordark="#ffffff">

        <tr>
          <td  height="32" align="center" class="bg"><strong><img src="img/save.gif" width="16" height="16">&nbsp;&nbsp;网&nbsp;&nbsp;站&nbsp;&nbsp;数&nbsp;&nbsp;据&nbsp;&nbsp;库&nbsp;&nbsp;在&nbsp;&nbsp;线&nbsp;&nbsp;管&nbsp;&nbsp;理</strong></td>
        </tr>
               <tr>
          <td  height="457" align="center">
       
<%'response.end'为了安全，不使用后台备份数据库功能
If SystemDatabaseType = 1 Then
	response.write "<li>本功能只适应ACCESS数据库，目前数据库为MS SQL，不能使用此功能!<br>MS SQL数据库管理可通过空间商给的数据库管理后台或远程桌面管理（要有权限才行）进行。具体咨询空间商。<br>"
	response.end
End if
Dim action,dbpath,temp,fileext,file1,file2,Engine   
Dim   fso,   f,   f1,   fc 
Call OpenConn
on error resume next
action=request.QueryString("action")
'filename=request("filename")  
select   case   action   
case   "beifen" 
Set   fso   =   CreateObject("Scripting.FileSystemObject")   
file1=fso.GetParentfoldername(server.MapPath(".\"))&"\"&Replace(DBFileName,"../","")   
file2=fso.GetParentfoldername(server.MapPath("./"))&"\databackup\"&request.Form("filename")   
fso.copyfile   file1,file2   
set   fso=nothing   
response.write "<script language='javascript'>" & chr(13)
response.write "alert('数据库备份成功！');" & Chr(13)
response.write "window.document.location.href='admin_data.asp';"&Chr(13)
response.write "</script>" & Chr(13)
Response.End()   
case   "yasuo"

dim founderr,errmsg
founderr=false
errmsg=""

dim boolIs97
dbpath = "/"&strInstallDir&"databackup/"&request("filename")
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
	response.write "<li>你的数据库"&request("filename")&"已经压缩成功!<br>"
	response.end

else
	response.write "<li>压缩过程中出现错误!<br>"
	response.end
end if
Set fso = nothing
Set Engine = nothing
Else
	response.write "<li>数据库名称或路径不正确. 请重试!"
	response.end
End If

End Function


Response.End()  
case   "huanyuan"   
Set   fso   =   CreateObject("Scripting.FileSystemObject")   
file1=fso.GetParentfoldername(server.MapPath("./"))&"\"&Replace(DBFileName,"../","")   
file2=fso.GetParentfoldername(server.MapPath("./"))&"\databackup\"&request.Form("filename")   
fso.copyfile   file2,file1   
set   fso=nothing   
response.write "<script language='javascript'>" & chr(13)
response.write "alert('还原完成！');" & Chr(13)
response.write "window.document.location.href='admin_data.asp';"&Chr(13)
response.write "</script>" & Chr(13)
Response.End() 

case   "del"
	Set fso=Server.CreateObject("Scripting.FileSystemObject")
    If fso.FileExists(Server.mappath("..\databackup\"&request.Form("filename")&"")) Then
       fso.DeleteFile Server.mappath("..\databackup\"&request.Form("filename")&"")
    End If
response.write "<script language='javascript'>" & chr(13)
response.write "alert('删除完成！');" & Chr(13)
response.write "window.document.location.href='admin_data.asp';"&Chr(13)
response.write "</script>" & Chr(13)
Response.End()

case   else   
dataform   
end   select   
    
function   dataform()   
%>   
<table   width="87%"     border="0"   cellspacing="0"   cellpadding="0">   
        <tr>   
          <td  width="100%" height="50" align="center" colspan="2"> 注：操作仅对ACCESS数据库有效，MS SQL数据库不能操作！;</td>   
      </tr>      
      <tr>   
          <td   width="31%" height="50" align="right"><strong><img src="img/database_add.png" width="16" height="16">&nbsp;备份：&nbsp;</strong></td>   
          <form   name="form1"   method="post"   action="admin_data.asp?action=beifen"><td   width="53%" align="left">   
              <input   name="filename"  type="text"  id="filename"   style="width:200px" readonly value="<%=filename()%>">   
              <input   type="submit" name="Submit"  value="开始备份" class="button1" onmouseover=this.className="button2";   onmouseout=this.className="button1"; style="height:26px;line-height:20px"> <br> 注：数据库名称由系统自动生成！  
          </td></form>   
          <td   width="16%">&nbsp;</td>   
      </tr>   
     
      <tr>   
          <td height="54" align="right"><strong><img src="img/database_go.png" width="16" height="16">&nbsp;还原：&nbsp;</strong></td>   
          <form   name="form3"   method="post"   action="admin_data.asp?action=huanyuan"><td align="left">   
              <select   name="filename"   id="filename"   style="width:200px   ">   
<%   
Dim   fso,   f,   f1,   fc   
Set   fso   =   CreateObject("Scripting.FileSystemObject")   
dbpath=fso.GetParentfoldername(server.MapPath("./"))&"\databackup"   
Set   f   =   fso.GetFolder(dbpath)   
Set   fc   =   f.Files   
For   Each   f1   in   fc   
temp=split(f1.name,".")   
fileext=temp(ubound(temp))   
if   fileext="asa"   then   
response.Write("<option   value='"&f1.name&"'>"&f1.name&"</option>")   
end   if   
next   
set   fso=nothing   
%>   
      </select>   
              &nbsp;<input   type="submit"   name="Submit"   value="开始还原"   class="button1"   onmouseover=this.className="button2";   onmouseout=this.className="button1";   onClick="if(confirm('真的还原吗？')){return   true;}else{return   false;}" style="height:26px;line-height:20px"> <br> 注：还原时请慎重正确选择源数据库！ 
          </td></form>   
          <td>&nbsp;</td>   
      </tr> 
         <tr>   
          <td height="52" align="right"><strong><img src="img/database_refresh.png" width="16" height="16">&nbsp;压缩：&nbsp;</strong></td>   
          <form   name="form2"   method="post"   action="admin_data.asp?action=yasuo"><td align="left">
		                <select   name="filename"   id="filename"   style="width:200px   ">   
<%   

Set   fso   =   CreateObject("Scripting.FileSystemObject")   
dbpath=fso.GetParentfoldername(server.MapPath("./"))&"\databackup"   
Set   f   =   fso.GetFolder(dbpath)   
Set   fc   =   f.Files   
For   Each   f1   in   fc   
temp=split(f1.name,".")   
fileext=temp(ubound(temp))   
if   fileext="asa"   then   
response.Write("<option   value='"&f1.name&"'>"&f1.name&"</option>")   
end   if   
next   
set   fso=nothing   
%>   
      </select>   
	  
              <!--input   name="filename"   type="text"   id="filename"   style="width:200px   "   value="<%=DBFileName%>" readonly="readonly"-->   
              <input   type="submit"   name="Submit"   value="开始压缩"   class="button1"   onmouseover=this.className="button2";   onmouseout=this.className="button1"; style="height:26px;line-height:20px"><br>
			  <input type="checkbox" name="boolIs97" value="True">如果使用 Access 97 数据库请选择
(默认为 Access 2000 数据库)<br>注：只能对备份的数据库进行压缩！   
          </td></form>   
          <td>&nbsp;</td>   
      </tr>
	  
      <tr>   
          <td height="54" align="right"><strong><img src="img/database_go.png" width="16" height="16">&nbsp;删除：&nbsp;</strong></td>   
          <form   name="form3"   method="post"   action="admin_data.asp?action=del"><td align="left">   
              <select   name="filename"   id="filename"   style="width:200px   ">   
<%   
Set   fso   =   CreateObject("Scripting.FileSystemObject")   
dbpath=fso.GetParentfoldername(server.MapPath("./"))&"\databackup"   
Set   f   =   fso.GetFolder(dbpath)   
Set   fc   =   f.Files   
For   Each   f1   in   fc   
temp=split(f1.name,".")   
fileext=temp(ubound(temp))   
if   fileext="asa"   then   
response.Write("<option   value='"&f1.name&"'>"&f1.name&"</option>")   
end   if   
next   
set   fso=nothing   
%>   
      </select>   
              &nbsp;<input   type="submit"   name="Submit"   value="开始删除"   class="button1"   onmouseover=this.className="button2";   onmouseout=this.className="button1";   onClick="if(confirm('真的删除吗？')){return   true;}else{return   false;}" style="height:26px;line-height:20px"> <br> <font color=red>注：删除选定的备份数据库，要看清那个是最新的！</font> 
          </td></form>   
          <td>&nbsp;</td>   
      </tr> 

</table>   
<%   
end   function   
    
function   filename()   
filename=year(now)   
if   len(month(now))=1   then   
filename=filename&"0"&month(now)   
else   
filename=filename&month(now)   
end   if   
if   len(day(now))=1   then   
filename=filename&"0"&day(now)   
else   
filename=filename&day(now)   
end   if   
if   len(hour(now))=1   then   
filename=filename&"0"&hour(now)   
else   
filename=filename&hour(now)   
end   if   
if   len(minute(now))=1   then   
filename=filename&"0"&minute(now)   
else   
filename=filename&minute(now)   
end   if   
if   len(second(now))=1   then   
filename=filename&"0"&second(now)   
else   
filename=filename&second(now)   
end   if   
filename=filename&"."&"asa"   
end   function   
%></td>   
 </tr>
</table>   
</body>   
</html> 