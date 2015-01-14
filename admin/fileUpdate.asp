<!--#include file="../conn.asp"-->
<!--#include file="../config.asp"-->
<!--#include file=cookies.asp-->
<!--#include file="../Include/filename.asp"-->
<!--#include file="inc/Function.asp"-->
<%call checkmanage("14")%>
<link rel="stylesheet" type="text/css" href="../css/css.css">
<BODY>
<style type="text/css">
<!--
body {
	background-color: #E3EBF9;
}
.style1 {
	color: #FFFFFF;
	font-weight: bold;
}
.style2 {color: #FF0000}
-->
</style>
<%If request("key")<>"ok" then%>
<table width="100%" border="0" cellpadding="0" cellspacing="0" bordercolor="#111111" style="border-collapse: collapse">
  <tr>
<td align="center">随机更改入口文件名的作用主要是防范机器群发注册会员和发布垃圾信息，可根据情况不定期更改一下入口文件名。如果没有灌水现象就不要更名。<br>
<%Dim i,n1,n2,n3,n4
n1=""
n2=""
n3=""
n4=""
Response.Write "<b>上次更名时间：</b>"&gmsj&"<br>"
Set fso = Server.CreateObject("Scripting.FileSystemObject") 
'if FSO.fileExists(Server.MapPath("../"&reg)) Then
'Response.Write "文件"&reg&"找到。<br>"
'n1="ok"
'Else
'Response.Write "<font color=red>文件"&reg&"没找到。</font><br>"
'End If
if FSO.fileExists(Server.MapPath("../"&regto)) Then
Response.Write "文件"&regto&"找到。<br>"
n2="ok"
Else
Response.Write "<font color=red>文件"&regto&"没找到。</font><br>"
End If
if FSO.fileExists(Server.MapPath("../"&addinfo)) Then
Response.Write "文件"&addinfo&"找到。<br>"
n3="ok"
Else
Response.Write "<font color=red>文件"&addinfo&"没找到。</font><br>"
End If
if FSO.fileExists(Server.MapPath("../"&adduserinfo)) Then
Response.Write "文件"&adduserinfo&"找到。<br>"
n4="ok"
Else
Response.Write "<font color=red>文件"&adduserinfo&"没找到。</font><br>"
End if
If n1="ok" And n2="ok" And n3="ok" And n4="ok" then
Response.Write "<a href=?key=ok><b>确定更名</b> </a>"
Else
Response.Write "<b>部分文件没找到，不能更名，请按下面方法操作</b>"
End if
%></td></tr>
<tr>
<td >如果更名不成功，或前台找不到入口文件，请手工操作：<br>
1、下载inc目录下的文件filename.asp，用记事本打开，修改引号内的文件名与网站上根目录下的相应文件名（用FTP查看）相同。文件名形式如下，红色部分为随机生成的，以你网站当前文件名为准。<br>
<!--a.Reg.<font color=red>eddfhe</font>.asp<br>-->
a.Regto.<font color=red>hhbiii</font>.asp<br>
a.addinfo.<font color=red>gegfgf</font>.asp<br>
a.adduserinfo.<font color=red>hfjiie</font>.asp<br>
2、将修改好的文件filename.asp上传到inc目录下覆盖。<br>
3、在本页再次确定更名，若再不成功请与客服联系。
</td>
</tr>
</table>
<%Else
dim ObjInstalled
Const Script_FSO="Scripting.FileSystemObject"
ObjInstalled=IsObjInstalled(Script_FSO)
If ObjInstalled=false Then
 Response.Write "<b><font color=red>你的服务器不支持 FSO(Scripting.FileSystemObject)! 不能使用本功能。 </font></b>"
 Response.end
End If

'创建fso操作对象 
'****************调用例子********************** 
'fileUpdate ""&Reg&"" ,""&get_rand_name("Reg")&".asp"
fileUpdate ""&Reg&"" ,"user_regagree.asp"
fileUpdate ""&Regto&"" ,""&get_rand_name("Regto")&".asp"
fileUpdate ""&addinfo&"" ,""&get_rand_name("addinfo")&".asp"
fileUpdate ""&adduserinfo&"" ,""&get_rand_name("adduserinfo")&".asp"
Dim fso,file 
Function fileUpdate(folderName ,extendName)
Set fso = Server.CreateObject("Scripting.FileSystemObject") 
if FSO.fileExists(Server.MapPath("../"&folderName)) then
fso.movefile Server.MapPath("../"&folderName&""),Server.MapPath("../"&extendName&"")

file=file+extendName+"|"
End If 
End Function 

'获取随机文件名字 
Function get_rand_name(nameExtend) 
Dim fileName 
Randomize 
fileName = Int(rnd()*1000000) 
fileName = change_number(fileName) 
fileName = "a."&nameExtend &"."& fileName 
get_rand_name = fileName 
End Function 
'改数字为字母 
Function change_number(number) 
Dim str 
str = CStr(number) 
Dim strArr 
strArr = Array("a","b","c","d","e","f","g","h","i","j") 
Dim strNew 
strNew = "" 
For i = 1 To Len(str) 
   If Mid(str ,i ,1) <> "" Then 
   strNew = strNew & strArr(CInt(Mid(str ,i ,1))) 
   End If 
Next 
change_number = strNew 
End Function
file=split(file,"|")
 dim hf
 set hf=fso.CreateTextFile(Server.mappath("../Include/filename.asp"),true)
 hf.write "<" & "%" & vbcrlf
 hf.write "dim gmsj,reg,regto,addinfo,adduserinfo" & vbcrlf
 hf.write "gmsj="""&now()&"""" & vbcrlf
 hf.write "reg="""&file(0)&"""" & vbcrlf
 hf.write "regto="""&file(1)&"""" & vbcrlf  
 hf.write "addinfo="""&file(2)&"""" & vbcrlf
 hf.write "adduserinfo="""&file(3)&"""" & vbcrlf 
 hf.write "%" & ">" & vbcrlf
 hf.close
response.write "已完成"
'销毁fso操作对象 
Set fso = Nothing
response.end
End If%> 