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
<td align="center">�����������ļ�����������Ҫ�Ƿ�������Ⱥ��ע���Ա�ͷ���������Ϣ���ɸ�����������ڸ���һ������ļ��������û�й�ˮ����Ͳ�Ҫ������<br>
<%Dim i,n1,n2,n3,n4
n1=""
n2=""
n3=""
n4=""
Response.Write "<b>�ϴθ���ʱ�䣺</b>"&gmsj&"<br>"
Set fso = Server.CreateObject("Scripting.FileSystemObject") 
'if FSO.fileExists(Server.MapPath("../"&reg)) Then
'Response.Write "�ļ�"&reg&"�ҵ���<br>"
'n1="ok"
'Else
'Response.Write "<font color=red>�ļ�"&reg&"û�ҵ���</font><br>"
'End If
if FSO.fileExists(Server.MapPath("../"&regto)) Then
Response.Write "�ļ�"&regto&"�ҵ���<br>"
n2="ok"
Else
Response.Write "<font color=red>�ļ�"&regto&"û�ҵ���</font><br>"
End If
if FSO.fileExists(Server.MapPath("../"&addinfo)) Then
Response.Write "�ļ�"&addinfo&"�ҵ���<br>"
n3="ok"
Else
Response.Write "<font color=red>�ļ�"&addinfo&"û�ҵ���</font><br>"
End If
if FSO.fileExists(Server.MapPath("../"&adduserinfo)) Then
Response.Write "�ļ�"&adduserinfo&"�ҵ���<br>"
n4="ok"
Else
Response.Write "<font color=red>�ļ�"&adduserinfo&"û�ҵ���</font><br>"
End if
If n1="ok" And n2="ok" And n3="ok" And n4="ok" then
Response.Write "<a href=?key=ok><b>ȷ������</b> </a>"
Else
Response.Write "<b>�����ļ�û�ҵ������ܸ������밴���淽������</b>"
End if
%></td></tr>
<tr>
<td >����������ɹ�����ǰ̨�Ҳ�������ļ������ֹ�������<br>
1������incĿ¼�µ��ļ�filename.asp���ü��±��򿪣��޸������ڵ��ļ�������վ�ϸ�Ŀ¼�µ���Ӧ�ļ�������FTP�鿴����ͬ���ļ�����ʽ���£���ɫ����Ϊ������ɵģ�������վ��ǰ�ļ���Ϊ׼��<br>
<!--a.Reg.<font color=red>eddfhe</font>.asp<br>-->
a.Regto.<font color=red>hhbiii</font>.asp<br>
a.addinfo.<font color=red>gegfgf</font>.asp<br>
a.adduserinfo.<font color=red>hfjiie</font>.asp<br>
2�����޸ĺõ��ļ�filename.asp�ϴ���incĿ¼�¸��ǡ�<br>
3���ڱ�ҳ�ٴ�ȷ�����������ٲ��ɹ�����ͷ���ϵ��
</td>
</tr>
</table>
<%Else
dim ObjInstalled
Const Script_FSO="Scripting.FileSystemObject"
ObjInstalled=IsObjInstalled(Script_FSO)
If ObjInstalled=false Then
 Response.Write "<b><font color=red>��ķ�������֧�� FSO(Scripting.FileSystemObject)! ����ʹ�ñ����ܡ� </font></b>"
 Response.end
End If

'����fso�������� 
'****************��������********************** 
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

'��ȡ����ļ����� 
Function get_rand_name(nameExtend) 
Dim fileName 
Randomize 
fileName = Int(rnd()*1000000) 
fileName = change_number(fileName) 
fileName = "a."&nameExtend &"."& fileName 
get_rand_name = fileName 
End Function 
'������Ϊ��ĸ 
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
response.write "�����"
'����fso�������� 
Set fso = Nothing
response.end
End If%> 