<!--#include file="../conn.asp"-->
<!--#include file="../config.asp"-->
<!--#include file="../Include/CheckFile.asp"-->
<!--#include file="../DNJY_clsUp.asp"-->
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
end if%>
<html>
<head>
<title>�ļ��ϴ�</title>
<meta name="Description" Content="">
<link href="../css/style.css" rel="stylesheet" type="text/css">
</head>
<body>
<%'======================================�����ϴ�����=====================================
Dim FileType,UpFileType
Const MaxFileSize=5000000  '�ϴ��ļ���С���ƣ���λ�ֽ�
ttup=session("ttup")
if ttup="slide1" Or ttup="slide2" Or ttup="slide3" Or ttup="slideb" Or ttup="adfp" Then
FileType="jpg;jpeg;"    '������,����������ϴ����ļ�����,���ļ��ĺ�׺�����ж�,���ִ�Сд,ÿ����׺����;�ŷֿ�
elseif ttup="flv" Then
FileType="flv;FLV;"    '������,����������ϴ����ļ�����,���ļ��ĺ�׺�����ж�,���ִ�Сд,ÿ����׺����;�ŷֿ�
Else
FileType="jpg;jpeg;gif;bmp;swf;png;"    '������,����������ϴ����ļ�����,���ļ��ĺ�׺�����ж�,���ִ�Сд,ÿ����׺����;�ŷֿ�
End if
UpFileType=FileType    '������,����������ϴ����ļ�����,���ļ��ĺ�׺�����ж�,���ִ�Сд,ÿ����׺����;�ŷֿ�
Const NoUpFileType = "asp;asa;exe;htm;html;aspx;cs;vb;js;dll;php;jsp;cgi;" '������,�������費���ϴ����ļ�����,���ļ��ĺ�׺�����ж�,���ִ�Сд,ÿ��ÿ׺����;�ŷֿ�,���������Ϊ��,���жϰ���
Const blnAllow=True           '��ֻ�����ǽ�ֹ�ϴ� UpFileType ��ָ���ĺ�׺,True����false��ֹ


'======================================���´��벻Ҫ����=====================================
Call OpenConn
On Error Resume Next
dim fupname,frmname
Dim upfile,formPath,SavePath,ServerPath,FSPath,formName,FileName,oFile,upfilecount,ttup,upname,strExt,fileExt,ranNum,my_string,objFSO,muma,msg,strJS,fso,whichfile,thisfile,dform

If blnAllow=True Then
strExt="ֻ����"
Else
strExt="��ֹ"
End If
ttup=session("ttup")
dform=session("dform")
fupname=session("fupname")  '�����ļ���
frmname=session("frmname")	'���ݱ���
if  fupname="" or frmname="" then
	response.write "<script language='javascript'>"
	response.write "alert('���ִ����������ϴ���');"
	response.write "location.href='javascript:history.go(-1)';"
	response.write "</script>"
	response.end
end if

if ttup="adv" then
savepath="../ads/pic" 	'��׼���ͼƬ
elseif ttup="slide1" Or ttup="slide2" Or ttup="slide3" then
savepath="../UploadFiles/slidea" '�õƹ��A
elseif ttup="slideb" then
savepath="../UploadFiles/slideb" '�õƹ��B
elseif ttup="adfp" then
savepath="../UploadFiles/adfp" '���ƹ��
elseif ttup="flv" then
savepath="../UploadFiles/flv" 'flv���
elseif ttup="weblogo" then
savepath="images" '����
Else
Response.Write "<script language='javascript'>alert('����Ĭ�ϵ��ϴ�ѡ������ϴ�!');history.go(-1);</script>"
response.end
end If
if right(SavePath,1)<>"/" then SavePath=SavePath&"/" '��Ŀ¼���(/)

upfilecount=0
set upfile=new clsUp ''�����ϴ�����
upfile.NoAllowExt=""&NoUpFileType&""	'�ϴ����͵ĺ�����
upfile.GetData (MaxFileSize)   'ȡ���ϴ�����,�����ļ���С
if upfile.fileSize>0 then
    if upfile.isErr then  '�������
    select case upfile.isErr
	case 1
	Response.Write "<script language='javascript'>alert('��û���ϴ�����ѽ???�ǲ��Ǹ����?');history.go(-1);</script>"
    response.end
		case 2
	Response.Write "<script language='javascript'>alert('���ϴ����ļ��������ǵ�����,���"&MaxFileSize/1000&"K');history.go(-1);</script>"
    response.end	
	end select
	Else


	FSPath=GetFilePath(Server.mappath("DNJY_upsave_ad.asp"),"\")'ȡ�õ�ǰ�ļ��ڷ�����·��
	ServerPath=GetFilePath(Request.ServerVariables("HTTP_REFERER"),"/")'ȡ������վ�ϵ�λ��
	for each formName in upfile.file '�г������ϴ��˵��ļ�
	set oFile=upfile.file(formname)

fileExt=lcase(ofile.FileExt)'����ļ���׺��
filename=SavePath&fupname&"."&fileExt'����ļ���������·��
upfile.SaveToFile formname,Server.mappath(filename)   '�����ļ� Ҳ����ʹ��AutoSave������,����һ��,���ǻ��Զ������µ��ļ���

If CheckFileExt(fileExt,UpFileType,blnAllow)=False then		
Response.Write "<script language='javascript'>alert('�����ļ����Ͳ������ϴ���\n\n"&strExt&"�ϴ��⼸���ļ����ͣ�" & UpFileType&"');history.go(-1);</script>"
response.end
end If

Set fso = CreateObject("Scripting.FileSystemObject")
if Right(filename,3)<>"swf" and Right(filename,3)<>"flv" then'20100722�ӣ��Դ����ļ������
if FSO.fileExists(Server.MapPath(""&filename&"")) then
whichfile=server.mappath(""& fileName & "")
set thisfile=fso.opentextfile(whichfile)
my_string=thisfile.readall
set thisfile=Nothing
If ChkImg(""&filename&"")=False or Chkfile(""&my_string&"")=False Then
muma=1
Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
if objFSO.fileExists(Server.MapPath(""&filename&"")) then
objFSO.DeleteFile(Server.MapPath(""&filename&""))
set objfso=nothing
end If
Response.Write "<script language='javascript'>alert('�ϴ�ʧ�ܣ�ԭ���ļ��������ռ�Ŀ¼û��дȨ�ޣ�������ͼ���ת����ʽ��༭���棬�����ÿռ�Ŀ¼дȨ�����Կ���');history.go(-1);</script>"
response.end
end If
Else
Response.Write "<script language='javascript'>alert('��ѡ�ļ��ϴ�ʲô��!');history.go(-1);</script>"
response.end
end If
end If'20100722��
	
	response.write "�ļ��ϴ��ɹ�!�ϴ��ļ�������·��Ϊ��"
	response.write "<font color=red>"&filename&"</font><br><br>"
	response.write "<a href='"&filename&"'  target='_blank'>���Ԥ���ϴ����ļ�</a>"
	response.write "<br><br><INPUT onclick='javascript:window.close();' type=submit value='�ϴ����'>"
%>
<script language=JavaScript>
function inserttext(){
<%
'����ttup�������жϻش�valueֵ��ʲô��
if ttup="adv" or ttup="flv" then
response.write "self.opener.form."&frmname&".value ='/"&strInstallDir&mid(filename,4)&"';"
elseif ttup="slide1" Or ttup="slide2" Or ttup="slide3" then
response.write "self.opener.form."&frmname&".value ='/"&strInstallDir&mid(filename,4)&"';"
elseif ttup="slideb" then
response.write "self.opener.form."&frmname&".value ='/"&strInstallDir&mid(filename,4)&"';"
elseif ttup="adfp" then
response.write "self.opener."&dform&"."&frmname&".value ='"&mid(filename,4)&"';"
elseif ttup="weblogo" then
response.write "self.opener.form1."&frmname&".value ='/"&strInstallDir&mid(filename,4)&"';"
end if
%>
}
inserttext();
</script>

<%
	set oFile=nothing
	Next
	end if
    set upfile=nothing  'ɾ���˶���
Else
response.write "<script language='javascript'>"
	response.write "alert('�ļ����ݲ���Ϊ�գ�������ȷ����������һҳ��');"
	response.write "location.href='javascript:history.go(-1)';"
	response.write "</script>"
	response.end
end if
set upload=nothing
session("ttup")=""
session("form")=""
session("fupname")=""
session("frmname")=""


function GetFilePath(FullPath,str)
  If FullPath <> "" Then
    GetFilePath = left(FullPath,InStrRev(FullPath, str))
    Else
    GetFilePath = ""
  End If
End function%>
</body> 
</html>