<!--#include file="conn.asp"-->
<!--#include file="config.asp"-->
<!--#include file="Include/Function.asp"-->
<!--#include file="Include/CheckFile.asp"-->
<!--#include file="DNJY_clsUp.asp"-->
<%
'=====================================
'���칩����Ϣ��վ����ϵͳ
'��������Ƽ������ҳ�Ʒ
'�ͷ�����http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>
<script LANGUAGE="Javascript">
<!--
setTimeout('window.close();', 3000); //3���ر�
// -->
</script>
<%'======================================�����ϴ�����=====================================
Const MaxFileSize=5000000  '�ϴ��ļ���С���ƣ���λ�ֽ�
Const UpFileType="jpg;jpeg;gif;bmp;png;"    '������,����������ϴ����ļ�����,���ļ��ĺ�׺�����ж�,���ִ�Сд,ÿ����׺����;�ŷֿ�
Const NoUpFileType = "asp;asa;exe;htm;html;aspx;cs;vb;js;dll;php;jsp;cgi;" '������,�������費���ϴ����ļ�����,���ļ��ĺ�׺�����ж�,���ִ�Сд,ÿ��ÿ׺����;�ŷֿ�,���������Ϊ��,���жϰ���
Const blnAllow=True           '��ֻ�����ǽ�ֹ�ϴ� UpFileType ��ָ���ĺ�׺,True����false��ֹ


'======================================���´��벻Ҫ����=====================================
dim upfile,formPath,SavePath,ServerPath,FSPath,formName,FileName,oFile,upfilecount,ttup,upname,strExt,fileExt,ranNum,my_string,objFSO,muma,msg,strJS,fso,whichfile,thisfile,xxname,id,oneid,twoid,threeid
'Const DelUpFiles="Yes"        'ɾ������ʱ�Ƿ�ͬʱɾ�������е��ϴ��ļ�
ttup=Request("ttup")          '�������,�ϴ�������
	id=strint(request("id"))
	oneid=strint(request("oneid"))
	twoid=strint(request("twoid"))
	threeid=strint(request("threeid"))
If blnAllow=True Then
strExt="ֻ����"
Else
strExt="��ֹ"
End If


If ttup="upnews" Then
SavePath = "UploadFiles/upnews"     '����ϴ�������ͼƬ�ļ���Ŀ¼
upname="pic_url"'�ش���������
elseIf ttup="photo" Then
SavePath = "UploadFiles/upnews"     '����ϴ���������ļ���Ŀ¼
upname=request("frmname")'�ش���������
elseIf ttup="link" Then
SavePath = "UploadFiles/link"       '������������ϴ��ļ���Ŀ¼
upname="logo"
ElseIf ttup="sdlogo" Then
SavePath = "UploadFiles/sdlogo"     '��ŵ��ͼƬ�ϴ��ļ���Ŀ¼
upname="tupian"
ElseIf ttup="sdBanner" Then
SavePath = "UploadFiles/sdlogo"     '���Banner�ϴ��ļ���Ŀ¼
upname="sdBanner"
ElseIf ttup="info" Then
SavePath = "UploadFiles/infopic"   '��������Ϣ�ϴ���ͼ��Ŀ¼
upname="tpname"
ElseIf ttup="thumbnail" Then
SavePath = "UploadFiles/magazine"   '��ŵ�����־�ϴ�ͼƬ��Ŀ¼
upname="thumbnail"
ElseIf ttup="largerpic" Then
SavePath = "UploadFiles/magazine"   '��ŵ�����־�ϴ�ͼƬ��Ŀ¼
upname="largerpic"
ElseIf ttup="certificate1" Then
SavePath = "UploadFiles/certificate"     '��ŵ�����Ƭ���֤���ϴ��ļ���Ŀ¼
upname="certificate1"
ElseIf ttup="certificate2" Then
SavePath = "UploadFiles/certificate"     '��ŵ�����Ƭ���֤���ϴ��ļ���Ŀ¼
upname="certificate2"
ElseIf ttup="certificate3" Then
SavePath = "UploadFiles/certificate"     '��ŵ�����Ƭ���֤���ϴ��ļ���Ŀ¼
upname="certificate3"
ElseIf ttup="certificate4" Then
SavePath = "UploadFiles/certificate"     '��ŵ�����Ƭ���֤���ϴ��ļ���Ŀ¼
upname="certificate4"
ElseIf ttup="certificate5" Then
SavePath = "UploadFiles/certificate"     '��ŵ�����Ƭ���֤���ϴ��ļ���Ŀ¼
upname="certificate5"
ElseIf ttup="certificate6" Then
SavePath = "UploadFiles/certificate"     '��ŵ�����Ƭ���֤���ϴ��ļ���Ŀ¼
upname="certificate6"
ElseIf ttup="certificate7" Then
SavePath = "UploadFiles/certificate"     '��ŵ�����Ƭ���֤���ϴ��ļ���Ŀ¼
upname="certificate7"
Else
Response.Write "<script language='javascript'>alert('����Ĭ�ϵ��ϴ�ѡ������ϴ�!');history.go(-1);</script>"
response.end
End if
if right(SavePath,1)<>"/" then SavePath=SavePath&"/" '��Ŀ¼���(/)
upfilecount=0
set upfile=new clsUp ''�����ϴ�����
upfile.NoAllowExt=""&NoUpFileType&""	'�ϴ����͵ĺ�����
upfile.GetData (MaxFileSize)   'ȡ���ϴ�����,�����ļ���С

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


	FSPath=GetFilePath(Server.mappath("DNJY_upfile.asp"),"\")'ȡ�õ�ǰ�ļ��ڷ�����·��
	ServerPath=GetFilePath(Request.ServerVariables("HTTP_REFERER"),"/")'ȡ������վ�ϵ�λ��
	for each formName in upfile.file '�г������ϴ��˵��ļ�
	set oFile=upfile.file(formname)

fileExt=lcase(ofile.FileExt)
randomize'��ʼ�������������
ranNum=int(900*rnd)+100
filename=SavePath&year(now) & right("0" & month(now),2) & right("0" & day(now),2) & right("0" & hour(now),2) & right("0" & minute(now),2) & right("0" & second(now),2)&ranNum&"."&fileExt'�������ʱ���ļ���
xxname=year(now) & right("0" & month(now),2) & right("0" & day(now),2) & right("0" & hour(now),2) & right("0" & minute(now),2) & right("0" & second(now),2)&ranNum&"."&fileExt'�����ϢͼƬ����ʱ���ļ��� 
upfile.SaveToFile formname,FSPath&FileName   '�����ļ� Ҳ����ʹ��AutoSave������,����һ��,���ǻ��Զ������µ��ļ���
		
If CheckFileExt(fileExt,UpFileType,blnAllow)=False then		
Response.Write "<script language='javascript'>alert('�����ļ����Ͳ������ϴ���\n\n"&strExt&"�ϴ��⼸���ļ����ͣ�" & UpFileType&"');history.go(-1);</script>"
response.end
end If

Set fso = CreateObject("Scripting.FileSystemObject")
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
Response.Write "<script language='javascript'>alert('�ϴ�ʧ�ܣ�ԭ���ļ������⣡������ͼ���ת����ʽ��༭�������ϴ���');</script>"
msg="��ѡ����ȷ�ļ�"
end If
Else
Response.Write "<script language='javascript'>alert('��ѡ�ļ��ϴ�ʲô��!');history.go(-1);</script>"
response.end
end If

strJS="<SCRIPT language=javascript>" & vbcrlf	   
if muma<>1 then			
Response.Write "<script language='javascript'>alert('ͼƬ�ϴ��ɹ���');</script>"
Response.Write "<a href=javascript:window.close()>�رձ�����</a>"



If ttup="info" Then fileName=xxName
if ttup="photo" then
strJS=strJS & "self.opener.addNEWS."&upname&".value ='"&filename&"';"
else
strJS=strJS & "parent.document.forms[0]."&upname&".value='"& fileName & "';" & vbcrlf    '�����ϴ�ͼƬ��ַ,����Ϊupname
end if
Else
strJS="<SCRIPT language=javascript>" & vbcrlf
strJS=strJS & "alert('" & msg & "');" & vbcrlf
strJS=strJS & "history.go(-1);" & vbcrlf
end if
strJS=strJS & "</script>" & vbcrlf
response.write strJS


	set oFile=nothing
	Next
	end if
    set upfile=nothing  'ɾ���˶���


function GetFilePath(FullPath,str)
  If FullPath <> "" Then
    GetFilePath = left(FullPath,InStrRev(FullPath, str))
    Else
    GetFilePath = ""
  End If
End function%>
