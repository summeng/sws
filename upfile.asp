<!--#include file="conn.asp"-->
<!--#include file=usercookies.asp-->
<%Server.ScriptTimeOut=5000%>
<!--#include file="config.asp"-->
<!--#include file="Include/head.asp"-->
<!--#include file="Include/Function.asp"-->
<!--#include file="Include/label.asp"-->
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
<%'======================================�����ϴ�����=====================================
Const MaxFileSize=500000  '�ϴ��ļ���С���ƣ���λ�ֽ�
Const UpFileType="jpg;jpeg;gif;bmp;png;"    '������,����������ϴ����ļ�����,���ļ��ĺ�׺�����ж�,���ִ�Сд,ÿ����׺����;�ŷֿ�
Const NoUpFileType = "asp;asa;exe;htm;html;aspx;cs;vb;js;dll;php;jsp;cgi;" '������,�������費���ϴ����ļ�����,���ļ��ĺ�׺�����ж�,���ִ�Сд,ÿ��ÿ׺����;�ŷֿ�,���������Ϊ��,���жϰ���
Const blnAllow=True           '��ֻ�����ǽ�ֹ�ϴ� UpFileType ��ָ���ĺ�׺,True����false��ֹ
SavePath = "UploadFiles/infopic"   '����ϴ���Ϣ��ͼƬ�ļ���Ŀ¼


'======================================���´��벻Ҫ����=====================================
Call OpenConn
if right(SavePath,1)<>"/" then SavePath=SavePath&"/" '��Ŀ¼���(/)

dim file,iCount,imgpath,SavePath,upfile,formPath,ServerPath,FSPath,formName,FileName,oFile,upfilecount,ttup,upname,xxmpname,strExt,fileExt,ranNum,my_string,objFSO,muma,msg,strJS,fso,whichfile,thisfile,username,user_c,tupian,vip,tpname,wcolor,ncolor 

username=trim(request("username"))
vip=trim(request("vip"))
id=trim(request("id"))
'id=upload.form("id")
iCount=0
'set upload = new sjCat_Upload
'Set file = upload.file("file1")
 
'===========�ϴ��ļ������==============
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

	FSPath=GetFilePath(Server.mappath("upfile.asp"),"\")'ȡ�õ�ǰ�ļ��ڷ�����·��
	ServerPath=GetFilePath(Request.ServerVariables("HTTP_REFERER"),"/")'ȡ������վ�ϵ�λ��
	for each formName in upfile.file '�г������ϴ��˵��ļ�
	set oFile=upfile.file(formname)

fileExt=lcase(ofile.FileExt)
randomize'��ʼ�������������
ranNum=int(900*rnd)+100
tpname=year(now) & right("0" & month(now),2) & right("0" & day(now),2) & right("0" & hour(now),2) & right("0" & minute(now),2) & right("0" & second(now),2)&ranNum&"."&fileExt'�������ʱ���ļ���
filename=SavePath&tpname

upfile.SaveToFile formname,FSPath&FileName   '�����ļ� Ҳ����ʹ��AutoSave������,����һ��,���ǻ��Զ������µ��ļ���

fileExt=lcase(ofile.FileExt)		
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
Response.Write "<script language='javascript'>alert('���ϴ����ļ������⣡�ϴ�ʧ�ܡ�������ͼ���ת����ʽ��༭�������Կ���');history.go(-1);</script>"
response.end
end If
Else
Response.Write "<script language='javascript'>alert('��ѡ�ļ��ϴ�ʲô��!');history.go(-1);</script>"
response.end
end If
 
	set oFile=nothing
	Next
	end if
    set upfile=nothing  'ɾ���˶���
    set file=nothing
iCount=iCount+1


'=========�ϴ��ļ�������������������


set rs=server.createobject("adodb.recordset")
sql = "select tupian,c,Readinfo from DNJY_info where id="&cstr(id)
rs.open sql,conn,1,3
tupian=rs("tupian")
Readinfo=rs("Readinfo")
if rs.eof then
response.write "<script language=JavaScript>" & chr(13) & "alert('��������');" &"history.back()'" & "</script>"
response.end
else
rs("tupian")=tpname
rs("c")=1
rs.update


imgpath="UploadFiles/infopic/" & tupian 
Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
if objFSO.fileExists(Server.MapPath(""&imgpath&"")) then
objFSO.DeleteFile(Server.MapPath(""&imgpath&""))
set objfso=Nothing
End if
Conn.Execute("Update [DNJY_user] Set c=c-1 where username='"&username&"'")'����һ����ͼ����
end if
Call CloseRs(rs)


response.write "<script language=JavaScript>" & chr(13) & "alert('ͼƬ�ɹ��ϴ���');" &"window.location='user_xxgl.asp?"&C_ID&"'" & "</script>"

function GetFilePath(FullPath,str)
  If FullPath <> "" Then
    GetFilePath = left(FullPath,InStrRev(FullPath, str))
    Else
    GetFilePath = ""
  End If
End Function

Call CloseDB(conn)%><br>