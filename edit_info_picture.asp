<!--#include file="conn.asp"-->
<!--#include file=usercookies.asp-->
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
<link href="/<%=strInstallDir%>css/css.css" rel="stylesheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<%'======================================�����ϴ�����=====================================
Const MaxFileSize=100000  '�ϴ��ļ���С���ƣ���λ�ֽ�
Const UpFileType="jpg;jpeg;gif;bmp;"    '������,����������ϴ����ļ�����,���ļ��ĺ�׺�����ж�,���ִ�Сд,ÿ����׺����;�ŷֿ�
Const NoUpFileType = "asp;asa;exe;htm;html;aspx;cs;vb;js;dll;php;jsp;cgi;" '������,�������費���ϴ����ļ�����,���ļ��ĺ�׺�����ж�,���ִ�Сд,ÿ��ÿ׺����;�ŷֿ�,���������Ϊ��,���жϰ���
Const blnAllow=True           '��ֻ�����ǽ�ֹ�ϴ� UpFileType ��ָ���ĺ�׺,True����false��ֹ
SavePath = "UploadFiles/infopic"   '����ϴ���ϢͼƬ�ļ���Ŀ¼


'======================================���´��벻Ҫ����=====================================

if right(SavePath,1)<>"/" then SavePath=SavePath&"/" '��Ŀ¼���(/)
Dim SavePath,upfile,formPath,ServerPath,FSPath,formName,FileName,oFile,upfilecount,ttup,upname,xxmpname,strExt,fileExt,ranNum,my_string,objFSO,muma,msg,strJS,fso,whichfile,thisfile,username,user_c,cc,tupian,vip,tpname,wcolor,ncolor

i=1
id=trim(request("id"))
username=request.cookies("dick")("username")
Call OpenConn
set rs = Server.CreateObject("ADODB.RecordSet")
sql="select c,vip from [DNJY_user] where username='"&username&"'"
rs.open sql,conn,1,1
if rs.eof then
response.write "��������"
Call CloseRs(rs)
Call CloseDB(conn)
response.end
end If
vip=rs("vip")
if request("dick")<>"chk" then
user_c=rs("c")
if user_c<=0 then
response.write "����ͼ����Ϊ"&user_c&"���������ϴ�ͼƬ�����ȹ�����ߣ�"
cl
Call CloseRs(rs)
Call CloseDB(conn)
response.end
end if
Call CloseRs(rs)
set rs = Server.CreateObject("ADODB.RecordSet")
sql="select c,tupian from [DNJY_info] where username='"&username&"' and id="&cstr(id)
rs.open sql,conn,1,1
if rs.eof then
response.write "��������"
cl
Call CloseRs(rs)
Call CloseDB(conn)
response.end
end if
cc=rs("c")
if cc=1 then
tupian=trim(rs("tupian"))
else
tupian="0"
end if
Call CloseRs(rs)
%>
<meta http-equiv="Content-Language" content="zh-cn">
<link rel="stylesheet" type="text/css" href="/<%=strInstallDir%>css/css.css">
<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%">
  <tr>
    <td width="100%">
    <p align="center">����Ϣ<%if cc=1 then%><a target=_blank href="UploadFiles/infopic/<%=tupian%>"><font color=#0000ff><b>������ͼ</b></font></a>���ϴ�ͼƬ<font color=#FF0000>��ֱ�Ӹ���ԭͼ</font><%else%>û����ͼ�����ϴ�ͼƬ<%end if%>��gif,jpg����Ƚ���500�����ڣ�</td>
  </tr>
</table>
<form name="form1" method="post" action="?dick=chk&id=<%=id%>&cc=<%if cc=1 then%>1<%else%>0<%end if%>" enctype="multipart/form-data">
<input type="hidden" name="id" value="<%=id%>">
<input type="hidden" name="cc" value="<%if cc=1 then%>1<%else%>0<%end if%>">
<p align="center">
<input type="file" name="file1" value="" size="17"><br>
<input type="submit" value="�ϴ�ͼƬ" name="B1"><br>
˵���������ϴ����޸Ķ���۳�<font color="#0000FF">1</font>��<font color="#0000FF">��ͼ����</font>.����������<font color="#FF0000"><%=user_c%></font>��</p>
</form>

<%else%>
<%Server.ScriptTimeOut=5000%>
<%dim upload,file,iCount,c_yn
iCount=0
id=cstr(request("id"))
c_yn=CInt(request("cc"))
username=request.cookies("dick")("username")

upfilecount=0
set upfile=new clsUp ''�����ϴ�����
upfile.NoAllowExt=""&NoUpFileType&""	'�ϴ����͵ĺ�����
upfile.GetData (MaxFileSize)   'ȡ���ϴ�����,�����ļ���С

    if upfile.isErr then  '�������
    select case upfile.isErr
	case 1
	Response.Write "<script language='javascript'>alert('��û���ϴ�����ѽ???�ǲ��Ǹ����?');history.go(-1);</script>"
    Call CloseDB(conn)
    response.end
		case 2
	Response.Write "<script language='javascript'>alert('���ϴ����ļ��������ǵ�����,���"&MaxFileSize/1000&"K');history.go(-1);</script>"
    Call CloseDB(conn)
    response.end	
	end select
	Else

	FSPath=GetFilePath(Server.mappath("edit_info_picture.asp"),"\")'ȡ�õ�ǰ�ļ��ڷ�����·��
	ServerPath=GetFilePath(Request.ServerVariables("HTTP_REFERER"),"/")'ȡ������վ�ϵ�λ��
	for each formName in upfile.file '�г������ϴ��˵��ļ�
	set oFile=upfile.file(formname)

fileExt=lcase(ofile.FileExt)
randomize'��ʼ�������������
ranNum=int(900*rnd)+100
tpname=year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)&ranNum&"."&fileExt'�������ʱ���ļ���
filename=SavePath&tpname

upfile.SaveToFile formname,FSPath&FileName   '�����ļ� Ҳ����ʹ��AutoSave������,����һ��,���ǻ��Զ������µ��ļ���

fileExt=lcase(ofile.FileExt)		
If CheckFileExt(fileExt,UpFileType,blnAllow)=False then		
Response.Write "<script language='javascript'>alert('�����ļ����Ͳ������ϴ���\n\n"&strExt&"�ϴ��⼸���ļ����ͣ�" & UpFileType&"');history.go(-1);</script>"
Call CloseDB(conn)
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
Call CloseDB(conn)
response.end
end If
Else
Response.Write "<script language='javascript'>alert('��ѡ�ļ��ϴ�ʲô��!');history.go(-1);</script>"
Call CloseDB(conn)
response.end
end If
 
	set oFile=nothing
	Next
	end if
    set upfile=nothing  'ɾ���˶���
    set file=nothing
iCount=iCount+1
set rs=server.createobject("adodb.recordset")
sql = "select tupian,c from DNJY_info where username='"&username&"' and id="&cstr(id) '����
rs.open sql,conn,1,3
if rs.eof then
response.write "��������"
cl
Call CloseRs(rs)
Call CloseDB(conn)
response.end
else
 if c_yn="1" then
 Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
 if objFSO.fileExists(Server.MapPath("UploadFiles\infopic\"&rs("tupian"))) then
 objFSO.DeleteFile(Server.MapPath("UploadFiles\infopic\"&rs("tupian")))
 end if
 set objFSO=nothing
 end If 
rs("tupian")=tpname
rs("c")=1 '����
rs.update
end If



Conn.Execute("Update [DNJY_user] Set c=c-1 where username='"&username&"'")'����һ����ͼ����
response.write "ͼƬ�ɹ��ϴ���"
cl
Call CloseDB(conn)
response.end
end if
sub cl()%>
<body onLoad="setTimeout(window.close, 2000)">
<%end Sub


function GetFilePath(FullPath,str)
  If FullPath <> "" Then
    GetFilePath = left(FullPath,InStrRev(FullPath, str))
    Else
    GetFilePath = ""
  End If
End Function
Call CloseDB(conn)%>