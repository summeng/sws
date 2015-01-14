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
'天天供求信息网站管理系统
'天天网络科技工作室出品
'客服中心http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>
<link href="/<%=strInstallDir%>css/css.css" rel="stylesheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<%'======================================设置上传参数=====================================
Const MaxFileSize=100000  '上传文件大小限制，单位字节
Const UpFileType="jpg;jpeg;gif;bmp;"    '白名单,在这里设可上传的文件类型,以文件的后缀名来判断,不分大小写,每个后缀名用;号分开
Const NoUpFileType = "asp;asa;exe;htm;html;aspx;cs;vb;js;dll;php;jsp;cgi;" '黑名单,在这里设不可上传的文件类型,以文件的后缀名来判断,不分大小写,每个每缀名用;号分开,如果黑名单为空,则判断白名
Const blnAllow=True           '是只允许还是禁止上传 UpFileType 中指定的后缀,True允许，false禁止
SavePath = "UploadFiles/infopic"   '存放上传信息图片文件的目录


'======================================以下代码不要动！=====================================

if right(SavePath,1)<>"/" then SavePath=SavePath&"/" '在目录后加(/)
Dim SavePath,upfile,formPath,ServerPath,FSPath,formName,FileName,oFile,upfilecount,ttup,upname,xxmpname,strExt,fileExt,ranNum,my_string,objFSO,muma,msg,strJS,fso,whichfile,thisfile,username,user_c,cc,tupian,vip,tpname,wcolor,ncolor

i=1
id=trim(request("id"))
username=request.cookies("dick")("username")
Call OpenConn
set rs = Server.CreateObject("ADODB.RecordSet")
sql="select c,vip from [DNJY_user] where username='"&username&"'"
rs.open sql,conn,1,1
if rs.eof then
response.write "参数错误！"
Call CloseRs(rs)
Call CloseDB(conn)
response.end
end If
vip=rs("vip")
if request("dick")<>"chk" then
user_c=rs("c")
if user_c<=0 then
response.write "你贴图道具为"&user_c&"个，不能上传图片，请先购买道具！"
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
response.write "参数错误！"
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
    <p align="center">该信息<%if cc=1 then%><a target=_blank href="UploadFiles/infopic/<%=tupian%>"><font color=#0000ff><b>已有贴图</b></font></a>，上传图片<font color=#FF0000>将直接覆盖原图</font><%else%>没有贴图，可上传图片<%end if%>（gif,jpg，宽度建议500像素内）</td>
  </tr>
</table>
<form name="form1" method="post" action="?dick=chk&id=<%=id%>&cc=<%if cc=1 then%>1<%else%>0<%end if%>" enctype="multipart/form-data">
<input type="hidden" name="id" value="<%=id%>">
<input type="hidden" name="cc" value="<%if cc=1 then%>1<%else%>0<%end if%>">
<p align="center">
<input type="file" name="file1" value="" size="17"><br>
<input type="submit" value="上传图片" name="B1"><br>
说明：不论上传或修改都会扣除<font color="#0000FF">1</font>个<font color="#0000FF">贴图道具</font>.道具数量：<font color="#FF0000"><%=user_c%></font>个</p>
</form>

<%else%>
<%Server.ScriptTimeOut=5000%>
<%dim upload,file,iCount,c_yn
iCount=0
id=cstr(request("id"))
c_yn=CInt(request("cc"))
username=request.cookies("dick")("username")

upfilecount=0
set upfile=new clsUp ''建立上传对象
upfile.NoAllowExt=""&NoUpFileType&""	'上传类型的黑名单
upfile.GetData (MaxFileSize)   '取得上传数据,限制文件大小

    if upfile.isErr then  '如果出错
    select case upfile.isErr
	case 1
	Response.Write "<script language='javascript'>alert('你没有上传数据呀???是不是搞错了?');history.go(-1);</script>"
    Call CloseDB(conn)
    response.end
		case 2
	Response.Write "<script language='javascript'>alert('你上传的文件超出我们的限制,最大"&MaxFileSize/1000&"K');history.go(-1);</script>"
    Call CloseDB(conn)
    response.end	
	end select
	Else

	FSPath=GetFilePath(Server.mappath("edit_info_picture.asp"),"\")'取得当前文件在服务器路径
	ServerPath=GetFilePath(Request.ServerVariables("HTTP_REFERER"),"/")'取得在网站上的位置
	for each formName in upfile.file '列出所有上传了的文件
	set oFile=upfile.file(formname)

fileExt=lcase(ofile.FileExt)
randomize'初始化随机数发生器
ranNum=int(900*rnd)+100
tpname=year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)&ranNum&"."&fileExt'获得日期时间文件名
filename=SavePath&tpname

upfile.SaveToFile formname,FSPath&FileName   '保存文件 也可以使用AutoSave来保存,参数一样,但是会自动建立新的文件名

fileExt=lcase(ofile.FileExt)		
If CheckFileExt(fileExt,UpFileType,blnAllow)=False then		
Response.Write "<script language='javascript'>alert('这种文件类型不允许上传！\n\n"&strExt&"上传这几种文件类型：" & UpFileType&"');history.go(-1);</script>"
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
Response.Write "<script language='javascript'>alert('您上传的文件有问题！上传失败。请用制图软件转换格式或编辑保存再试看。');history.go(-1);</script>"
Call CloseDB(conn)
response.end
end If
Else
Response.Write "<script language='javascript'>alert('不选文件上传什么啊!');history.go(-1);</script>"
Call CloseDB(conn)
response.end
end If
 
	set oFile=nothing
	Next
	end if
    set upfile=nothing  '删除此对象
    set file=nothing
iCount=iCount+1
set rs=server.createobject("adodb.recordset")
sql = "select tupian,c from DNJY_info where username='"&username&"' and id="&cstr(id) '玉林
rs.open sql,conn,1,3
if rs.eof then
response.write "参数错误！"
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
rs("c")=1 '玉林
rs.update
end If



Conn.Execute("Update [DNJY_user] Set c=c-1 where username='"&username&"'")'减少一个贴图道具
response.write "图片成功上传！"
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