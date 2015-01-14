<!--#include file="../conn.asp"-->
<!--#include file="../config.asp"-->
<!--#include file="../Include/CheckFile.asp"-->
<!--#include file="../DNJY_clsUp.asp"-->
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
end if%>
<html>
<head>
<title>文件上传</title>
<meta name="Description" Content="">
<link href="../css/style.css" rel="stylesheet" type="text/css">
</head>
<body>
<%'======================================设置上传参数=====================================
Dim FileType,UpFileType
Const MaxFileSize=5000000  '上传文件大小限制，单位字节
ttup=session("ttup")
if ttup="slide1" Or ttup="slide2" Or ttup="slide3" Or ttup="slideb" Or ttup="adfp" Then
FileType="jpg;jpeg;"    '白名单,在这里设可上传的文件类型,以文件的后缀名来判断,不分大小写,每个后缀名用;号分开
elseif ttup="flv" Then
FileType="flv;FLV;"    '白名单,在这里设可上传的文件类型,以文件的后缀名来判断,不分大小写,每个后缀名用;号分开
Else
FileType="jpg;jpeg;gif;bmp;swf;png;"    '白名单,在这里设可上传的文件类型,以文件的后缀名来判断,不分大小写,每个后缀名用;号分开
End if
UpFileType=FileType    '白名单,在这里设可上传的文件类型,以文件的后缀名来判断,不分大小写,每个后缀名用;号分开
Const NoUpFileType = "asp;asa;exe;htm;html;aspx;cs;vb;js;dll;php;jsp;cgi;" '黑名单,在这里设不可上传的文件类型,以文件的后缀名来判断,不分大小写,每个每缀名用;号分开,如果黑名单为空,则判断白名
Const blnAllow=True           '是只允许还是禁止上传 UpFileType 中指定的后缀,True允许，false禁止


'======================================以下代码不要动！=====================================
Call OpenConn
On Error Resume Next
dim fupname,frmname
Dim upfile,formPath,SavePath,ServerPath,FSPath,formName,FileName,oFile,upfilecount,ttup,upname,strExt,fileExt,ranNum,my_string,objFSO,muma,msg,strJS,fso,whichfile,thisfile,dform

If blnAllow=True Then
strExt="只允许"
Else
strExt="禁止"
End If
ttup=session("ttup")
dform=session("dform")
fupname=session("fupname")  '传递文件名
frmname=session("frmname")	'传递表单名
if  fupname="" or frmname="" then
	response.write "<script language='javascript'>"
	response.write "alert('出现错误，请重新上传！');"
	response.write "location.href='javascript:history.go(-1)';"
	response.write "</script>"
	response.end
end if

if ttup="adv" then
savepath="../ads/pic" 	'标准广告图片
elseif ttup="slide1" Or ttup="slide2" Or ttup="slide3" then
savepath="../UploadFiles/slidea" '幻灯广告A
elseif ttup="slideb" then
savepath="../UploadFiles/slideb" '幻灯广告B
elseif ttup="adfp" then
savepath="../UploadFiles/adfp" '翻牌广告
elseif ttup="flv" then
savepath="../UploadFiles/flv" 'flv广告
elseif ttup="weblogo" then
savepath="images" '备用
Else
Response.Write "<script language='javascript'>alert('不是默认的上传选项，不能上传!');history.go(-1);</script>"
response.end
end If
if right(SavePath,1)<>"/" then SavePath=SavePath&"/" '在目录后加(/)

upfilecount=0
set upfile=new clsUp ''建立上传对象
upfile.NoAllowExt=""&NoUpFileType&""	'上传类型的黑名单
upfile.GetData (MaxFileSize)   '取得上传数据,限制文件大小
if upfile.fileSize>0 then
    if upfile.isErr then  '如果出错
    select case upfile.isErr
	case 1
	Response.Write "<script language='javascript'>alert('你没有上传数据呀???是不是搞错了?');history.go(-1);</script>"
    response.end
		case 2
	Response.Write "<script language='javascript'>alert('你上传的文件超出我们的限制,最大"&MaxFileSize/1000&"K');history.go(-1);</script>"
    response.end	
	end select
	Else


	FSPath=GetFilePath(Server.mappath("DNJY_upsave_ad.asp"),"\")'取得当前文件在服务器路径
	ServerPath=GetFilePath(Request.ServerVariables("HTTP_REFERER"),"/")'取得在网站上的位置
	for each formName in upfile.file '列出所有上传了的文件
	set oFile=upfile.file(formname)

fileExt=lcase(ofile.FileExt)'获得文件后缀名
filename=SavePath&fupname&"."&fileExt'获得文件名及保存路径
upfile.SaveToFile formname,Server.mappath(filename)   '保存文件 也可以使用AutoSave来保存,参数一样,但是会自动建立新的文件名

If CheckFileExt(fileExt,UpFileType,blnAllow)=False then		
Response.Write "<script language='javascript'>alert('这种文件类型不允许上传！\n\n"&strExt&"上传这几种文件类型：" & UpFileType&"');history.go(-1);</script>"
response.end
end If

Set fso = CreateObject("Scripting.FileSystemObject")
if Right(filename,3)<>"swf" and Right(filename,3)<>"flv" then'20100722加，对此类文件不检测
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
Response.Write "<script language='javascript'>alert('上传失败，原因：文件有问题或空间目录没有写权限！请用制图软件转换格式或编辑保存，或设置空间目录写权限再试看。');history.go(-1);</script>"
response.end
end If
Else
Response.Write "<script language='javascript'>alert('不选文件上传什么啊!');history.go(-1);</script>"
response.end
end If
end If'20100722加
	
	response.write "文件上传成功!上传文件的物理路径为："
	response.write "<font color=red>"&filename&"</font><br><br>"
	response.write "<a href='"&filename&"'  target='_blank'>点击预览上传的文件</a>"
	response.write "<br><br><INPUT onclick='javascript:window.close();' type=submit value='上传完成'>"
%>
<script language=JavaScript>
function inserttext(){
<%
'根据ttup的类型判断回传value值到什么表单
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
    set upfile=nothing  '删除此对象
Else
response.write "<script language='javascript'>"
	response.write "alert('文件内容不能为空，单击“确定”返回上一页！');"
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