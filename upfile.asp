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
'天天供求信息网站管理系统
'天天网络科技工作室出品
'客服中心http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>
<%'======================================设置上传参数=====================================
Const MaxFileSize=500000  '上传文件大小限制，单位字节
Const UpFileType="jpg;jpeg;gif;bmp;png;"    '白名单,在这里设可上传的文件类型,以文件的后缀名来判断,不分大小写,每个后缀名用;号分开
Const NoUpFileType = "asp;asa;exe;htm;html;aspx;cs;vb;js;dll;php;jsp;cgi;" '黑名单,在这里设不可上传的文件类型,以文件的后缀名来判断,不分大小写,每个每缀名用;号分开,如果黑名单为空,则判断白名
Const blnAllow=True           '是只允许还是禁止上传 UpFileType 中指定的后缀,True允许，false禁止
SavePath = "UploadFiles/infopic"   '存放上传信息主图片文件的目录


'======================================以下代码不要动！=====================================
Call OpenConn
if right(SavePath,1)<>"/" then SavePath=SavePath&"/" '在目录后加(/)

dim file,iCount,imgpath,SavePath,upfile,formPath,ServerPath,FSPath,formName,FileName,oFile,upfilecount,ttup,upname,xxmpname,strExt,fileExt,ranNum,my_string,objFSO,muma,msg,strJS,fso,whichfile,thisfile,username,user_c,tupian,vip,tpname,wcolor,ncolor 

username=trim(request("username"))
vip=trim(request("vip"))
id=trim(request("id"))
'id=upload.form("id")
iCount=0
'set upload = new sjCat_Upload
'Set file = upload.file("file1")
 
'===========上传文件并检测==============
upfilecount=0
set upfile=new clsUp ''建立上传对象
upfile.NoAllowExt=""&NoUpFileType&""	'上传类型的黑名单
upfile.GetData (MaxFileSize)   '取得上传数据,限制文件大小

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

	FSPath=GetFilePath(Server.mappath("upfile.asp"),"\")'取得当前文件在服务器路径
	ServerPath=GetFilePath(Request.ServerVariables("HTTP_REFERER"),"/")'取得在网站上的位置
	for each formName in upfile.file '列出所有上传了的文件
	set oFile=upfile.file(formname)

fileExt=lcase(ofile.FileExt)
randomize'初始化随机数发生器
ranNum=int(900*rnd)+100
tpname=year(now) & right("0" & month(now),2) & right("0" & day(now),2) & right("0" & hour(now),2) & right("0" & minute(now),2) & right("0" & second(now),2)&ranNum&"."&fileExt'获得日期时间文件名
filename=SavePath&tpname

upfile.SaveToFile formname,FSPath&FileName   '保存文件 也可以使用AutoSave来保存,参数一样,但是会自动建立新的文件名

fileExt=lcase(ofile.FileExt)		
If CheckFileExt(fileExt,UpFileType,blnAllow)=False then		
Response.Write "<script language='javascript'>alert('这种文件类型不允许上传！\n\n"&strExt&"上传这几种文件类型：" & UpFileType&"');history.go(-1);</script>"
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
response.end
end If
Else
Response.Write "<script language='javascript'>alert('不选文件上传什么啊!');history.go(-1);</script>"
response.end
end If
 
	set oFile=nothing
	Next
	end if
    set upfile=nothing  '删除此对象
    set file=nothing
iCount=iCount+1


'=========上传文件检测结束＝＝＝＝＝＝


set rs=server.createobject("adodb.recordset")
sql = "select tupian,c,Readinfo from DNJY_info where id="&cstr(id)
rs.open sql,conn,1,3
tupian=rs("tupian")
Readinfo=rs("Readinfo")
if rs.eof then
response.write "<script language=JavaScript>" & chr(13) & "alert('参数错误！');" &"history.back()'" & "</script>"
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
Conn.Execute("Update [DNJY_user] Set c=c-1 where username='"&username&"'")'减少一个贴图道具
end if
Call CloseRs(rs)


response.write "<script language=JavaScript>" & chr(13) & "alert('图片成功上传！');" &"window.location='user_xxgl.asp?"&C_ID&"'" & "</script>"

function GetFilePath(FullPath,str)
  If FullPath <> "" Then
    GetFilePath = left(FullPath,InStrRev(FullPath, str))
    Else
    GetFilePath = ""
  End If
End Function

Call CloseDB(conn)%><br>