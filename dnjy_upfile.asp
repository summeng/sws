<!--#include file="conn.asp"-->
<!--#include file="config.asp"-->
<!--#include file="Include/Function.asp"-->
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
<script LANGUAGE="Javascript">
<!--
setTimeout('window.close();', 3000); //3秒后关闭
// -->
</script>
<%'======================================设置上传参数=====================================
Const MaxFileSize=5000000  '上传文件大小限制，单位字节
Const UpFileType="jpg;jpeg;gif;bmp;png;"    '白名单,在这里设可上传的文件类型,以文件的后缀名来判断,不分大小写,每个后缀名用;号分开
Const NoUpFileType = "asp;asa;exe;htm;html;aspx;cs;vb;js;dll;php;jsp;cgi;" '黑名单,在这里设不可上传的文件类型,以文件的后缀名来判断,不分大小写,每个每缀名用;号分开,如果黑名单为空,则判断白名
Const blnAllow=True           '是只允许还是禁止上传 UpFileType 中指定的后缀,True允许，false禁止


'======================================以下代码不要动！=====================================
dim upfile,formPath,SavePath,ServerPath,FSPath,formName,FileName,oFile,upfilecount,ttup,upname,strExt,fileExt,ranNum,my_string,objFSO,muma,msg,strJS,fso,whichfile,thisfile,xxname,id,oneid,twoid,threeid
'Const DelUpFiles="Yes"        '删除文章时是否同时删除文章中的上传文件
ttup=Request("ttup")          '共享组件,上传的类型
	id=strint(request("id"))
	oneid=strint(request("oneid"))
	twoid=strint(request("twoid"))
	threeid=strint(request("threeid"))
If blnAllow=True Then
strExt="只允许"
Else
strExt="禁止"
End If


If ttup="upnews" Then
SavePath = "UploadFiles/upnews"     '存放上传新闻主图片文件的目录
upname="pic_url"'回传表单变量名
elseIf ttup="photo" Then
SavePath = "UploadFiles/upnews"     '存放上传新闻相册文件的目录
upname=request("frmname")'回传表单变量名
elseIf ttup="link" Then
SavePath = "UploadFiles/link"       '存放友情链接上传文件的目录
upname="logo"
ElseIf ttup="sdlogo" Then
SavePath = "UploadFiles/sdlogo"     '存放店标图片上传文件的目录
upname="tupian"
ElseIf ttup="sdBanner" Then
SavePath = "UploadFiles/sdlogo"     '存放Banner上传文件的目录
upname="sdBanner"
ElseIf ttup="info" Then
SavePath = "UploadFiles/infopic"   '存放添加信息上传主图的目录
upname="tpname"
ElseIf ttup="thumbnail" Then
SavePath = "UploadFiles/magazine"   '存放电子杂志上传图片的目录
upname="thumbnail"
ElseIf ttup="largerpic" Then
SavePath = "UploadFiles/magazine"   '存放电子杂志上传图片的目录
upname="largerpic"
ElseIf ttup="certificate1" Then
SavePath = "UploadFiles/certificate"     '存放店铺名片审核证照上传文件的目录
upname="certificate1"
ElseIf ttup="certificate2" Then
SavePath = "UploadFiles/certificate"     '存放店铺名片审核证照上传文件的目录
upname="certificate2"
ElseIf ttup="certificate3" Then
SavePath = "UploadFiles/certificate"     '存放店铺名片审核证照上传文件的目录
upname="certificate3"
ElseIf ttup="certificate4" Then
SavePath = "UploadFiles/certificate"     '存放店铺名片审核证照上传文件的目录
upname="certificate4"
ElseIf ttup="certificate5" Then
SavePath = "UploadFiles/certificate"     '存放店铺名片审核证照上传文件的目录
upname="certificate5"
ElseIf ttup="certificate6" Then
SavePath = "UploadFiles/certificate"     '存放店铺名片审核证照上传文件的目录
upname="certificate6"
ElseIf ttup="certificate7" Then
SavePath = "UploadFiles/certificate"     '存放店铺名片审核证照上传文件的目录
upname="certificate7"
Else
Response.Write "<script language='javascript'>alert('不是默认的上传选项，不能上传!');history.go(-1);</script>"
response.end
End if
if right(SavePath,1)<>"/" then SavePath=SavePath&"/" '在目录后加(/)
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


	FSPath=GetFilePath(Server.mappath("DNJY_upfile.asp"),"\")'取得当前文件在服务器路径
	ServerPath=GetFilePath(Request.ServerVariables("HTTP_REFERER"),"/")'取得在网站上的位置
	for each formName in upfile.file '列出所有上传了的文件
	set oFile=upfile.file(formname)

fileExt=lcase(ofile.FileExt)
randomize'初始化随机数发生器
ranNum=int(900*rnd)+100
filename=SavePath&year(now) & right("0" & month(now),2) & right("0" & day(now),2) & right("0" & hour(now),2) & right("0" & minute(now),2) & right("0" & second(now),2)&ranNum&"."&fileExt'获得日期时间文件名
xxname=year(now) & right("0" & month(now),2) & right("0" & day(now),2) & right("0" & hour(now),2) & right("0" & minute(now),2) & right("0" & second(now),2)&ranNum&"."&fileExt'获得信息图片日期时间文件名 
upfile.SaveToFile formname,FSPath&FileName   '保存文件 也可以使用AutoSave来保存,参数一样,但是会自动建立新的文件名
		
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
Response.Write "<script language='javascript'>alert('上传失败，原因：文件有问题！请用制图软件转换格式或编辑保存再上传。');</script>"
msg="请选项正确文件"
end If
Else
Response.Write "<script language='javascript'>alert('不选文件上传什么啊!');history.go(-1);</script>"
response.end
end If

strJS="<SCRIPT language=javascript>" & vbcrlf	   
if muma<>1 then			
Response.Write "<script language='javascript'>alert('图片上传成功！');</script>"
Response.Write "<a href=javascript:window.close()>关闭本窗口</a>"



If ttup="info" Then fileName=xxName
if ttup="photo" then
strJS=strJS & "self.opener.addNEWS."&upname&".value ='"&filename&"';"
else
strJS=strJS & "parent.document.forms[0]."&upname&".value='"& fileName & "';" & vbcrlf    '回填上传图片地址,参数为upname
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
    set upfile=nothing  '删除此对象


function GetFilePath(FullPath,str)
  If FullPath <> "" Then
    GetFilePath = left(FullPath,InStrRev(FullPath, str))
    Else
    GetFilePath = ""
  End If
End function%>
