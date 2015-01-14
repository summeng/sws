<!--#include file="../conn.asp"-->
<!--#include file="../config.asp"-->
<!--#include file="../SqlIn.Asp"-->
<!--#include file="../Include/Function.asp"-->
<!--#include file="../Include/label.asp"-->
<%
'=====================================
'天天供求信息网站管理系统
'天天网络科技工作室出品
'客服中心http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>
<%function HTMLEncode(fString)
if not isnull(fString) then
    fString = replace(fString, ">", "&gt;")
    fString = replace(fString, "<", "&lt;")
    fString = Replace(fString, CHR(32), "&nbsp;")
    fString = Replace(fString, CHR(9), "&nbsp;")
    fString = Replace(fString, CHR(34), "&quot;")
    fString = Replace(fString, CHR(39), "&#39;")
    fString = Replace(fString, CHR(13), "")
    fString = Replace(fString, CHR(10) & CHR(10), "</P><P> ")
    fString = Replace(fString, CHR(10), "<BR> ")
    HTMLEncode = fString
end if
end function
dim str,InstallDir,bt3,bt4,HomeElite,HomeElite1,HomeElite2,HomeElite3,HomeElite4,HomeElite5
Call OpenConn
set rsconfig = Server.CreateObject("ADODB.RecordSet")
sql="select * from DNJY_config order by id "&DN_OrderType&""
rsconfig.open sql,conn,3,3

title=rsconfig("title")
web=rsconfig("web")
InstallDir=rsconfig("InstallDir")
If InstallDir<>"" then
If Right(InstallDir, 1) <> "/" Then
strInstallDir = InstallDir & "/"
End If
else
strInstallDir=InstallDir
End If

'判断文件夹是否存在
Dim M_fso
Set M_fso = CreateObject("Scripting.FileSystemObject")
If Not (M_fso.FolderExists(Server.MapPath("/"&InstallDir&""))) Then
REsponse.write("你设置的系统安装目录“"&InstallDir&"”不存在，无法生成配置文件，系统将不能正常运行。请手工建立该目录，并将程序放在该目录下！")
response.end
End If
Set M_fso=Nothing

'判断是否在根目录
If InstallDir="" then
Dim filespec,fso
filespec=server.mappath("/"&strInstallDir&"DNJY_clsUp.asp")     
Set fso=CreateObject("Scripting.FileSystemObject")   
If Not (fso.FileExists(filespec)) Then        
response.write "你安装本系统的目录路径与后台网站参数管理中设置的系统安装目录不一致，系统将无法正常运行，请检查并正确放置本系统！"
response.end
End If   
End If
Set fso=Nothing
logo=rsconfig("logo")
Version=rsconfig("Version")
CacheName=rsconfig("CacheName")
kefu=rsconfig("kefu")
kefuid=rsconfig("kefuid")
qqskin=rsconfig("qqskin")
qqshow=rsconfig("qqshow")
qqimg=rsconfig("qqimg")
kefuskin=rsconfig("kefuskin")
stopsm=HTMLEncode(rsconfig("stopsm"))
icp=rsconfig("icp")
leixing=rsconfig("leixing")
z_hb=rsconfig("z_hb")
z_a=rsconfig("z_a")
z_b=rsconfig("z_b")
z_c=rsconfig("z_c")
z_d=rsconfig("z_d")
jf_1=rsconfig("jf_1")
jf_2=rsconfig("jf_2")
jf_3=rsconfig("jf_3")
jf_4=rsconfig("jf_4")
jf_5=rsconfig("jf_5")
jf_hb=rsconfig("jf_hb")
jf_ck=rsconfig("jf_ck")
adjfs=rsconfig("adjfs")
g_a=rsconfig("g_a")
g_b=rsconfig("g_b")
g_c=rsconfig("g_c")
g_d=rsconfig("g_d")
rmb_mc=rsconfig("rmb_mc")
rmb_hb=rsconfig("rmb_hb")
vipje=rsconfig("vipje")
huiyuansp=rsconfig("huiyuansp")
huiyuanxx=rsconfig("huiyuanxx")
jiaoyi=rsconfig("jiaoyi")
overdue=rsconfig("overdue")
addxinxi=rsconfig("addxinxi")
xinxish=rsconfig("xinxish")
usernews=rsconfig("usernews")
userlink=rsconfig("userlink")
adjf=rsconfig("adjf")
vipsj=rsconfig("vipsj")
vipno=rsconfig("vipno")
onoff=rsconfig("onoff")
lockip=rsconfig("lockip")
addip=rsconfig("addip")
hydlip=rsconfig("hydlip")
htdlip=rsconfig("htdlip")
gpwip=rsconfig("gpwip")
bookip=rsconfig("bookip")
infoip=rsconfig("infoip")
sqlip=rsconfig("sqlip")
infosj=rsconfig("infosj")
b_y=rsconfig("b_y")
tui_y=rsconfig("tui_y")
xinxis=rsconfig("xinxis")
xnp=rsconfig("xnp")
xnpsj=rsconfig("xnpsj")
killword=rsconfig("killword")
killmemo=rsconfig("killmemo")
killname=rsconfig("killname")
codekg=rsconfig("codekg")
codekg=split(codekg,"|")
codekg1=codekg(0)
codekg2=codekg(1)
codekg3=codekg(2)
codekg4=codekg(3)
codekg5=codekg(4)
lmkg=rsconfig("lmkg")
lmkg=split(lmkg,"|")
lmkg1=lmkg(0)
lmkg2=lmkg(1)
lmkg3=lmkg(2)
lmkg4=lmkg(3)
lmkg5=lmkg(4)
lmkg6=lmkg(5)
lmkg7=lmkg(6)
lmkg8=lmkg(7)
lmkg9=lmkg(8)
lmkg10=lmkg(9)
lmkg11=lmkg(10)
lmkg12=lmkg(11)
lmkg13=lmkg(12)
lmkg14=lmkg(13)
lmkg15=lmkg(14)
cip=rsconfig("cip")
shows=rsconfig("shows")
shows=split(shows,"|")
shows1=shows(0)
shows2=shows(1)
shows3=shows(2)
shows4=shows(3)
shows5=shows(4)
shows6=shows(5)
shows7=shows(6)
shows8=shows(7)
shows9=shows(8)
shows10=shows(9)
shows11=shows(10)
shows12=shows(11)
zcfbsj=rsconfig("zcfbsj")
hdlb=rsconfig("hdlb")
bt3=hdlb
contribute=rsconfig("contribute")
slidelx=rsconfig("slidelx")
biaotinum=rsconfig("biaotinum")
memonum=rsconfig("memonum")
bt4=memonum
addlink=rsconfig("addlink")
linklogo=rsconfig("linklogo")
times=rsconfig("times")
zcNum=rsconfig("zcNum")
times=rsconfig("times")
hotbz=rsconfig("hotbz")
SY_Type=rsconfig("SY_Type")
SY_interval=rsconfig("SY_interval")
SY_Text=rsconfig("SY_Text")
SY_FontName=rsconfig("SY_FontName")
SY_FontSize=rsconfig("SY_FontSize")
SY_FontColor=rsconfig("SY_FontColor")
SY_Bold=rsconfig("SY_Bold")
SY_FileName=rsconfig("SY_FileName")
SY_Width=rsconfig("SY_Width")
SY_X=rsconfig("SY_X")
SY_Y=rsconfig("SY_Y")
SY_Transparence=rsconfig("SY_Transparence")
SY_BackgroundColor=rsconfig("SY_BackgroundColor")
SY_coordinates=rsconfig("SY_coordinates")
HomeElite=rsconfig("HomeElite")
HomeElite=split(HomeElite,"|")
HomeElite1=HomeElite(0)
HomeElite2=HomeElite(1)
HomeElite3=HomeElite(2)
HomeElite4=HomeElite(3)
HomeElite5=HomeElite(4)
regmail=rsconfig("regmail")
mailqr=rsconfig("mailqr")
usersh=rsconfig("usersh")
zdsh=rsconfig("zdsh")
regyxq=rsconfig("regyxq")
citys=rsconfig("citys")
asphtm=rsconfig("asphtm")
content_length=rsconfig("content_length")
HtmlHome=rsconfig("HtmlHome")
Filterhtm=rsconfig("Filterhtm")
gxkf=rsconfig("gxkf")
newspl=rsconfig("newspl")
tgjf=rsconfig("tgjf")
webeditor=rsconfig("webeditor")
sms_kg=rsconfig("sms_kg")
sms_del=rsconfig("sms_del")
plsh=rsconfig("plsh")
rsconfig.close
set rsconfig=nothing
Call CloseDB(conn)
str=""
str=str &"<!--#include file=""api/config.asp""-->"&vbCR
str=str &"<%"&vbCR
str=str &"'=================================可修改网站安装路径====================================" & vbcrlf
str=str &"strInstallDir="""&strInstallDir&""" '根目录留空，子目录在后面加“/”，如test/"&vbCR
str=str &"'======================================================================================" & vbcrlf
str=str  & vbcrlf
str=str &"dim "
str=str &"contentss,Readinfo,rsconfig,title,webname,web,strInstallDir,logo,Version,CacheName,vipje,huiyuansp,huiyuanxx,jiaoyi,overdue,addxinxi,xinxish,vipsj,vipno,onoff,kefu,kefuid,qqskin,qqshow,qqimg,kefuskin,stopsm,icp,leixing,lockip,addip,hydlip,htdlip,gpwip,bookip,infoip,sqlip,infosj,b_y,tui_y,killword,killmemo,killname,codekg,codekg1,codekg2,codekg3,codekg4,codekg5,lmkg,lmkg1,lmkg2,lmkg3,lmkg4,lmkg5,lmkg6,lmkg7,lmkg8,lmkg9,lmkg10,lmkg11,lmkg12,lmkg13,lmkg14,lmkg15,biaotinum,memonum,z_hb,z_a,z_b,z_c,z_d,jf_1,jf_2,jf_3,jf_4,jf_5,jf_hb,jf_ck,adjfs,g_a,g_b,g_c,g_d,rmb_hb,rmb_mc,xinxis,xnp,xnpsj,zcfbsj,hdlb,contribute,slidelx,addlink,linklogo,times,zcNum,hotbz,SY_Type,SY_interval,SY_Text,SY_FontName,SY_FontSize,SY_FontColor,SY_Bold,SY_FileName,SY_Width,SY_X,SY_Y,SY_Transparence,SY_BackgroundColor,SY_coordinates,regmail,citys,asphtm,content_length,HtmlHome,mailqr,usersh,zdsh,regyxq,Filterhtm,gxkf,cip,shows,shows1,shows2,shows3,shows4,shows5,shows6,shows7,shows8,shows9,shows10,shows11,shows12,usernews,userlink,adjf,newspl,tgjf,webeditor,sms_kg,sms_del,plsh"&vbCR
str=str &"webname="""&title&""""&vbCR
str=str &"web="""&web&""""&vbCR
str=str &"logo="""&logo&""""&vbCR
str=str &"cip="&cip&""&vbCR
str=str &"Version="""&Version&""""&vbCR
str=str &"CacheName="""&CacheName&""""&vbCR
str=str &"kefu="&kefu&""&vbCR
str=str &"kefuid="""&kefuid&""""&vbCR
'str=str &"qqskin="&qqskin&""&vbCR
str=str &"qqshow="&qqshow&""&vbCR
str=str &"qqimg="&qqimg&""&vbCR
str=str &"kefuskin="&kefuskin&""&vbCR
str=str &"stopsm="""&stopsm&""""&vbCR
str=str &"icp="""&icp&""""&vbCR
str=str &"leixing="""&leixing&""""&vbCR
str=str &"z_hb="&z_hb&""&vbCR
str=str &"z_a="&z_a&""&vbCR
str=str &"z_b="&z_b&""&vbCR
str=str &"z_c="&z_c&""&vbCR
str=str &"z_d="&z_d&""&vbCR
str=str &"jf_1="&jf_1&""&vbCR
str=str &"jf_2="&jf_2&""&vbCR
str=str &"jf_3="&jf_3&""&vbCR
str=str &"jf_4="&jf_4&""&vbCR
str=str &"jf_5="&jf_5&""&vbCR
str=str &"jf_hb="&jf_hb&""&vbCR
str=str &"jf_ck="&jf_ck&""&vbCR
str=str &"adjfs="&adjfs&""&vbCR
str=str &"g_a="&g_a&""&vbCR
str=str &"g_b="&g_b&""&vbCR
str=str &"g_c="&g_c&""&vbCR
str=str &"g_d="&g_d&""&vbCR
str=str &"rmb_mc="""&rmb_mc&""""&vbCR
str=str &"rmb_hb="&rmb_hb&""&vbCR
str=str &"vipje="&vipje&""&vbCR
str=str &"huiyuansp="&huiyuansp&""&vbCR
str=str &"huiyuanxx="&huiyuanxx&""&vbCR
str=str &"jiaoyi="&jiaoyi&""&vbCR
str=str &"overdue="&overdue&""&vbCR
str=str &"addxinxi="&addxinxi&""&vbCR
str=str &"xinxish="&xinxish&""&vbCR
str=str &"usernews="&usernews&""&vbCR
str=str &"userlink="&userlink&""&vbCR
str=str &"adjf="&adjf&""&vbCR
str=str &"vipsj="&vipsj&""&vbCR
str=str &"vipno="&vipno&""&vbCR
str=str &"onoff="&onoff&""&vbCR
str=str &"lockip="""&lockip&""""&vbCR
str=str &"addip="""&addip&""""&vbCR
str=str &"hydlip="&hydlip&""&vbCR
str=str &"htdlip="&htdlip&""&vbCR
str=str &"gpwip="&gpwip&""&vbCR
str=str &"bookip="&bookip&""&vbCR
str=str &"infoip="&infoip&""&vbCR
str=str &"sqlip="&sqlip&""&vbCR
str=str &"infosj="&infosj&""&vbCR
str=str &"b_y="&b_y&""&vbCR
str=str &"tui_y="&tui_y&""&vbCR
str=str &"xinxis="&xinxis&""&vbCR
str=str &"xnp="&xnp&""&vbCR
str=str &"xnpsj="""&xnpsj&""""&vbCR
str=str &"killword="""&killword&""""&vbCR
str=str &"killmemo="""&killmemo&""""&vbCR
str=str &"killname="""&killname&""""&vbCR
str=str &"codekg1="&codekg1&""&vbCR
str=str &"codekg2="&codekg2&""&vbCR
str=str &"codekg3="&codekg3&""&vbCR
str=str &"codekg4="&codekg4&""&vbCR
str=str &"codekg5="&codekg5&""&vbCR
str=str &"lmkg1="&lmkg1&""&vbCR
str=str &"lmkg2="&lmkg2&""&vbCR
str=str &"lmkg3="&lmkg3&""&vbCR
str=str &"lmkg4="&lmkg4&""&vbCR
str=str &"lmkg5="&lmkg5&""&vbCR
str=str &"lmkg6="&lmkg6&""&vbCR
str=str &"lmkg7="&lmkg7&""&vbCR
str=str &"lmkg8="&lmkg8&""&vbCR
str=str &"lmkg9="&lmkg9&""&vbCR
str=str &"lmkg10="&lmkg10&""&vbCR
str=str &"lmkg11="&lmkg11&""&vbCR
str=str &"lmkg12="&lmkg12&""&vbCR
str=str &"lmkg13="&lmkg13&""&vbCR
str=str &"lmkg14="&lmkg14&""&vbCR
str=str &"lmkg15="&lmkg15&""&vbCR
str=str &"shows1="&shows1&""&vbCR
str=str &"shows2="&shows2&""&vbCR
str=str &"shows3="&shows3&""&vbCR
str=str &"shows4="&shows4&""&vbCR
str=str &"shows5="&shows5&""&vbCR
str=str &"shows6="&shows6&""&vbCR
str=str &"shows7="&shows7&""&vbCR
str=str &"shows8="&shows8&""&vbCR
str=str &"shows9="&shows9&""&vbCR
str=str &"shows10="&shows10&""&vbCR
str=str &"shows11="&shows11&""&vbCR
str=str &"shows12="&shows12&""&vbCR
str=str &"zcfbsj="&zcfbsj&""&vbCR
str=str &"hdlb="&bt3&""&vbCR
str=str &"contribute="&contribute&""&vbCR
str=str &"slidelx="&slidelx&""&vbCR
str=str &"biaotinum="&biaotinum&""&vbCR
str=str &"memonum="&bt4&""&vbCR
str=str &"addlink="&addlink&""&vbCR
str=str &"linklogo="&linklogo&""&vbCR
str=str &"times="&times&""&vbCR
str=str &"zcNum="&zcNum&""&vbCR
str=str &"hotbz="&hotbz&""&vbCR
str=str &"SY_Type="&SY_Type&""&vbCR
str=str &"SY_interval="&SY_interval&""&vbCR
str=str &"SY_Text="""&SY_Text&""""&vbCR
str=str &"SY_FontName="""&SY_FontName&""""&vbCR
str=str &"SY_FontSize="&SY_FontSize&""&vbCR
str=str &"SY_FontColor="""&SY_FontColor&""""&vbCR
str=str &"SY_Bold="&SY_Bold&""&vbCR
str=str &"SY_FileName="""&SY_FileName&""""&vbCR
str=str &"SY_Width="&SY_Width&""&vbCR
str=str &"SY_X="&SY_X&""&vbCR
str=str &"SY_Y="&SY_Y&""&vbCR
str=str &"SY_Transparence="&SY_Transparence&""&vbCR
str=str &"SY_BackgroundColor="""&SY_BackgroundColor&""""&vbCR
str=str &"SY_coordinates="&SY_coordinates&""&vbCR
str=str &"'首页头条图片标题参数"&vbCR
str=str &"Const HomeEliteWidth = """&HomeElite1&""" '头条图片宽度"&vbCR
str=str &"Const HomeEliteHeight = """&HomeElite2&""" '头条图片高度"&vbCR
str=str &"Const HomeEliteFontSize = """&HomeElite3&""" '头条字体大小"&vbCR
str=str &"Const HomeEliteFontFamily = """&HomeElite4&""" '头条字体"&vbCR
str=str &"Const HomeElitePath = """&HomeElite5&""" '头条图片存储位置，例如UploadFiles/Home/，不可以“/”或者“..”开头"&vbCR
str=str &"Const HomeEliteKey = ""HereIsYourKey"" '头条插件认证码，请尽量复杂"&vbCR
str=str &"regmail="&regmail&""&vbCR
str=str &"mailqr="&mailqr&""&vbCR
str=str &"usersh="&usersh&""&vbCR
str=str &"zdsh="&zdsh&""&vbCR
str=str &"regyxq="&regyxq&""&vbCR
str=str &"citys="&citys&""&vbCR
str=str &"asphtm="&asphtm&""&vbCR
str=str &"content_length="&content_length&""&vbCR
str=str &"HtmlHome="&HtmlHome&""&vbCR
str=str &"Filterhtm="&Filterhtm&""&vbCR
str=str &"gxkf="&gxkf&""&vbCR
str=str &"newspl="&newspl&""&vbCR
str=str &"tgjf="&tgjf&""&vbCR
str=str &"webeditor="""&webeditor&""""&vbCR
str=str &"sms_kg="&sms_kg&""&vbCR
str=str &"sms_del="&sms_del&""&vbCR
str=str &"plsh="&plsh&""&vbCR
str=str &"%"
str=str &">"

dim fs,f
Set fs = CreateObject("Scripting.FileSystemObject")
Set f = fs.CreateTextFile(Server.MapPath("/")&"/"&""&strInstallDir&"config.asp", True)
f.write(str)
f.close
Set f = nothing
Set fs = Nothing
If HtmlHome=1 Then
Call HomeHtml()'重新生成首页
Else
Dim objFSO
Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
if objFSO.fileExists(Server.MapPath("../index.html")) then
objFSO.DeleteFile(Server.MapPath("../index.html"))'删除静态首页
end if
set objfso=Nothing
End if
if request("page")<>"" then
Response.Write "<script language=javascript>setTimeout(""location.replace('"&request("page")&"')"",0)</script>"
Else
Response.Write "配置文件config.asp已生成!"
Response.Write "<br>"
Response.Write "<a href=javascript:window.close()>关闭本窗口</a>"
end if
%>
