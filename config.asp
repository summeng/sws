<!--#include file="api/config.asp"--><%'=================================可修改网站安装路径====================================
strInstallDir="" '根目录留空，子目录在后面加“/”，如test/'======================================================================================

dim contentss,Readinfo,rsconfig,title,webname,web,strInstallDir,logo,Version,CacheName,vipje,huiyuansp,huiyuanxx,jiaoyi,overdue,addxinxi,xinxish,vipsj,vipno,onoff,kefu,kefuid,qqskin,qqshow,qqimg,kefuskin,stopsm,icp,leixing,lockip,addip,hydlip,htdlip,gpwip,bookip,infoip,sqlip,infosj,b_y,tui_y,killword,killmemo,killname,codekg,codekg1,codekg2,codekg3,codekg4,codekg5,lmkg,lmkg1,lmkg2,lmkg3,lmkg4,lmkg5,lmkg6,lmkg7,lmkg8,lmkg9,lmkg10,lmkg11,lmkg12,lmkg13,lmkg14,lmkg15,biaotinum,memonum,z_hb,z_a,z_b,z_c,z_d,jf_1,jf_2,jf_3,jf_4,jf_5,jf_hb,jf_ck,adjfs,g_a,g_b,g_c,g_d,rmb_hb,rmb_mc,xinxis,xnp,xnpsj,zcfbsj,hdlb,contribute,slidelx,addlink,linklogo,times,zcNum,hotbz,SY_Type,SY_interval,SY_Text,SY_FontName,SY_FontSize,SY_FontColor,SY_Bold,SY_FileName,SY_Width,SY_X,SY_Y,SY_Transparence,SY_BackgroundColor,SY_coordinates,regmail,citys,asphtm,content_length,HtmlHome,mailqr,usersh,zdsh,regyxq,Filterhtm,gxkf,cip,shows,shows1,shows2,shows3,shows4,shows5,shows6,shows7,shows8,shows9,shows10,shows11,shows12,usernews,userlink,adjf,newspl,tgjf,webeditor,sms_kg,sms_del,plshwebname="天天供求信息网站管理系统"web="test.ip126.com"logo="images/logo.gif"cip=FalseVersion="4.4"CacheName="ttgq"kefu=1kefuid="ttgq"qqshow=5qqimg=6kefuskin=4stopsm="欢迎光临本站，系统正在维护，约今晚9时开放，对您造成不便请见谅。"icp="ICP备09501888号"leixing="出售|求购|出租|求租|求职|招聘|合作|其它"z_hb=5z_a=2z_b=3z_c=2z_d=0jf_1=5jf_2=2jf_3=2jf_4=5jf_5=2jf_hb=20jf_ck=4adjfs=20g_a=2g_b=3g_c=2g_d=1rmb_mc="虚拟币"rmb_hb=1vipje=5huiyuansp=6huiyuanxx=5jiaoyi=1overdue=1addxinxi=0xinxish=0usernews=0userlink=0adjf=0vipsj=12vipno=1onoff=1lockip="1"addip="1"hydlip=5htdlip=5gpwip=5bookip=2infoip=5sqlip=3infosj=60b_y=2tui_y=100xinxis=1xnp=10xnpsj="2006-8-30 下午 09:56:30"killword="六合彩|六盒彩|六合采|发票|法轮功|枪|无抵押|贷款|免抵押|侦探|窃听器|文凭|群发"killmemo="六合彩|六盒彩|六合采|发票|法轮功|枪|窃听器|文凭|群发"killname="付本元|高立胜"codekg1=0codekg2=0codekg3=0codekg4=0codekg5=1lmkg1=1lmkg2=0lmkg3=1lmkg4=1lmkg5=1lmkg6=1lmkg7=1lmkg8=1lmkg9=1lmkg10=1lmkg11=1lmkg12=1lmkg13=1lmkg14=1lmkg15=1shows1=1shows2=1shows3=1shows4=1shows5=1shows6=1shows7=1shows8=0shows9=1shows10=1shows11=1shows12=1zcfbsj=60hdlb=2contribute=1slidelx=1biaotinum=40memonum=10000addlink=0linklogo=1times=24zcNum=1hotbz=100SY_Type=1SY_interval=20SY_Text="天天网络科技"SY_FontName="黑体"SY_FontSize=18SY_FontColor="#FFFFFF"SY_Bold=1SY_FileName="\images/watermark.gif"SY_Width=150SY_X=10SY_Y=10SY_Transparence=0.6SY_BackgroundColor="#0000FF"SY_coordinates=4'首页头条图片标题参数Const HomeEliteWidth = "380" '头条图片宽度Const HomeEliteHeight = "50" '头条图片高度Const HomeEliteFontSize = "40" '头条字体大小Const HomeEliteFontFamily = "黑体" '头条字体Const HomeElitePath = "UploadFiles/biaoti/" '头条图片存储位置，例如UploadFiles/Home/，不可以“/”或者“..”开头Const HomeEliteKey = "HereIsYourKey" '头条插件认证码，请尽量复杂regmail=1mailqr=0usersh=0zdsh=0regyxq=3citys=16asphtm=0content_length=200HtmlHome=0Filterhtm=0gxkf=2newspl=0tgjf=2webeditor=""sms_kg=Truesms_del=30plsh=1%>