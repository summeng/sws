<!--#include file="../conn.asp"-->
<!--#include file=cookies.asp-->
<!--#include file=../config.asp-->
<!--#include file="../Include/head.asp"-->
<!--#include file="../Include/Function.asp"-->
<!--#include file="../Include/label.asp"-->
<!--#include file="inc/Function.asp"-->
<!--#include file="inc/ubbcode.asp"-->
<%
'=====================================
'天天供求信息网站管理系统
'天天网络科技工作室出品
'客服中心http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>
<%
response.buffer=true
%>
<%Call OpenConn
Dim ItemID
Dim RsItem,SqlItem,FoundErr,ErrMsg
Dim UrlTest,TsString,ToString,CsString,CoString
Dim DateType,DsString,DoString,DateTime
Dim StopDateType,SDsString,SDoString,StopDateTime,StopDate
Dim AuthorType,AsString,AoString,AuthorStr,TelType,TelsString,TeloString,TelStr,sjType,sjsString,sjoString,sjStr,QQType,QsString,QoString,QQStr,dzType,dzsString,dzoString,dzStr,ybType,ybsString,yboString,ybStr,jgType,jgsString,jgoString,jgStr
Dim CopyFromType,FsString,FoString,CopyFromStr,leixingType,LxsString,LxoString,leixingStr,mailType,mailsString,mailoString,mailStr


Dim NewsUrl,NewsCode
Dim ConTent,Author,CopyFrom,jg,sj,dz,yb,telc,mail,qqc,leixingc
Dim strChannelDir
Dim cj_strInstallDir

cj_strInstallDir=trim(request.ServerVariables("SCRIPT_NAME"))
cj_strInstallDir=left(cj_strInstallDir,instrrev(lcase(cj_strInstallDir),"/")-1)
cj_strInstallDir=left(cj_strInstallDir,instrrev(lcase(cj_strInstallDir),"/"))
strChannelDir="Test"

ItemID=strint(Request.Form("ItemID"))           '项目ID
UrlTest=Trim(Request.Form("UrlTest"))           '一条信息
TsString=Request.Form("TsString")               '标题标签开始
ToString=Request.Form("ToString")               '标题标签结束
CsString=Request.Form("CsString")               '正文标签开始
CoString=Request.Form("CoString")               '正文标签结束

DateType=strint(Request.Form("DateType"))       '发布时间类型 0.当前时间 1.设置标签
DsString=Request.Form("DsString")               '发布时间标签开始
DoString=Request.Form("DoString")               '发布时间标签结束

StopDateType=strint(Request.Form("StopDateType")) '到期时间类型 0.当前时间+ 1.设置标签
SDsString=Request.Form("SDsString")             '到期时间标签开始
SDoString=Request.Form("SDoString")             '到期时间标签结束
StopDate=strint(request.form("StopDate"))       '到期时间+

jgType=strint(Request.Form("jgType"))   '交易价格
jgsString=Request.Form("jgsString")     
jgoString=Request.Form("jgoString")     
jgStr=Trim(Request.Form("jgStr"))  

AuthorType=strint(Request.Form("AuthorType"))   '作者设置类型 0.不作设置 1.设置标签 2.指定作者
AsString=Request.Form("AsString")               '作者标签开始
AoString=Request.Form("AoString")               '作者标签结束
AuthorStr=Trim(Request.Form("AuthorStr"))       '指定作者

TelType=strint(Request.Form("TelType"))   '电话设置类型 0.不作设置 1.设置标签 2.指定电话
TelsString=Request.Form("TelsString")               '电话标签开始
TeloString=Request.Form("TeloString")               '电话标签结束
TelStr=Trim(Request.Form("telStr"))       '指定电话

sjType=strint(Request.Form("sjType"))   '手机设置类型 0.不作设置 1.设置标签 2.指定号码
sjsString=Request.Form("sjsString")               '手机标签开始
sjoString=Request.Form("sjoString")               '手机标签结束
sjStr=Trim(Request.Form("sjStr"))       '指定手机

mailType=strint(Request.Form("mailType"))   '信箱设置类型 0.不作设置 1.设置标签 2.指定信箱
mailsString=Request.Form("mailsString")               '信箱标签开始
mailoString=Request.Form("mailoString")               '信箱标签结束
mailStr=Trim(Request.Form("mailStr"))       '指定信箱

QQType=strint(Request.Form("QQType"))   'QQ设置类型 0.不作设置 1.设置标签 2.指定QQ
QsString=Request.Form("QsString")               'QQ标签开始
QoString=Request.Form("QoString")               'QQ标签结束
QQStr=Trim(Request.Form("QQStr"))       '指定QQ

dzType=strint(Request.Form("dzType"))   '地址
dzsString=Request.Form("dzsString")     
dzoString=Request.Form("dzoString")     
dzStr=Trim(Request.Form("dzStr"))       

ybType=strint(Request.Form("ybType"))   '邮编
ybsString=Request.Form("ybsString")      
yboString=Request.Form("yboString")      
ybStr=Trim(Request.Form("ybStr"))   

CopyFromType=strint(Request.Form("CopyFromType"))'信息来源设置类型 0.不作设置 1.设置标签 2.指定来源 3.被采集页
FsString=Request.Form("FsString")               '信息来源标签开始
FoString=Request.Form("FoString")               '信息来源标签结束
CopyFromStr=Trim(Request.Form("CopyFromStr"))   '指定来源

leixingType=strint(Request.Form("leixingType"))'信息类型设置类型 0.不作设置 1.设置标签 2.指定类型
LxsString=Request.Form("LxsString")               '信息类型标签开始
LxoString=Request.Form("LxoString")               '信息类型标签结束
leixingStr=Trim(Request.Form("leixingStr"))   '指定类型

If ItemID=0 Then
   FoundErr=True
   ErrMsg=ErrMsg & "<br><li>参数错误，请从有效链接进入</li>"
End If

If UrlTest="" Then
   FoundErr=True
   ErrMsg=ErrMsg & "<br><li>参数错误，数据传递时发生错误</li>"
Else
   NewsUrl=UrlTest
End If

If TsString="" Then
   FoundErr=True
   ErrMsg=ErrMsg & "<br><li>标题开始标记不能为空</li>"
End If
If ToString="" Then
   FoundErr=True
   ErrMsg=ErrMsg & "<br><li>标题结束标记不能为空</li>" 
End If
If CsString="" Then
   FoundErr=True
   ErrMsg=ErrMsg & "<br><li>正文开始标记不能为空</li>"
End If
If CoString="" Then
   FoundErr=True
   ErrMsg=ErrMsg & "<br><li>正文结束标记不能为空</li>" 
End If


If DateType=1 Then
  If DsString="" or DoString="" Then
	 FoundErr=True
	 ErrMsg=ErrMsg & "<br><li>请将发布时间的开始/结束标记填写完整</li>" 
  End If
End If

If StopDateType=1 Then
  If SDsString="" or SDoString="" Then
	 FoundErr=True
	 ErrMsg=ErrMsg & "<br><li>请将到时期时间的开始/结束标记填写完整</li>" 
  End If
End If

If jgType=1 Then
  If jgsString="" or jgoString="" Then
	 FoundErr=True
	 ErrMsg=ErrMsg & "<br><li>请将交易价格的开始/结束标记填写完整</li>" 
  End If
ElseIf jgType=2 Then
  If jgStr="" Then
	 FoundErr=True
	 ErrMsg=ErrMsg & "<br><li>请指定交易价格</li>" 
  End If
End If
If IsNumeric(jgStr)=False And jgType=1 Then
         FoundErr=True
         ErrMsg=ErrMsg & "<br><li>交易价格只能是数字！</li>"
End If
      

If AuthorType=1 Then
  If AsString="" or AoString="" Then
	 FoundErr=True
	 ErrMsg=ErrMsg & "<br><li>请将作者的开始/结束标记填写完整</li>" 
  End If
ElseIf AuthorType=2 Then
  If AuthorStr="" Then
	 FoundErr=True
	 ErrMsg=ErrMsg & "<br><li>请指定作者</li>" 
  End If
End If

If TelType=1 Then
  If TelsString="" or TeloString="" Then
	 FoundErr=True
	 ErrMsg=ErrMsg & "<br><li>请将固定电话的开始/结束标记填写完整</li>" 
  End If
ElseIf TelType=2 Then
  If TelStr="" Then
	 FoundErr=True
	 ErrMsg=ErrMsg & "<br><li>请指定固定电话</li>" 
  End If
End If

If sjType=1 Then
  If sjsString="" or sjoString="" Then
	 FoundErr=True
	 ErrMsg=ErrMsg & "<br><li>请将手机号码的开始/结束标记填写完整</li>" 
  End If
ElseIf sjType=2 Then
  If sjStr="" Then
	 FoundErr=True
	 ErrMsg=ErrMsg & "<br><li>请指定手机号码</li>" 
  End If
End If

If mailType=1 Then
  If mailsString="" or mailoString="" Then
	 FoundErr=True
	 ErrMsg=ErrMsg & "<br><li>请将信箱的开始/结束标记填写完整</li>" 
  End If
ElseIf mailType=2 Then
  If mailStr="" Then
	 FoundErr=True
	 ErrMsg=ErrMsg & "<br><li>请指定信箱</li>" 
  End If
End If

If QQType=1 Then
  If QsString="" or QoString="" Then
	 FoundErr=True
	 ErrMsg=ErrMsg & "<br><li>请将QQ的开始/结束标记填写完整</li>" 
  End If
ElseIf QQType=2 Then
  If QQStr="" Then
	 FoundErr=True
	 ErrMsg=ErrMsg & "<br><li>请指定QQ</li>" 
  End If
End If

If dzType=1 Then
  If dzsString="" or dzoString="" Then
	 FoundErr=True
	 ErrMsg=ErrMsg & "<br><li>请将通讯地址的开始/结束标记填写完整</li>" 
  End If
ElseIf dzType=2 Then
  If dzStr="" Then
	 FoundErr=True
	 ErrMsg=ErrMsg & "<br><li>请指定通讯地址</li>" 
  End If
End If

If ybType=1 Then
  If ybsString="" or yboString="" Then
	 FoundErr=True
	 ErrMsg=ErrMsg & "<br><li>请将邮编的开始/结束标记填写完整</li>" 
  End If
ElseIf ybType=2 Then
  If ybStr="" Then
	 FoundErr=True
	 ErrMsg=ErrMsg & "<br><li>请指定邮编</li>" 
  End If
End If

If CopyFromType=1 Then
  If FsString="" or FoString="" Then
	 FoundErr=True
	 ErrMsg=ErrMsg & "<br><li>请将来源的开始/结束标记填写完整！</li>" 
  End If
ElseIf CopyFromType=2 Then
  If CopyFromStr="" Then
	 FoundErr=True
	 ErrMsg=ErrMsg & "<br><li>请指定来源</li>" 
  End If
End If

If leixingType=1 Then
  If LxsString="" or LxoString="" Then
	 FoundErr=True
	 ErrMsg=ErrMsg & "<br><li>请将类型的开始/结束标记填写完整！</li>" 
  End If
ElseIf leixingType=2 Then
  If leixingStr="" Then
	 FoundErr=True
	 ErrMsg=ErrMsg & "<br><li>请指定类型</li>" 
  End If
End If

'打开数据库



If FoundErr<>True Then
   SqlItem="Select * from DNJY_xx_Item Where ItemID=" & ItemID
   Set RsItem=server.CreateObject("adodb.recordset")
   RsItem.Open SqlItem,Conn,2,3

   RsItem("TsString")=TsString
   RsItem("ToString")=ToString
   RsItem("CsString")=CsString
   RsItem("CoString")=CoString

   RsItem("DateType")=DateType
   If DateType=1 Then
      RsItem("DsString")=DsString
      RsItem("DoString")=DoString
   End If

   RsItem("StopDateType")=StopDateType
   If StopDateType=0 Then
	  RsItem("StopDate")=StopDate
   ElseIf StopDateType=1 Then
	  RsItem("SDsString")=SDsString
	  RsItem("SDoString")=SDoString
   End If

   RsItem("jgType")=jgType
   If jgType=1 Then
      RsItem("jgsString")=jgsString
      RsItem("jgoString")=jgoString
   ElseIf jgType=2 Then
      RsItem("jgStr")=jgStr
   End If

   RsItem("AuthorType")=AuthorType
   If AuthorType=1 Then
      RsItem("AsString")=AsString
      RsItem("AoString")=AoString
   ElseIf AuthorType=2 Then
      RsItem("AuthorStr")=AuthorStr
   End If

   RsItem("TelType")=TelType
   If TelType=1 Then
      RsItem("TelsString")=TelsString
      RsItem("TeloString")=TeloString
   ElseIf TelType=2 Then
      RsItem("TelStr")=TelStr
   End If

   RsItem("sjType")=sjType
   If sjType=1 Then
      RsItem("sjsString")=sjsString
      RsItem("sjoString")=sjoString
   ElseIf sjType=2 Then
      RsItem("sjStr")=sjStr
   End If

   RsItem("mailType")=mailType
   If mailType=1 Then
      RsItem("mailsString")=mailsString
      RsItem("mailoString")=mailoString
   ElseIf mailType=2 Then
      RsItem("mailStr")=mailStr
   End If

   RsItem("QQType")=QQType
   If QQType=1 Then
      RsItem("QsString")=QsString
      RsItem("QoString")=QoString
   ElseIf QQType=2 Then
      RsItem("QQStr")=QQStr
   End If

   RsItem("dzType")=dzType
   If dzType=1 Then
      RsItem("dzsString")=dzsString
      RsItem("dzoString")=dzoString
   ElseIf dzType=2 Then
      RsItem("dzStr")=dzStr
   End If

   RsItem("ybType")=ybType
   If ybType=1 Then
      RsItem("ybsString")=ybsString
      RsItem("yboString")=yboString
   ElseIf ybType=2 Then
      RsItem("ybStr")=ybStr
   End If

   RsItem("CopyFromType")=CopyFromType
   If CopyFromType=1 Then
      RsItem("FsString")=FsString
      RsItem("FoString")=FoString
   ElseIf CopyFromType=2 Then
      RsItem("CopyFromStr")=CopyFromStr
   End If

   RsItem("leixingType")=leixingType
   If leixingType=1 Then
      RsItem("LxsString")=LxsString
      RsItem("LxoString")=LxoString
   ElseIf leixingType=2 Then
      RsItem("leixingStr")=leixingStr
   End If

   RsItem.UpDate
   RsItem.Close
   Set RsItem=Nothing
End If
'关闭数据库链接
conn.close:Set conn=nothing

If FoundErr<>True Then
   NewsCode=GetHttpPage(NewsUrl)
   If NewsCode<>"$False$" Then
      Title=GetBody(NewsCode,TsString,ToString,False,False)
      Content=GetBody(NewsCode,CsString,CoString,False,False)
      If Title="$False$" Then
         FoundErr=True
         ErrMsg=ErrMsg & "<br><li>在截取标题的时候发生错误：" & NewsUrl & "</li>"
      Else
         Title=FpHtmlEnCode(Title)
      End If

      If Content="$False$" Then
         FoundErr=True
         ErrMsg=ErrMsg & "<br><li>在截取正文的时候发生错误：" & NewsUrl & "</li>"
      End If
   Else
      FoundErr=True
      ErrMsg=ErrMsg & "<br><li>在获取源码时发生错误："& NewsUrl &"</li>"
   End If 
End If

If FoundErr<>True Then
      If DateType=0 then
         DateTime=date()
      Else
         DateTime=GetBody(NewsCode,DsString,DoString,False,False)
         DateTime=FpHtmlEncode(DateTime)
         DateTime=Trim(Replace(DateTime,"&nbsp;"," "))
         If IsDate(DateTime)=True Then
            DateTime=CDate(DateTime)
         Else
            DateTime=date()
         End If
      End If
	  DateTime=FormatDateTime(DateTime,2)

      If StopDateType=0 then
         StopDateTime=dateadd("d",StopDate,date())
      Else
         StopDateTime=GetBody(NewsCode,SDsString,SDoString,False,False)
         StopDateTime=FpHtmlEncode(StopDateTime)
         StopDateTime=Trim(Replace(StopDateTime,"&nbsp;"," "))
         If IsDate(StopDateTime)=True Then
            StopDateTime=CDate(StopDateTime)
         Else
            StopDateTime=dateadd("d",31,date())
         End If
      End If
	  StopDateTime=FormatDateTime(StopDateTime,2)

      If jgType=1 Then
         jg=GetBody(NewsCode,jgsString,jgoString,False,False)
      ElseIf jgType=2 Then
         jg=jgStr
      End If
      If jg="$False$" Or IsNull(jg) Then
         jg=0
      Else
      If IsInteger(jg)=True then
         jg=CLng(jg)
      Else
      jg=0
      End If
      End If

      If AuthorType=1 Then
         Author=GetBody(NewsCode,AsString,AoString,False,False)
      ElseIf AuthorType=2 Then
         Author=AuthorStr
      End If
      If Author="$False$" Or IsNull(Author) Then
         Author="佚名"
      Else
         Author=FpHtmlEnCode(Author)
      End If

      If TelType=1 Then
         Telc=GetBody(NewsCode,TelsString,TeloString,False,False)
      ElseIf TelType=2 Then
         Telc=TelStr
      End If
      If Telc="$False$" Or IsNull(Telc) Then
         Telc="未填"
      Else
         Telc=FpHtmlEnCode(Telc)
      End If

      If sjType=1 Then
         sj=GetBody(NewsCode,sjsString,sjoString,False,False)
      ElseIf sjType=2 Then
         sj=sjStr
      End If
      If sj="$False$" Or IsNull(sj) Then
         sj="未填"
      Else
         sj=FpHtmlEnCode(sj)
      End If

      If mailType=1 Then
         mail=GetBody(NewsCode,mailsString,mailoString,False,False)
      ElseIf mailType=2 Then
         mail=mailStr
      End If
      If mail="$False$" Or IsNull(mail) Then
         mail="未填"
      Else
         mail=FpHtmlEnCode(mail)
      End If

      If QQType=1 Then
         QQc=GetBody(NewsCode,qsString,qoString,False,False)
      ElseIf qqType=2 Then
         qqc=qqStr
      End If
      If qqc="$False$" Or IsNull(qqc) Then
         qqc="未填"
      Else
         qqc=FpHtmlEnCode(qqc)
      End If

      If dzType=1 Then
         dz=GetBody(NewsCode,dzsString,dzoString,False,False)
      ElseIf dzType=2 Then
         dz=dzStr
      End If
      If dz="$False$" Or IsNull(dz) Then
         dz="未填"
      Else
         dz=FpHtmlEnCode(dz)
      End If

      If ybType=1 Then
         yb=GetBody(NewsCode,ybsString,yboString,False,False)
      ElseIf ybType=2 Then
         yb=ybStr
      End If
      If yb="$False$" Or IsNull(yb) Then
         yb="未填"
      Else
         yb=FpHtmlEnCode(yb)
      End If
	  
      If CopyFromType=1 Then
         CopyFrom=GetBody(NewsCode,FsString,FoString,False,False)
      ElseIf CopyFromType=2 Then
         CopyFrom=CopyFromStr
      ElseIf CopyFromType=3 Then
         CopyFrom=NewsUrl
      End If
      If CopyFrom="$False$" Or IsNull(CopyFrom) Then
         CopyFrom="不详"
      Else
         CopyFrom=FpHtmlEnCode(CopyFrom)
      End If

      If leixingType=1 Then
         leixingc=GetBody(NewsCode,LxsString,LxoString,False,False)
      ElseIf leixingType=2 Then
         leixingc=leixingStr
      End If
      If leixingc="$False$" Or IsNull(leixingc) Then
         leixingc="未填"
      Else
         leixingc=FpHtmlEnCode(leixingc)
      End If

End If

If FoundErr<>True Then
   Content=ReplaceSaveRemoteFile(Content,cj_strInstallDir,strChannelDir,False,NewsUrl)
End If

If FoundErr=True Then
   Call WriteErrMsg(ErrMsg)
Else
   Call Main()
End if

%>
<%Sub Main()%>
<HTML>
<HEAD>
<TITLE>数据采集系统</TITLE>
<META http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="../css/manage.css" type="text/css">
</HEAD>
<BODY leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<TABLE width="100%" border="0" align="center" cellpadding="0" cellspacing="1" class="border">
  <TR> 
    <TD height="22" colspan="2" align="center" class="topbg"><strong>采&nbsp;&nbsp;集&nbsp;&nbsp;系&nbsp;&nbsp;统&nbsp;&nbsp;项&nbsp;&nbsp;目&nbsp;&nbsp;管&nbsp;&nbsp;理</strong></TD>
  </TR>
  <TR class="tdbg"> 
    <TD width="65" height="30"><STRONG>管理导航：</STRONG></TD>
    <TD height="30"><A href="collect_add_a.asp">添加项目</A> >> <A href="collect_editor_a.asp?ItemID=<%=ItemID%>">基本设置</A> >> <A href="collect_editor_b.asp?ItemID=<%=ItemID%>">列表设置</A> >> <A href="collect_editor_c.asp?ItemID=<%=ItemID%>">链接设置</A> >> <A href="collect_editor_d.asp?ItemID=<%=ItemID%>">正文设置</A> >> <FONT color=red>采样测试</FONT> >> 属性设置 >> 完成</TD>         
  </TR>         
</TABLE> 
<BR>        
<TABLE width="100%" border="0" align="center" cellpadding="0" cellspacing="1" class="border" >
    <TR> 
      <TD height="22" colspan="2" class="title"> <DIV align="center"><STRONG>添 加 新 项 目--采 样 测 试</STRONG></DIV></TD>
    </TR>
</TABLE>

<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" class="border" >
  <TR align="center" class="tdbg"> 
    <TD height="40" colspan="2" valign="bottom"><SPAN lang="zh-cn"><font size="5"><%=Title%></font></SPAN></TD>
  </TR>
  <TR align="center" class="tdbg">
    <TD colspan="2">
        来源：<%=CopyFrom%>&nbsp;&nbsp;作者：<%=Author%>&nbsp;&nbsp;发布时间：<%=DateTime%>&nbsp;&nbsp;到期时间：<%=StopDateTime%>
    </TD>
  </TR>
  <TR class="tdbg">
    <TD colspan="2">
      <TABLE width="100%" height="100%" border="0" cellpadding="0" cellspacing="5">
        <TR>
          <TD height="200" valign="top"><P><SPAN lang="zh-cn"><%=Content%></SPAN></P>
          </TD>
        </TR>
      </TABLE>
    </TD>
  </TR>
</table>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" class="border" >
  <tr class="tdbg">
    <td colspan="2">
      <table width="100%" height="100%" border="0" cellpadding="0" cellspacing="5">
        <tr>
          <td height="200" valign="top"><b>采集检测的其它相关项目：</b><p>
		  <p>信息类型：<%=leixingc%></p><p>交易价格：<%=jg%></p><p>固话：<%=Telc%></p><p>手机：<%=sj%></p><p>QQ：<%=QQc%></p><p>邮箱：<%=mail%></p><p>通讯地址：<%=dz%></p><p>邮编：<%=yb%></p>
</td>
        </tr>
      </table>
    </td>
  </tr>
</table>
<FORM method="post" action="collect_attribute.asp" name="form1">
<TABLE width="100%" border="0" align="center" cellpadding="0" cellspacing="1" class="border" >
    <TR class="tdbg"> 
      <TD colspan="2" align="center" class="tdbg">
        <INPUT name="ItemID" type="hidden" value="<%=ItemID%>">
        <INPUT name="button1" type="button" id="Cancel" value=" 上&nbsp;一&nbsp;步 " onClick="window.location.href='javascript:history.go(-1)'" style="cursor: hand;background-color: #cccccc;">
        &nbsp; 
        <INPUT  type="submit" name="Submit" value="  下&nbsp;一&nbsp;步 " style="cursor: hand;background-color: #cccccc;"></TD>
    </TR>
</TABLE>
</FORM>       
</BODY>         
</HTML>
<%End Sub%>


