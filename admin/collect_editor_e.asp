<!--#include file="../conn.asp"-->
<!--#include file="../config.asp"-->
<!--#include file=cookies.asp-->
<!--#include file="inc/Function.asp"-->
<!--#include file="inc/ubbcode.asp"-->
<!--#include file="../Include/Function.asp"-->
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

<%
Dim ItemID,Action
Dim RsItem,SqlItem,SqlF,RsF,FoundErr,ErrMsg
Dim LoginType,LoginUrl,LoginPostUrl,LoginUser,LoginPass,LoginFalse,LoginResult,LoginData
Dim ListStr,LsString,LoString,ListPaingType,LPsString,LPoString,ListPaingStr1,ListPaingStr2,ListPaingID1,ListPaingID2,ListPaingStr3,HsString,HoString,HttpUrlType,HttpUrlStr
Dim TsString,ToString,CsString,CoString,DateType,DsString,DoString,DateTime,AuthorType,AsString,AoString,AuthorStr,CopyFromType,FsString,FoString,CopyFromStr,NPsString,NPoString,NewsPaingStr,NewsPaingHtml,TelType,TelsString,TeloString,TelStr,Tel,sjType,sjsString,sjoString,sjStr,sj,mailType,mailsString,mailoString,mailStr,QQType,QsString,QoString,QQStr,QQ,dzType,dzsString,dzoString,dzStr,dz,ybType,ybsString,yboString,ybStr,yb,jgType,jgsString,jgoString,jgStr,jg,StopDateType,SDsString,SDoString,StopDateTime,StopDate,leixingType,LxsString,LxoString,leixingStr
Dim NewsPaingNext,NewsPaingNextCode,ContentTemp
Dim UrlTest,ListUrl,ListCode
Dim NewsUrl,NewsCode,NewsArrayCode,NewsArray
Dim Content,Author,CopyFrom,mail
Dim Arr_Filters,Filteri,FilterStr
Dim UpDateType

Dim cj_strInstallDir,UploadFiles,strChannelDir
cj_strInstallDir=trim(request.ServerVariables("SCRIPT_NAME"))
cj_strInstallDir=left(cj_strInstallDir,instrrev(lcase(cj_strInstallDir),"/")-1)
cj_strInstallDir=left(cj_strInstallDir,instrrev(lcase(cj_strInstallDir),"/"))
strChannelDir="Test"
FoundErr=False

ItemID=strint(Request("ItemID"))
Action=Trim(Request("Action"))

If ItemID=0 Then
   FoundErr=True
   ErrMsg=ErrMsg & "<br><li>参数错误，项目ID不能为空</li>"
End If

Call OpenConn
If Action="SaveEdit" And FoundErr<>True Then
   Call SaveEdit()
End If

If FoundErr<>True Then
   Call GetTest()
End If
'关闭数据库链接
conn.close:Set conn=nothing
If FoundErr<>True Then
   Call Main()
Else
   Call WriteErrMsg(ErrMsg)
End If
%>
<%Sub Main()%>
<html>
<head>
<title>采集系统</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="../css/manage.css" type="text/css">
</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" class="border">
<TR>
    <TD height="20" bgcolor="#799AE1" align="center" colspan="13"><FONT color="#FFFFFF" style="font-size:14px">采 集 系 统 项 目 管 理</FONT></TD>
 </TR>
  <tr class="tdbg"> 
    <td width="65" height="30"><strong>管理导航：</strong></td>
    <td height="30">项目编辑 >> <a href="collect_editor_a.asp?ItemID=<%=ItemID%>">基本设置</a> >> <a href="collect_editor_b.asp?ItemID=<%=ItemID%>">列表设置</a> >> <a href="collect_editor_c.asp?ItemID=<%=ItemID%>">链接设置</a> >> <a href="collect_editor_d.asp?ItemID=<%=ItemID%>">正文设置</a> >>  
	<a href="collect_editor_e.asp?ItemID=<%=ItemID%>"><font color=red>采样测试</font></a> >> <a href="collect_attribute.asp?ItemID=<%=ItemID%>">属性设置</a> >> 完成</td>         
  </tr>         
</table> 
<br>        
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" class="border" >
    <tr> 
      <td height="22" colspan="2" class="title"> <div align="center"><strong>编 辑 项 目--采 样 测 试</strong></div></td>
    </tr>
</table>

<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" class="border" >
  <tr align="center" class="tdbg"> 
    <td colspan="2" valign="bottom"><span lang="zh-cn">
	<font size="5"><%=Title%></font></span></td>
  </tr>
  <tr align="center" class="tdbg">
    <td colspan="2">
        来源：
	<%IF CopyFromType=3 Then
			Response.write NewsUrl
	  Else
		  Response.write  CopyFrom 
	  End IF%>
		   &nbsp;&nbsp;作者：<%=Author%>&nbsp;&nbsp;更新时间：<%=DateTime%>&nbsp;到期时间:<%=StopDateTime%>
    </td>
  </tr>
  <tr class="tdbg">
    <td colspan="2">
      <table width="100%" height="100%" border="0" cellpadding="0" cellspacing="5">
        <tr>
          <td height="200" valign="top"><p><span lang="zh-cn"><%=Content%></span></p>		  
</td>
        </tr>
      </table>
    </td>
  </tr>
</table>

<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" class="border" >
  <tr class="tdbg">
    <td colspan="2">
      <table width="100%" height="100%" border="0" cellpadding="0" cellspacing="5">
        <tr>
          <td height="200" valign="top"><b>采集检测的其它相关项目：</b><p>
		  <p>信息类型：<%=leixing%></p><p>交易价格：<%=jg%></p><p>固话：<%=Tel%></p><p>手机：<%=sj%></p><p>QQ：<%=QQ%></p><p>邮箱：<%=mail%></p><p>通讯地址：<%=dz%></p><p>邮编：<%=yb%></p>
</td>
        </tr>
      </table>
    </td>
  </tr>
</table>

<form method="post" action="collect_attribute.asp" name="form1">
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" class="border" >
    <tr class="tdbg"> 
      <td colspan="2" align="center" class="tdbg">
        <input name="Action" type="hidden" id="Action" value="SaveEdit">
        <input name="ItemID" type="hidden" id="ItemID" value="<%=ItemID%>">
        <input name="Cancel" type="button" id="Cancel" value=" 上&nbsp;一&nbsp;步 " onClick="window.location.href='javascript:history.go(-1)'" style="cursor: hand;background-color: #cccccc;">
        &nbsp; 
        <input  type="submit" name="Submit" value="  下&nbsp;一&nbsp;步 " style="cursor: hand;background-color: #cccccc;"></td>
    </tr>
</table>
</form>    
</body>         
</html>
<%End Sub%>

<%
'==================================================
'过程名：SaveEdit
'作  用：保存设置
'参  数：无
'==================================================
Sub SaveEdit

TsString=Request.Form("TsString")
ToString=Request.Form("ToString")
CsString=Request.Form("CsString")
CoString=Request.Form("CoString")

CopyFromType=strint(Request.Form("CopyFromType"))'信息来源设置类型 0.不作设置 1.设置标签 2.指定来源 3.被采集页
FsString=Request.Form("FsString")               '信息来源标签开始
FoString=Request.Form("FoString")               '信息来源标签结束
CopyFromStr=Trim(Request.Form("CopyFromStr"))   '指定来源

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
TelStr=Trim(Request.Form("TelStr"))       '指定电话

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
QQStr=Trim(Request.Form("QQStr"))       'QQ

dzType=strint(Request.Form("dzType"))   '地址
dzsString=Request.Form("dzsString")     
dzoString=Request.Form("dzoString")     
dzStr=Trim(Request.Form("dzStr"))       

ybType=strint(Request.Form("ybType"))   '邮编
ybsString=Request.Form("ybsString")      
yboString=Request.Form("yboString")      
ybStr=Trim(Request.Form("ybStr"))       


leixingType=strint(Request.Form("leixingType"))   '信息类型设置 0.不作设置 1.设置标签 2.指定类型
LxsString=Request.Form("LxsString")               '类型标签开始
LxoString=Request.Form("LxoString")               '类型标签结束
leixingStr=Trim(Request.Form("leixingStr"))       '类型

if instr(dostring,"&nbsp;") then
response.write "ok"
end if
UrlTest=Trim(Request.Form("UrlTest"))

If ItemID=0 Then
   FoundErr=True
   ErrMsg=ErrMsg & "<br><li>参数错误，项目ID不能为空</li>"
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
  If dsString="" or doString="" Then
	 FoundErr=True
	 ErrMsg=ErrMsg & "<br><li>请将发布时间的开始/结束标记填写完整</li>" 
  End If
End If

If StopDateType=1 Then
  If sdsString="" or sdoString="" Then
	 FoundErr=True
	 ErrMsg=ErrMsg & "<br><li>请将到期时间的开始/结束标记填写完整</li>" 
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

If telType=1 Then
  If telsString="" or teloString="" Then
	 FoundErr=True
	 ErrMsg=ErrMsg & "<br><li>请将固定电话的开始/结束标记填写完整</li>" 
  End If
ElseIf telType=2 Then
  If telStr="" Then
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
	 ErrMsg=ErrMsg & "<br><li>请将邮箱的开始/结束标记填写完整</li>" 
  End If
ElseIf mailType=2 Then
  If mailStr="" Then
	 FoundErr=True
	 ErrMsg=ErrMsg & "<br><li>请指定邮箱</li>" 
  End If
End If

If qqType=1 Then
  If qsString="" or qoString="" Then
	 FoundErr=True
	 ErrMsg=ErrMsg & "<br><li>请将QQ号的开始/结束标记填写完整</li>" 
  End If
ElseIf qqType=2 Then
  If qqStr="" Then
	 FoundErr=True
	 ErrMsg=ErrMsg & "<br><li>请指定QQ号</li>" 
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
	 ErrMsg=ErrMsg & "<br><li>请将信息类型的开始/结束标记填写完整</li>" 
  End If
ElseIf leixingType=2 Then
  If leixingStr="" Then
	 FoundErr=True
	 ErrMsg=ErrMsg & "<br><li>请指定信息类型</li>" 
  End If
End If

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
   RsItem.Close:Set RsItem=Nothing
End If
End Sub

'==================================================
'过程名：GetTest
'作  用：采集测试
'参  数：无
'==================================================
Sub GetTest()
   SqlItem="Select * from DNJY_xx_Item Where ItemID=" & ItemID
   Set RsItem=server.CreateObject("adodb.recordset")
   RsItem.Open SqlItem,Conn,1,1
   If  RsItem.Eof  And  RsItem.Bof  Then
         FoundErr=True
      ErrMsg=ErrMsg  &  "<br><li>参数错误，找不到该项目</li>"
   Else
      LoginType=RsItem("LoginType")
      LoginUrl=RsItem("LoginUrl")
      LoginPostUrl=RsItem("LoginPostUrl")
      LoginUser=RsItem("LoginUser")
      LoginPass=RsItem("LoginPass")
      LoginFalse=RsItem("LoginFalse")
      
      ListStr=RsItem("ListStr")
      LsString=RsItem("LsString")
      LoString=RsItem("LoString")
      ListPaingType=RsItem("ListPaingType")
      LPsString=RsItem("LPsString")
      LPoString=RsItem("LPoString")
      ListPaingStr1=RsItem("ListPaingStr1")
      ListPaingStr2=RsItem("ListPaingStr2")
      ListPaingID1=RsItem("ListPaingID1")
      ListPaingID2=RsItem("ListPaingID2")
      ListPaingStr3=RsItem("ListPaingStr3")
      
      HsString=RsItem("HsString")
      HoString=RsItem("HoString")
      HttpUrlType=RsItem("HttpUrlType")
      HttpUrlStr=RsItem("HttpUrlStr")
      
      TsString=RsItem("TsString")
      ToString=RsItem("ToString")
      CsString=RsItem("CsString")
      CoString=RsItem("CoString")
      
      DateType=RsItem("DateType")
      DsString=RsItem("DsString")
      DoString=RsItem("DoString")

      StopDateType=RsItem("StopDateType")
      sdsString=RsItem("sdsString")
      sdoString=RsItem("sdoString")

      jgType=RsItem("jgType")
      jgsString=RsItem("jgsString")
      jgoString=RsItem("jgoString")
      jgStr=RsItem("jgStr")

      AuthorType=RsItem("AuthorType")
      AsString=RsItem("AsString")
      AoString=RsItem("AoString")
      AuthorStr=RsItem("AuthorStr")

      telType=RsItem("telType")
      telsString=RsItem("telsString")
      teloString=RsItem("teloString")
      telStr=RsItem("telStr")

      sjType=RsItem("sjType")
      sjsString=RsItem("sjsString")
      sjoString=RsItem("sjoString")
      sjStr=RsItem("sjStr")

      mailType=RsItem("mailType")
      mailsString=RsItem("mailsString")
      mailoString=RsItem("mailoString")
      mailStr=RsItem("mailStr")

      qqType=RsItem("qqType")
      qsString=RsItem("qsString")
      qoString=RsItem("qoString")
      qqStr=RsItem("qqStr")

      dzType=RsItem("dzType")
      dzsString=RsItem("dzsString")
      dzoString=RsItem("dzoString")
      dzStr=RsItem("dzStr")

      ybType=RsItem("ybType")
      ybsString=RsItem("ybsString")
      yboString=RsItem("yboString")
      ybStr=RsItem("ybStr")

      CopyFromType=RsItem("CopyFromType")
      FsString=RsItem("FsString")
      FoString=RsItem("FoString")
      CopyFromStr=RsItem("CopyFromStr")

      leixingType=RsItem("leixingType")
      lxsString=RsItem("lxsString")
      lxoString=RsItem("lxoString")
      leixingStr=RsItem("leixingStr")

   End  If   
   RsItem.Close
   Set RsItem=Nothing

   If LoginType=1 Then
      If LoginUrl="" or LoginPostUrl="" or LoginUser="" Or LoginPass="" Or LoginFalse="" Then
         FoundErr=True
         ErrMsg=ErrMsg & "<br><li>您要采集的网站需要登录！请将登录信息填写完整</li>"
      End If
   End If
   If LsString="" Then
      FoundErr=True
      ErrMsg=ErrMsg & "<br><li>列表开始标记不能为空！</li>"
   End If
   If LoString="" Then
      FoundErr=True
      ErrMsg=ErrMsg & "<br><li>列表结束标记不能为空！</li>"
   End If
   If ListPaingType=0 Or ListPaingType=1 Then
      If ListStr="" Then
         FoundErr=True
         ErrMsg=ErrMsg & "<br><li>列表索引页不能为空！</li>"
      End If
      If ListPaingType=1 Then    
         If LPsString="" Or LPoString="" Then
            FoundErr=True
            ErrMsg=ErrMsg & "<br><li>索引分页开始、结束标记不能为空！</li>"
         End If
      End If      
      If  ListPaingStr1<>""  And  Len(ListPaingStr1)<15  Then
         FoundErr=True
         ErrMsg=ErrMsg & "<br><li>索引分页重定向设置不正确！</li>"
            End  IF
   ElseIf ListPaingType=2 Then
      If ListPaingStr2="" Then
         FoundErr=True
         ErrMsg=ErrMsg & "<br><li>批量生成原字符串不能为空！</li>"
      End If
      If IsNumeric(ListPaingID1)=False or IsNumeric(ListPaingID2)=False Then
         FoundErr=True
         ErrMsg=ErrMsg & "<br><li>批量生成的范围只能是数字！</li>"
      Else
         ListPaingID1=Clng(ListPaingID1)
         ListPaingID2=Clng(ListPaingID2)
         If ListPaingID1=0 And ListPaingID2=0 Then
            FoundErr=True
            ErrMsg=ErrMsg & "<br><li>批量生成的范围不正确！</li>"
         End If
      End If 
   ElseIf ListPaingType=3 Then
      If ListPaingStr3="" Then
         FoundErr=True
         ErrMsg=ErrMsg & "<br><li>索引分页不能为空！</li>"
      End If
   Else
      FoundErr=True
      ErrMsg=ErrMsg & "<br><li>请选择索引分页类型</li>"
   End If
     If  HsString=""  or  HoString=""  Then
      FoundErr=True
      ErrMsg=ErrMsg & "<br><li>链接开始/结束标记不能为空！</li>"
      End  If


   If FoundErr<>True And  Action<>"SaveEdit"  Then
      Select Case ListPaingType
      Case 0,1
            ListUrl=ListStr
      Case 2
         ListUrl=Replace(ListPaingStr2,"{$ID}",CStr(ListPaingID1))
      Case 3
         If Instr(ListPaingStr3,"|")> 0 Then
            ListUrl=Left(ListPaingStr3,Instr(ListPaingStr3,"|")-1)
         Else
            ListUrl=ListPaingStr3
         End If
      End Select
   End If
   
      If  FoundErr<>True  And  Action<>"SaveEdit"  And  LoginType=1  Then
      LoginData=UrlEncoding(LoginUser & "&" & LoginPass)
      LoginResult=PostHttpPage(LoginUrl,LoginPostUrl,LoginData)
      If Instr(LoginResult,LoginFalse)>0 Then
         FoundErr=True
         ErrMsg=ErrMsg & "<br><li>登录网站时发生错误，请确认登录信息的正确性！</li>"
      End If
      End  If
   
   If  FoundErr<>True  And  Action<>"SaveEdit"  Then
         ListCode=GetHttpPage(ListUrl)
         If ListCode<>"$False$" Then
            ListCode=GetBody(ListCode,LsString,LoString,False,False)
            If ListCode<>"$False$" Then
               NewsArrayCode=GetArray(ListCode,HsString,HoString,False,False)
               If NewsArrayCode<>"$False$" Then
                  If Instr(NewsArrayCode,"$Array$")>0 Then
                     NewsArray=Split(NewsArrayCode,"$Array$")
                     If HttpUrlType=1 Then
                        NewsUrl=Trim(Replace(HttpUrlStr,"{$ID}",NewsArray(0)))
                     Else
                        NewsUrl=Trim(DefiniteUrl(NewsArray(0),ListUrl))
                     End If
                  Else
                     FoundErr=True
                     ErrMsg=ErrMsg & "<br><li>只发现一个有效链接？：" & NewsArrayCode & "</li>"
                 End If
              Else
                 FoundErr=True
                 ErrMsg=ErrMsg & "<br><li>在获取链接列表时出错。</li>"
              End If   
           Else
               FoundErr=True
              ErrMsg=ErrMsg & "<br><li>在截取列表时发生错误。</li>"
           End If
        Else
            FoundErr=True
           ErrMsg=ErrMsg & "<br><li>在获取:" & ListUrl & "网页源码时发生错误。</li>"
        End If
     End If

If FoundErr<>True Then
   NewsCode=GetHttpPage(NewsUrl)
   If NewsCode<>"$False$" Then
      Title=GetBody(NewsCode,TsString,ToString,False,False)
      Content=GetBody(NewsCode,CsString,CoString,False,False)
      If Title="$False$" or  Content="$False$"  Then
         FoundErr=True
		 If Title="$False$" Then  ErrMsg=ErrMsg & "<br><li>在截取信息标题的时候发生错误：" & NewsUrl & "</li>"
		 If Content="$False$" Then ErrMsg=ErrMsg & "<br><li>在截取信息正文的时候发生错误：" & len(NewsCode) & "</li>"
      Else
         Title=FpHtmlEnCode(Title)
         Title=dvhtmlencode(Title)

      If DateType=0 then
         DateTime=date()
      Else
         DateTime=GetBody(NewsCode,DsString,DoString,False,False)
         DateTime=FpHtmlEncode(DateTime)
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

         '作者
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
         Tel=GetBody(NewsCode,TelsString,TeloString,False,False)
      ElseIf TelType=2 Then
         Tel=TelStr
      End If
      If Tel="$False$" Or IsNull(Tel) Then
         Tel="未填"
      Else
         Tel=FpHtmlEnCode(Tel)
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
         QQ=GetBody(NewsCode,QsString,QoString,False,False)
      ElseIf QQType=2 Then
         QQ=QQStr
      End If
      If QQ="$False$" Or IsNull(QQ) Then
         QQ="未填"
      Else
         QQ=FpHtmlEnCode(QQ)
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

         '来源
         If CopyFromType=1 Then
            CopyFrom=GetBody(NewsCode,FsString,FoString,False,False)
         ElseIf CopyFromType=2 Then
            CopyFrom=CopyFromStr
         End If
         If CopyFrom="$False$" Or IsNull(CopyFrom) Then
            CopyFrom="不详"
         Else
            CopyFrom=FpHtmlEnCode(CopyFrom)
         End If

      If leixingType=1 Then
         leixing=GetBody(NewsCode,LxsString,LxoString,False,False)
      ElseIf leixingType=2 Then
         leixing=leixingStr
      End If
      If leixing="$False$" Or IsNull(leixing) Then
         leixing="未填"
      Else
         leixing=FpHtmlEnCode(leixing)
      End If

     End  If    
   Else
     FoundErr=True
     ErrMsg=ErrMsg & "<br><li>在获取源码时发生错误："& NewsUrl &"</li>"
   End If 
End If

If FoundErr<>True Then
   Call GetFilters
   Call Filters
   Content=ReplaceSaveRemoteFile(Content,cj_strInstallDir,strChannelDir,False,NewsUrl)
End If

End Sub


'==================================================
'过程名：GetFilters
'作  用：提取过滤信息
'参  数：无
'==================================================
Sub GetFilters()
   SqlF ="Select * from DNJY_xx_Filters Where Flag="&DN_True&" And (PublicTf="&DN_True&" Or ItemID=" & ItemID & ") order by FilterID ASC"
   Set RSF=Conn.Execute(SqlF)
   If RsF.Eof And RsF.Bof Then
      Arr_Filters=""
   Else
      Arr_Filters=RsF.GetRows()
   End If
   RsF.Close
   Set RsF=Nothing
End Sub


'==================================================
'过程名：Filters
'作  用：过滤
'==================================================
Sub Filters()
If IsArray(Arr_Filters)=False Then
   Exit Sub
End if

   For Filteri=0 to Ubound(Arr_Filters,2)
      FilterStr=""
      If Arr_Filters(1,Filteri)=ItemID Or Arr_Filters(10,Filteri)=True Then
         If Arr_Filters(3,Filteri)=1 Then'标题过滤
            If Arr_Filters(4,Filteri)=1 Then
               Title=Replace(Title,Arr_Filters(5,Filteri),Arr_Filters(8,Filteri))
            ElseIf Arr_Filters(4,Filteri)=2 Then
               FilterStr=GetBody(Title,Arr_Filters(6,Filteri),Arr_Filters(7,Filteri),True,True)
               Do While FilterStr<>"$False$"
                  Title=Replace(Title,FilterStr,Arr_Filters(8,Filteri))
                  FilterStr=GetBody(Title,Arr_Filters(6,Filteri),Arr_Filters(7,Filteri),True,True)
               Loop
            End If
         ElseIf Arr_Filters(3,Filteri)=2 Then'正文过滤
            If Arr_Filters(4,Filteri)=1 Then
               Content=Replace(Content,Arr_Filters(5,Filteri),Arr_Filters(8,Filteri))
            ElseIf Arr_Filters(4,Filteri)=2 Then
               FilterStr=GetBody(Content,Arr_Filters(6,Filteri),Arr_Filters(7,Filteri),True,True)
               Do While FilterStr<>"$False$"
                  Content=Replace(Content,FilterStr,Arr_Filters(8,Filteri))
                  FilterStr=GetBody(Content,Arr_Filters(6,Filteri),Arr_Filters(7,Filteri),True,True)
               Loop
            End If
         End If
      End If
   Next
End Sub
%>


