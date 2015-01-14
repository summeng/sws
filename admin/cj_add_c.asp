<!--#include file="../conn.asp"-->
<!--#include file=cookies.asp-->
<!--#include file=../config.asp-->
<!--#include file="../Include/head.asp"-->
<!--#include file="../Include/Function.asp"-->
<!--#include file="../Include/label.asp"-->
<!--#include file="inc/Article_Function.asp"-->
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
<html>
<head>
<title>采集系统</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="../css/manage.css" type="text/css">
</head>
<%Call OpenConn
Dim SqlItem,RsItem,ItemID,FoundErr,ErrMsg
Dim ListStr,LsString,LoString,ListPaingType,LPsString,LPoString,ListPaingStr1,ListPaingStr2,ListPaingID1,ListPaingID2,ListPaingStr3
Dim ListUrl,ListCode
Dim LoginType,LoginUrl,LoginPostUrl,LoginUser,LoginPass,LoginFalse,LoginResult,LoginData
Dim  ListPaingNext
ItemID=Trim(Request.Form("ItemID"))
ListStr=Trim(Request.Form("ListStr"))
LsString=Request.Form("LsString")
LoString=Request.Form("LoString")
ListPaingType=Request.Form("ListPaingType")
LPsString=Request.Form("LPsString")
LPoString=Request.Form("LPoString")
ListPaingStr1=Trim(Request.Form("ListPaingStr1"))
ListPaingStr2=Trim(Request.Form("ListPaingStr2"))
ListPaingID1=Trim(Request.Form("ListPaingID1"))
ListPaingID2=Trim(Request.Form("ListPaingID2"))
ListPaingStr3=Trim(Request.Form("ListPaingStr3"))
FoundErr=False
If ItemID=""  Then
   FoundErr=True
   ErrMsg=ErrMsg & "<br><li>参数错误，请从有效链接进入</li>"
Else
   ItemID=Clng(ItemID)
End If
If LsString="" Then
   FoundErr=True
   ErrMsg=ErrMsg & "<br><li>列表开始标记不能为空</li>"
End If
If LoString="" Then
   FoundErr=True
   ErrMsg=ErrMsg & "<br><li>列表结束标记不能为空</li>" 
End If
If ListPaingType="" Then
   FoundErr=True
   ErrMsg=ErrMsg & "<br><li>请选择列表索引分页类型</li>" 
Else
   ListPaingType=Clng(ListPaingType)
   Select Case ListPaingType
   Case 0,1
            If ListStr="" Then
               FoundErr=True
               ErrMsg=ErrMsg & "<br><li>列表索引页不能为空</li>"
            Else
               ListStr=Trim(ListStr)
            End If
      If  ListPaingType=1  Then
            If LPsString="" or LPoString="" Then
               FoundErr=True
               ErrMsg=ErrMsg & "<br><li>索引分页开始/结束标记不能为空</li>" 
            End If
            If ListPaingStr1<>"" and Len(ListPaingStr1)<15 Then
               FoundErr=True
               ErrMsg=ErrMsg & "<br><li>索引分页重定向设置不正确(至少15个字符)</li>" 
            End If
      End  If
   Case 2
      If ListPaingStr2="" Then
         FoundErr=True
         ErrMsg=ErrMsg & "<br><li>批量生成字符不能为空</li>"
      End If
      If isNumeric(ListPaingID1)=False or isNumeric(ListPaingID2)=False Then
         FoundErr=True
         ErrMsg=ErrMsg & "<br><li>批量生成的范围只能是数字</li>"
      Else
         ListPaingID1=Clng(ListPaingID1)
         ListPaingID2=Clng(ListPaingID2)
         If ListPaingID1=0 And ListPaingID2=0 Then
            FoundErr=True
            ErrMsg=ErrMsg & "<br><li>批量生成范围设置不正确</li>"
         End If
      End If
   Case 3
      If ListPaingStr3="" Then
         FoundErr=True
         ErrMsg=ErrMsg & "<br><li>列表索引分页不能为空，请手动添加</li>"
      Else
         ListPaingStr3=Replace(ListPaingStr3,CHR(13),"|") 
      End If
   Case Else
      FoundErr=True
      ErrMsg=ErrMsg & "<br><li>请选择列表索引分页类型</li>" 
   End Select
End if


If FoundErr<>True Then
   SqlItem="Select * from DNJY_Article_Item Where ItemID=" & ItemID
   Set RsItem=server.CreateObject("adodb.recordset")
   RsItem.Open SqlItem,Conn,2,3

   RsItem("LsString")=LsString
   RsItem("LoString")=LoString
   RsItem("ListPaingType")=ListPaingType
   Select Case ListPaingType
   Case 0,1
         RsItem("ListStr")=ListStr
      If ListPaingType=1  Then
            RsItem("LPsString")=LPsString
            RsItem("LPoString")=LPoString
            If ListPaingStr1<>"" Then
               RsItem("ListPaingStr1")=ListPaingStr1
            End If
      End  If
      ListUrl=ListStr
   Case 2
      RsItem("ListPaingStr2")=ListPaingStr2
      RsItem("ListPaingID1")=ListPaingID1
      RsItem("ListPaingID2")=ListPaingID2
      ListUrl=Replace(ListPaingStr2,"{$ID}",CStr(ListPaingID1))
   Case 3
      RsItem("ListPaingStr3")=ListPaingStr3
      If  Instr(ListPaingStr3,"|")>0  Then
            ListUrl=Left(ListPaingStr3,Instr(ListPaingStr3,"|")-1)
      Else
            ListUrl=ListPaingStr3
      End  If
   End Select
   LoginType=RsItem("LoginType")
   LoginUrl=RsItem("LoginUrl")
   LoginPostUrl=RsItem("LoginPostUrl")
   LoginUser=RsItem("LoginUser")
   LoginPass=RsItem("LoginPass")
   LoginFalse=RsItem("LoginFalse")
   RsItem.UpDate
   RsItem.Close
   Set RsItem=Nothing

   If LoginType=1 then
      LoginData=UrlEncoding(LoginUser & "&" & LoginPass)
      LoginResult=PostHttpPage(LoginUrl,LoginPostUrl,LoginData)
      If Instr(LoginResult,LoginFalse)>0 Then
         FoundErr=True
         ErrMsg=ErrMsg & "<br><li>网站登录失败，请检查登录参数！</li>"
      End If
   End If
   
   If FoundErr<>True Then
      ListCode=GetHttpPage(ListUrl)
      If ListCode<>"$False$" Then
         If ListPaingType=1  Then
            ListPaingNext=GetPaing(ListCode,LPsString,LPoString,False,False)
                  If ListPaingNext<>"$False$"  Then
                     If ListPaingStr1<>""  Then  
                        ListPaingNext=Replace(ListPaingStr1,"{$ID}",ListPaingNext)
               Else
                        ListPaingNext=DefiniteUrl(ListPaingNext,ListUrl)
               End  If
            End  If
         End If
         ListCode=GetBody(ListCode,LsString,Lostring,False,False)
         If ListCode="$False$" Then
            FoundErr=True
            ErrMsg=ErrMsg & "<br><li>在截取:" & ListUrl & "新闻列表时发生错误</li>"
         End If
      Else
         FoundErr=True
         ErrMsg=ErrMsg & "<br><li>在获取:" & ListUrl & "网页源码时发生错误</li>"
      End If
   End If
End If

If FoundErr=True Then
   Call WriteErrMsg(ErrMsg)
Else
   Call Main
End If   
'关闭数据库链接
conn.close:Set conn=nothing
%>

<%Sub Main%>
<html>
<head>
<title>采集系统</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" type="text/css" href="Admin_Style.css">
</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" class="border">
  <tr> 
    <td height="22" colspan="2" align="center" class="topbg"><strong>新闻采集系统项目管理</td>
  </tr>
  <tr class="tdbg"> 
    <td width="65" height="30"><strong>管理导航：</strong></td>
    <td height="30"><a href="cj_add_a.asp">添加项目</a> >> <a href="cj_editor_a.asp?ItemID=<%=ItemID%>">基本设置</a> >> <a href="cj_editor_b.asp?ItemID=<%=ItemID%>">列表设置</a> >> <font color=red>链接设置</font> >> 正文设置 >> 采样测试 >> 属性设置 >> 完成</td>         
  </tr>         
</table>         
<br>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" class="border" >
    <tr> 
      <td height="22" colspan="2" class="title"> <div align="center"><strong>添 加 新 
		项 目--列 表 截 取 测 试</strong></div></td>
    </tr>
    <tr class="tdbg"> 
      <td height="22" colspan="2" class="tdbg">
       <%=ListCode%>
      </td>
    </tr>
</table>
<%If ListPaingNext<>"" And ListPaingNext<>"$False$" Then%>
<br>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" class="border" >
    <tr class="tdbg"> 
      <td height="22" colspan="2" class="tdbg" >
      <%Response.Write "<br>下一页列表：<a  href='" & ListPaingNext  &  "' target=_blank><font  color=red>"  &  ListPaingNext  &  "</font></a>"%>
      </td>
    </tr>
</table>
<%End If%>
<br>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" class="border" >
<form method="post" action="cj_add_d.asp" name="form1">
    <tr> 
      <td height="22" colspan="2" class="title"> <div align="center"><strong>添 加 新 
		项 目--链 接 设 置</strong></div></td>
    </tr>
    <tr class="tdbg"> 
      <td width="20%" class="tdbg"  align="right"><strong>链接开始标记：</strong></td>
      <td class="tdbg" width="75%">
      <textarea name="HsString" cols="49" rows="7"></textarea></td>
    </tr>
    <tr class="tdbg"> 
      <td width="20%" class="tdbg"  align="right"><strong>链接结束标记：</strong></td>
      <td class="tdbg" width="75%">
      <textarea name="HoString" cols="49" rows="7"></textarea></td>
    </tr>
    <tr>
      <td width="20%" class="tdbg" align="right"><strong>链接处理类型：</strong></td>
      <td class="tdbg" width="75%">
		<input type="radio" value="0" name="HttpUrlType" checked onClick="HttpUrl1.style.display='none'">自动处理&nbsp;
		<input type="radio" value="1" name="HttpUrlType" onClick="HttpUrl1.style.display=''">重新定向
	   </td>
    </tr>
	<tr class="tdbg" id="HttpUrl1" style="display:none">
      <td width="20%" class="tdbg" align="right"><strong>重新定向链接字符：</strong></td>
      <td class="tdbg" width="75%">
	<input name="HttpUrlStr" type="text" size="49" maxlength="100" value=""><br>
        格式：http://www.scuta.net/Article_Show.asp?ID={$ID}
      </td>
    </tr>
    <tr class="tdbg"> 
      <td colspan="2" align="center" class="tdbg">
        <input name="ItemID" type="hidden" value="<%=ItemID%>">
        <input  type="button" name="button1" value="上&nbsp;一&nbsp;步" onClick="window.location.href='javascript:history.go(-1)'"  style="cursor: hand;background-color: #cccccc;">
        &nbsp;&nbsp;&nbsp;&nbsp; 
        <input  type="submit" name="Submit" value="下&nbsp;一&nbsp;步" style="cursor: hand;background-color: #cccccc;"></td>
    </tr>
</form>
</table>       
</body>         
</html>
<%End Sub%>
