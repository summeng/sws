<!--#include file="../conn.asp"-->
<!--#include file="../config.asp"-->
<!--#include file=cookies.asp-->
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
<%
Dim SqlItem,RsItem,ItemID,FoundErr,ErrMsg,Action
Dim ListStr,LsString,LoString,ListPaingType,LPsString,LPoString,ListPaingStr1,ListPaingStr2,ListPaingID1,ListPaingID2,ListPaingStr3
Dim LoginType,LoginUrl,LoginPostUrl,LoginUser,LoginPass,LoginFalse,LoginResult,LoginData
Dim HsString,HoString,HttpUrlType,HttpUrlStr
Dim ListUrl,ListCode,ListPaingNext
ItemID=Trim(Request("ItemID"))
Action=Trim(Request("Action"))

Call OpenConn
If ItemID="" Then
   FoundErr=True
   ErrMsg=ErrMsg & "<br><li>参数错误，项目ID不能为空！</li>"
Else
   ItemID=Clng(ItemID)
End If
If Action="SaveEdit" And FoundErr<>True Then
   Call SaveEdit()
End If

If FoundErr<>True Then
   Call GetTest()
End If

If FoundErr=True Then
   Call WriteErrMsg(ErrMsg)
Else
   Call Main()
End If
'关闭数据库链接
conn.close:Set conn=nothing
%>

<%Sub Main%>
<html>
<head>
<title>采集系统</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="../css/manage.css" type="text/css">
</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" class="border">
  <tr> 
    <td height="22" colspan="2" align="center" class="topbg"><strong>新闻采集系统项目管理</td>
  </tr>
  <tr class="tdbg"> 
    <td width="65" height="30"><strong>管理导航：</strong></td>
    <td height="30">项目编辑 >> <a href="cj_editor_a.asp?ItemID=<%=ItemID%>">基本设置</a> >> <a href="cj_editor_b.asp?ItemID=<%=ItemID%>">列表设置</a> >> <a href="cj_editor_c.asp?ItemID=<%=ItemID%>"><font color=red>链接设置</font></a> >> <a href="cj_editor_d.asp?ItemID=<%=ItemID%>">正文设置</a> >>  
	<a href="cj_editor_e.asp?ItemID=<%=ItemID%>">采样测试</a> >> <a href="cj_attribute.asp?ItemID=<%=ItemID%>">属性设置</a> >> 完成</td>         
  </tr>         
</table>         
<br>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" class="border" >
    <tr> 
      <td height="22" colspan="5" class="title"> <div align="center"><strong>编 辑 项 目--列 表 截 取 测 试</strong></div></td>
    </tr>
    <tr class="tdbg"> 
      <td height="22" colspan="5" class="tdbg">
       <%=ListCode%>
      </td>
    </tr>
</table>
<%If ListPaingType=1 Then%>
<br>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" class="border" >
    <tr class="tdbg"> 
      <td height="22" colspan="2" class="tdbg" >
      <%Response.Write "<br>下一页列表：<a  href='" & ListPaingNext  &  "' target=_blank><font  color=red>"  &  ListPaingNext  &  "</font></a>"%>
      </td>
    </tr>
</table>
<br>
<%End If%>

<form method="post" action="cj_editor_d.asp" name="form1">
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" class="border" >
    <tr> 
      <td height="22" colspan="2" class="title"> <div align="center"><strong>编 辑 项 目--链 接 设 置</strong></div></td>
    </tr>
    <tr class="tdbg"> 
      <td width="20%" class="tdbg"  align="right"><strong>链接开始标记：</strong></td>
      <td class="tdbg" width="75%">
      <textarea name="HsString" cols="49" rows="7"><%=HsString%></textarea></td>
    </tr>
    <tr class="tdbg"> 
      <td width="20%" class="tdbg" align="right" ><strong>链接结束标记：</strong></td>
      <td class="tdbg" width="75%">
      <textarea name="HoString" cols="49" rows="7"><%=HoString%></textarea></td>
    </tr>
    <tr>
      <td width="20%" class="tdbg" align="right"><strong> 链接特殊处理：</strong></td>
      <td class="tdbg" width="75%">
		<input type="radio" value="0" name="HttpUrlType" <%If HttpUrlType=0 Then Response.Write "checked"%> onClick="HttpUrl1.style.display='none'"> 自动处理&nbsp;
		<input type="radio" value="1" name="HttpUrlType" <%If HttpUrlType=1 Then Response.Write "checked"%> onClick="HttpUrl1.style.display=''"> 重新定位
	   </td>
    </tr>
	<tr class="tdbg" id="HttpUrl1" style="display:'<%If HttpUrlType=0 Then Response.Write "none"%>'">
      <td width="20%" class="tdbg" align="right"><strong>绝对链接字符：</strong></td>
      <td class="tdbg" width="75%">
		<input name="HttpUrlStr" type="text" size="49" maxlength="100" value="<%=HttpUrlStr%>"></td>
    </tr>
    <tr class="tdbg"> 
      <td colspan="2" align="center" class="tdbg">
        <input name="Action" type="hidden" id="Action" value="SaveEdit">
        <input name="ItemID" type="hidden" id="ItemID" value="<%=ItemID%>">
        <input  type="button" name="button1" value="上&nbsp;一&nbsp;步" onClick="window.location.href='javascript:history.go(-1)'"  style="cursor: hand;background-color: #cccccc;">
        &nbsp;&nbsp;&nbsp;&nbsp; 
        <input  type="submit" name="Submit" value="下&nbsp;一&nbsp;步" style="cursor: hand;background-color: #cccccc;"></td>
    </tr>
</table>
</form>    
</body>         
</html>
<%End Sub%>
<%
Sub SaveEdit
   ListStr=Trim(Request.Form("ListStr"))
   LsString=Request.Form("LsString")
   LoString=Request.Form("LoString")
   ListPaingType=Request.Form("ListPaingType")
   LPsString=Request.Form("LPsString")
   LPoString=Request.Form("LPoString")
   ListPaingStr1=Trim(Request.Form("ListPaingStr1"))
   ListPaingStr2=Trim(Request.Form("ListPaingStr2"))
   ListPaingID1=Request.Form("ListPaingID1")
   ListPaingID2=Request.Form("ListPaingID2")
   ListPaingStr3=Request.Form("ListPaingStr3")

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
   RsItem.Open SqlItem,conn,2,3

   RsItem("LsString")=LsString
   RsItem("LoString")=LoString
   RsItem("ListPaingType")=ListPaingType
   Select Case ListPaingType
   Case 0,1
      RsItem("ListStr")=ListStr
      If ListPaingType=1  Then
         RsItem("LPsString")=LPsString
         RsItem("LPoString")=LPoString
         RsItem("ListPaingStr1")=ListPaingStr1
      End If
   Case 2
      RsItem("ListPaingStr2")=ListPaingStr2
      RsItem("ListPaingID1")=ListPaingID1
      RsItem("ListPaingID2")=ListPaingID2
   Case 3
      RsItem("ListPaingStr3")=ListPaingStr3
   End Select
   RsItem.UpDate
   RsItem.Close
   Set RsItem=Nothing
End If
End Sub


'==================================================
'过程名：GetTest
'作  用：测试
'参  数：无
'==================================================
Sub GetTest
   SqlItem="Select * from DNJY_Article_Item Where ItemID=" & ItemID
   Set RsItem=server.CreateObject("adodb.recordset")
   RsItem.Open SqlItem,conn,1,1
   If RsItem.Eof And RsItem.Bof Then
      FoundErr=True
      ErrMsg=ErrMsg & "<br><li>参数错误，项目ID不能为空</li>"
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
   End If
   RsItem.Close
   Set RsItem=Nothing     
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
            ErrMsg=ErrMsg & "<br><li>索引分页开始/结束标记不能为空！</li>"
         End If
         If ListPaingStr1<>"" And Len(ListPaingStr1)<15 Then
            FoundErr=True
            ErrMsg=ErrMsg & "<br><li>索引分页绝对链接设置不正确(请留空或者字符>15个)！</li>"
         End If
      End If      
   ElseIf ListPaingType=2 Then
      If ListPaingStr2="" Then
         FoundErr=True
         ErrMsg=ErrMsg & "<br><li>批量生成原字符串不能为空！</li>"
      End If
      If IsNumeric(ListPaingID1)=False or IsNumeric(ListPaingID2)=False Then
         FoundErr=True
         ErrMsg=ErrMsg & "<br><li>批量生成的范围不正确！无</li>"
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
      ErrMsg=ErrMsg & "<br><li>参数错误，请选择索引分页类型</li>"
   End If
 
   If LoginType=1 Then
      If LoginUrl="" or LoginPostUrl="" or LoginUser="" Or LoginPass="" Or LoginFalse="" Then
         FoundErr=True
         ErrMsg=ErrMsg & "<br><li>请将登录信息填写完整</li>"
      End If
   End If

   If FoundErr<>True Then
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

      If LoginType=1 then
         LoginData=UrlEncoding(LoginUser & "&" & LoginPass)
         LoginResult=PostHttpPage(LoginUrl,LoginPostUrl,LoginData)
         If Instr(LoginResult,LoginFalse)>0 Then
            FoundErr=True
            ErrMsg=ErrMsg & "<br><li>登录网站时发生错误，请确认登录信息的正确性！</li>"
         End If
      End If
   End If
   If FoundErr<>True Then
      ListCode=GetHttpPage(ListUrl)
      If ListCode<>"$False$" Then
         If ListPaingType=1 Then
            ListPaingNext=GetPaing(ListCode,LPsString,LPoString,False,False)
            If ListPaingNext<>"$False$" then
               If ListPaingStr1<>"" Then
                  ListPaingNext=Replace(ListPaingStr1,"{$ID}",ListPaingNext)
               Else
                  ListPaingNext=DefiniteUrl(ListPaingNext,ListUrl)
               End If
            End If
         End If
         ListCode=GetBody(ListCode,LsString,LoString,False,False)
         If ListCode="$False$" Then
            FoundErr=True
            ErrMsg=ErrMsg & "<br><li>在截取列表时发生错误。</li>"
         End If
      Else
         FoundErr=True
         ErrMsg=ErrMsg & "<br><li>在获取:" & ListUrl & "网页源码时发生错误。</li>"
      End If
   End If 
End Sub
%>
