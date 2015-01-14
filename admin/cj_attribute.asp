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
Dim SqlItem,RsItem,Action,FoundErr,ErrMsg
Dim ItemID,ItemName
Dim  Hits,UpdateType,UpdateTime,IncludePicYn,DefaultPicYn,OnTop,Hot
Dim  Script_Iframe,Script_Object,Script_Script,Script_Div,Script_Class,Script_Span,Script_Img,Script_Font,Script_A,Script_Html,Script_Table,Script_Tr,Script_Td
Dim  CollecListNum,CollecNewsNum,Passed,SaveFiles,CollecOrder,LinkUrlYn,InputerType,Inputer,EditorType,Editor,ShowCommentLink,cj_watermark
FoundErr=False
ItemID=strint(Request("ItemID"))
Action=Trim(Request("Action"))
If ItemID=0 Then
   FoundErr=True
   ErrMsg=ErrMsg & "<br><li>参数错误，项目ID不能为空！</li>"
End If

If  FoundErr<>True  Then

Call OpenConn
      Call  GetTest()
conn.close:SEt conn=nothing
End  If
If FoundErr<>True Then
Call OpenConn
   Call Main()
End If

'关闭数据库链接
conn.close:SEt conn=nothing


Sub Main
%>
<html>
<head>
<title>采集系统</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="../css/manage.css" type="text/css">
</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" class="border">
  <tr> 
    <td height="22" colspan="2" align="center" class="topbg"><strong>新闻采集系统模板管理</td>
  </tr>
  <tr class="tdbg"> 
    <td width="65" height="30"><strong>管理导航：</strong></td>
    <td height="30">项目编辑 >> <a href="cj_editor_a.asp?ItemID=<%=ItemID%>">基本设置</a> >> <a href="cj_editor_b.asp?ItemID=<%=ItemID%>">列表设置</a> >> <a href="cj_editor_c.asp?ItemID=<%=ItemID%>">链接设置</a> >> <a href="cj_editor_d.asp?ItemID=<%=ItemID%>">正文设置</a> >>  
	<a href="cj_editor_e.asp?ItemID=<%=ItemID%>">采样测试</a> >> <a href="cj_attribute.asp?ItemID=<%=ItemID%>"><font color=red>属性设置</font></a> >> 完成</td>         
  </tr>         
</table>
<br>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" class="border">
    <tr> 
      <td height="22" colspan="2" class="title"> <div align="center"><strong>编 辑 项 目--属 性 设 置</strong></div></td>
    </tr>
</table>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" class="border,tdbg">
<form method="post" action="cj_success.asp" name="myform">
    <tr class="tdbg">
      <td height="30" width="20%" align="right"><b>点击数初始值：</b></td>
      <td><input name="Hits" type="text" id="Hits" value="<%=Hits%>" size="10" maxlength="10">
        <font color='#0000FF'>这功能是提供给管理员作弊用的。不过尽量不要用呀！^_^</font>      </td>
    </tr>
    <tr class="tdbg">
      <td height="30" width="20%" align="right"><b>新闻性质：</b></td>
      <td> 
        <input name="IncludePicYn" type="checkbox" value="yes" <%If IncludePicYn=-1 Then Response.write "checked"%> style="border: 0px;background-color: #eeeeee;">
        包含图片
        <input name="DefaultPicYn" type="checkbox" value="yes" <%If DefaultPicYn=-1 Then Response.write "checked"%> style="border: 0px;background-color: #eeeeee;">
        推荐新闻
        <input name="OnTop" type="checkbox" value="yes" <%If OnTop=-1 Then Response.write "checked"%> style="border: 0px;background-color: #eeeeee;">
        固顶新闻
		<input name="Hot" type="checkbox" value="yes" <%If Hot=-1 Then Response.write "checked"%> style="border: 0px;background-color: #eeeeee;">
        热点新闻</td>
    </tr>  

    <tr class="tdbg">
      <td height="30" width="20%" align="right"><b>标签过滤：</b></td>
      <td>
        <input name="Script_Iframe" type="checkbox" value="yes" <%If Script_Iframe=-1 Then response.write "checked"%> style="border: 0px;background-color: #eeeeee;">
        Iframe
        <input name="Script_Object" type="checkbox" value="yes" <%If Script_Object=-1 Then response.write "checked"%> onclick='return confirm("确定要选择该标记吗？这将删除正文中的所有Object标记，结果将导致该文章中的所有Flash动画被删除！");' style="border: 0px;background-color: #eeeeee;">
        Object
        <input name="Script_Script" type="checkbox" value="yes" <%If Script_Script=-1 Then response.write "checked"%> style="border: 0px;background-color: #eeeeee;">
        Script
        <input name="Script_Div" type="checkbox"  value="yes" <%If Script_Div=-1 Then response.write "checked"%> style="border: 0px;background-color: #eeeeee;">
        Div
        <input name="Script_Class" type="checkbox"  value="yes" <%If Script_Class=-1 Then response.write "checked"%> style="border: 0px;background-color: #eeeeee;">
        Class
        <input name="Script_Table" type="checkbox"  value="yes" <%If Script_Table=-1 Then response.write "checked"%> style="border: 0px;background-color: #eeeeee;">
        Table
        <input name="Script_Tr" type="checkbox"  value="yes" <%If Script_Tr=-1 Then response.write "checked"%> style="border: 0px;background-color: #eeeeee;">
        Tr
        <br>
        <input name="Script_Span" type="checkbox"  value="yes" <%If Script_Span=-1 Then response.write "checked"%> style="border: 0px;background-color: #eeeeee;">
        Span&nbsp;&nbsp;
        <input name="Script_Img" type="checkbox" value="yes" <%If Script_Img=-1 Then response.write "checked"%> style="border: 0px;background-color: #eeeeee;">
        Img&nbsp;&nbsp;&nbsp;
        <input name="Script_Font" type="checkbox"  value="yes" <%If Script_Font=-1 Then response.write "checked"%> style="border: 0px;background-color: #eeeeee;">
        Font&nbsp;&nbsp;
        <input name="Script_A" type="checkbox" value="yes" <%If Script_A=-1 Then response.write "checked"%> style="border: 0px;background-color: #eeeeee;">
        A&nbsp;&nbsp;
        <input name="Script_Html" type="checkbox" value="yes" <%If Script_Html=-1 Then response.write "checked"%> onclick='return confirm("确定要选择该标记吗？这将删除正文中的所有Html标记，结果将导致该文章的可阅读性降低！");' style="border: 0px;background-color: #eeeeee;">
        Html&nbsp;
        <input name="Script_Td" type="checkbox"  value="yes" <%If Script_Td=-1 Then response.write "checked"%> style="border: 0px;background-color: #eeeeee;">
        Td        </td>
    </tr>
    <tr class="tdbg">
      <td height="30" width="20%" align="right"> <b>列表深度：</b></td>
      <td>
        <input name="CollecListNum" type="text" id="CollecListNum" value="<%=CollecListNum%>" size="10" maxlength="10">&nbsp;&nbsp;&nbsp;
		<font color='#0000FF'>0为所有的列表</font></td>
    </tr>
    <tr class="tdbg">
      <td height="30" width="20%" align="right"> <b>新闻数量：</b></td>
      <td>
        <input name="CollecNewsNum" type="text" id="CollecNewsNum" value="<%=CollecNewsNum%>" size="10" maxlength="10">&nbsp;&nbsp;&nbsp; 
		<font color='#0000FF'>0为所有的新闻<span lang="en-us">(</span>每一列表的新闻限制数量<span lang="en-us">)</span></font></td>
    </tr>
    <tr class="tdbg">
      <td height="30" width="20%" align="right"> <b>采集选项：</b></td>
      <td><input name="Passed" type="checkbox" value="yes" <%if Passed=-1 then response.write "checked"%> style="border: 0px;background-color: #eeeeee;">
        自动审核		
        <input name="SaveFiles" type="checkbox" value="yes" <%if SaveFiles=-1 then response.write "checked"%> <%If IsObjInstalled("Scripting.FileSystemObject")=False Then Response.Write "disabled"%> style="border: 0px;background-color: #eeeeee;">
        保存图片
		<input name="cj_watermark" type="checkbox" value="yes" <%if cj_watermark=-1 then response.write "checked"%> <%If IsObjInstalled("Scripting.FileSystemObject")=False Then Response.Write "disabled"%> style="border: 0px;background-color: #eeeeee;">
        图片水印
        <input name="CollecOrder" type="checkbox" value="yes" <%if CollecOrder=-1 then response.write "checked"%> style="border: 0px;background-color: #eeeeee;">
        倒序采集
        <input name="LinkUrlYn" type="checkbox" value="yes" <%if LinkUrlYn=-1 then response.write "checked"%> style="border: 0px;background-color: #eeeeee;">
        URL跳转</td>
    </tr>
    <tr class="tdbg">
    <td height="30" width="20%" align="right"></td>
    <td align-"center"><center>
       <input type="hidden" value="<%=ItemID%>" name="ItemID">        
       <input type="submit" value=" 完&nbsp;&nbsp;成 " name="submit" style="cursor: hand;background-color: #cccccc;">  </center>       </td> 
    </tr>              
</form>         
</table>  
</body>         
</html>
<%End Sub%>
<%
Sub GetTest()
   SqlItem="Select * from DNJY_Article_Item Where ItemID=" & ItemID
   Set RsItem=server.CreateObject("adodb.recordset")
   RsItem.Open SqlItem,Conn,1,1
   If  RsItem.Eof  And  RsItem.Bof  Then
      FoundErr=True
      ErrMsg=ErrMsg & "<br><li>参数错误，找不到该项目</li>"
   Else
      Hits=RsItem("Hits")
      UpdateType=RsItem("UpdateType")
      UpdateTime=RsItem("UpdateTime")
      IncludePicYn=RsItem("IncludePicYn")
      DefaultPicYn=RsItem("DefaultPicYn")
      OnTop=RsItem("OnTop")
      Hot=RsItem("Hot")
      Script_Iframe=RsItem("Script_Iframe")
      Script_Object=RsItem("Script_Object")
      Script_Script=RsItem("Script_Script")
      Script_Div=RsItem("Script_Div")
      Script_Class=RsItem("Script_Class")
      Script_Span=RsItem("Script_Span")
      Script_Img=RsItem("Script_Img")
      Script_Font=RsItem("Script_Font")
      Script_A=RsItem("Script_A")
      Script_Html=RsItem("Script_Html")
      Passed=RsItem("Passed")
      SaveFiles=RsItem("SaveFiles")
      CollecOrder=RsItem("CollecOrder")
      LinkUrlYn=RsItem("LinkUrlYn")
      CollecListNum=RsItem("CollecListNum")
      CollecNewsNum=RsItem("CollecNewsNum")
      InputerType=RsItem("InputerType")
      Inputer=RsItem("Inputer")
      EditorType=RsItem("EditorType")
      Editor=RsItem("Editor")
      ShowCommentLink=RsItem("ShowCommentLink")
      Script_Table=RsItem("Script_Table")
      Script_Tr=RsItem("Script_Tr")
      Script_Td=RsItem("Script_Td")
	  cj_watermark=RsItem("cj_watermark")
   End  If
   RsItem.Close
   Set RsItem=Nothing
End  Sub
%>
