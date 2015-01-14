<!--#include file="../conn.asp"-->
<!--#include file="../config.asp"-->
<!--#include file=cookies.asp-->
<!--#include file="inc/Article_Function.asp"-->
<!--#include file="inc/clsCache.asp"-->
<%
'=====================================
'天天供求信息网站管理系统
'天天网络科技工作室出品
'客服中心http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>

<%
'option explicit
'Response.Buffer = True

Dim SqlItem,RsItem,Action,ItemID,FilterName,FilterID,FilterObject,FilterType,FilterContent,FisString,FioString,FilterRep,Flag,PublicTf
Dim FoundErr,ErrMsg
Action=Trim(Request("Action"))
FilterID=Trim(Request("FilterID"))
Call OpenConn
If FilterID="" Then
   FoundErr=True
   ErrMsg=ErrMsg & "<br><li>参数错误，请从有效链接进入</li>" 
Else
   FilterID=Clng(FilterID)
End If


If Action="SaveEdit" and FoundErr<>True Then
   Call SaveEdit
Else
   Call Test()
End If
If FoundErr<>True Then
   Call Main()
Else
   Call WriteErrMsg(ErrMsg)
End If
'关闭数据库链接
Conn.Close:Set Conn=Nothing

Sub Main
%>
<html>
<head>
<title>采集系统</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="../css/manage.css" type="text/css">
<SCRIPT language=javascript>
function showset(thisform)
{
		if(thisform.FilterType.selectedIndex==1)
		{
        	FilterType1.style.display = "none";
			FilterType2.style.display = "";
		}
		else
		{
        	FilterType1.style.display = "";
			FilterType2.style.display = "none";
	    }

}
</script>
</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" class="border">
  <tr class='topbg'> 
    <td height="22" colspan="2" align="center" ><strong>采 集 系 统 过 滤 管 理</strong></td>
  </tr>
</table>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" class="border">
  <tr class="tdbg"> 
    <td width="65" height="30"><strong>管理导航：</strong></td>
    <td height="30"><a href=cj_Settings.asp>管理首页</a> | <a href="cj_filter.asp">添加新过滤</a></td>
  </tr>
</table>
<br>

<form method="post" action="cj_modify.asp" name="form1">
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" class="border" >                                
    <tr>                                 
      <td height="22" colspan="2" class="title"> <div align="center"><strong>编  辑 过 滤</strong></div></td>                                
    </tr>                                
    <tr class="tdbg">                                 
      <td width="150" class="tdbg"><strong> 过滤名称：</strong></td>                                
      <td class="tdbg"><input name="FilterName" type="text" id="FilterName" value="<%=FilterName%>" size="25" maxlength="30">                                 
        &nbsp;</td>                                
    </tr>                                
    <tr class="tdbg">                                 
      <td width="150" class="tdbg"><strong> 所属项目：</strong></td>                                
      <td class="tdbg"><%Call Admin_ShowItem_Option(ItemID)%>                                
      </td>                                
    </tr>                                
    <tr class="tdbg">                                 
      <td width="150" class="tdbg"><strong> 过滤对象：<strong></td>                               
      <td class="tdbg">                               
         <select name="FilterObject" id="FilterObject">                              
            <option value="1" <%if FilterObject=1 Then Response.Write "selected"%>>标题过滤</option>                              
            <option value="2" <%if FilterObject=2 Then Response.Write "selected"%>>正文过滤</option>                              
         </select>                              
      </td>                               
    </tr>                               
    <tr class="tdbg">                                
      <td width="150" class="tdbg"><strong> 过滤类型：</strong></td>                               
      <td class="tdbg">                               
         <select name="FilterType" id="FilterType" onchange=showset(this.form)>                              
            <option value="1" <%if FilterType=1 Then Response.Write "selected"%> >简单替换</option>                              
            <option value="2" <%if FilterType=2 Then Response.Write "selected"%>>高级过滤</option>                              
         </select>                              
      </td>                                
    </tr>     
    <tr class="tdbg">                                
      <td width="150" class="tdbg"><strong> 使用状态：</strong></td>                               
      <td class="tdbg">                               
         <select name="Flag" id="Flag">                              
            <option value="yes" <%If Flag=True Then Response.Write "selected"%>>启用</option>                              
            <option value="no" <%If Flag=False Then Response.Write "selected"%>>禁用</option>                              
         </select>                              
      </td>                                
    </tr>                                     
    <tr class="tdbg">                                
      <td width="150" class="tdbg"><strong> 使用范围：</strong></td>                               
      <td class="tdbg">                               
         <select name="PublicTf" id="PublicTf">                              
            <option value="yes" <%If PublicTf=True Then Response.Write "selected"%>>公有</option>                            
            <option value="no" <%If PublicTf=False Then Response.Write "selected"%>>私有</option>                            
         </select>                              
      </td>                                
    </tr>                                     
</table>                                
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" class="border" id="FilterType1" style="display:<%if FilterType<>1 Then Response.Write "none"%>">                                
    <tr class="tdbg">                                 
      <td width="150" class="tdbg"><strong> 内容：</strong></td>                               
      <td class="tdbg"><textarea name="FilterContent" cols="49" rows="5"><%=FilterContent%></textarea>                                
      </td>                               
    </tr>                            
</table>                 
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" class="border" id="FilterType2" style="display:<%if FilterType<>2 Then Response.Write "none"%>">                                   
    <tr class="tdbg">                               
      <td width="150" class="tdbg"><strong> 开始标记：</strong></td>                              
      <td class="tdbg"><textarea name="FisString" cols="49" rows="5"><%=FisString%></textarea>                               
        &nbsp;</td>                                
    </tr>                                
    <tr class="tdbg">                                 
      <td width="150" class="tdbg"><strong> 结束标记：</strong></td>                               
      <td class="tdbg"><textarea name="FioString" cols="49" rows="5"><%=FioString%></textarea>                                
        &nbsp;</td>                                
    </tr>                         
</table>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" class="border" >
    <tr class="tdbg"> 
      <td width="150" class="tdbg"><strong> 替换：</strong></td>
      <td class="tdbg"><textarea name="FilterRep" cols="49" rows="5" id="FilterRep"><%=FilterRep%></textarea></td>
    </tr>
</table>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" class="border" >
    <tr class="tdbg"> 
      <td colspan="2" align="center" class="tdbg">
        <input name="Action" type="hidden" id="Action" value="SaveEdit">
        <input name="FilterID" type="hidden" id="FilterID" value="<%=FilterID%>">
        <input name="Cancel" type="button" id="Cancel" value="返&nbsp;&nbsp;回" onClick="window.location.href='cj_Settings.asp'" style="cursor: hand;background-color: #cccccc;">
        &nbsp;
        <input  type="submit" name="Submit" value="确&nbsp;&nbsp;定" style="cursor: hand;background-color: #cccccc;"> 
        </td>
    </tr>
</table>
</form>          
</body>         
</html>
<%End Sub%>

<%
Sub SaveEdit
FilterID=Trim(Request("FilterID"))
FilterName=Trim(Request.Form("FilterName"))
ItemID=Trim(Request.Form("ItemID"))
FilterObject=Trim(Request.Form("FilterObject"))
FilterType=Trim(Request.Form("FilterType"))
FilterContent=Request.Form("FilterContent")
FisString=Request.Form("FisString")
FioString=Request.Form("FioString")
FilterRep=Request.Form("FilterRep")
Flag=Trim(Request.Form("Flag"))
PublicTf=Trim(Request.Form("PublicTf"))

If FilterID="" Then                                
   FoundErr=True                                
   ErrMsg=ErrMsg & "<br><li>参数错误！请从有效链接进入</li>"
Else
   FilterID=Clng(FilterID)                                 
End If                                

If FilterName="" Then                                
   FoundErr=True                                
   ErrMsg=ErrMsg & "<br><li>过滤名称不能为空</li>"                                 
End If                                
If ItemID="" Then                                
   FoundErr=True                                
   ErrMsg=ErrMsg & "<br><li>请选择过滤所属项目</li>"                                 
Else                                
   ItemID=Clng(ItemID)  
   If ItemID=0 Then  
      FoundErr=True                                
      ErrMsg=ErrMsg & "<br><li>请选择过滤所属项目</li>"   
   End If                                
End If       
If FilterObject="" Then                                
   FoundErr=True                                
   ErrMsg=ErrMsg & "<br><li>请选择过滤对象</li>"                                 
Else                                
   FilterObject=Clng(FilterObject)   
End If                                
                            
If FilterType="" Then                                
   FoundErr=True                                
   ErrMsg=ErrMsg & "<br><li>请选择过滤类型</li>"                                 
Else                                
   FilterType=Clng(FilterType)                                
   If FilterType=1 Then     
      If FilterContent="" Then                              
         FoundErr=True                                
         ErrMsg=ErrMsg & "<br><li>过滤的内容不能为空</li>"   
      End If   
   ElseIf FilterType=2 Then   
      If FisString="" or FioString="" Then   
         FoundErr=True                                
         ErrMsg=ErrMsg & "<br><li>开始/结束标记不能为空</li>"   
      End If   
   Else   
      FoundErr=True                                
      ErrMsg=ErrMsg & "<br><li>参数错误，请从有效链接进入</li>"   
   End If   
End If                     
If Flag="yes" Then
   Flag=True
Else
   Flag=False
End If
If PublicTf="yes" Then
   PublicTf=True
Else
   PublicTf=False
End If 

If FoundErr<>True Then                                
   SqlItem ="select top 1 *  from DNJY_Article_Filters Where FilterID=" & FilterID                                
   Set RsItem=server.CreateObject("adodb.recordset")                                
   RsItem.open SqlItem,Conn,2,3                                                              
   RsItem("FilterName")=FilterName                                
   RsItem("ItemID")=ItemID   
   RsItem("FilterObject")=FilterObject                                
   RsItem("FilterType")=FilterType                                
   If FilterType=1 Then                                
      RsItem("FilterContent")=FilterContent                                
   ElseIf FilterType=2 Then                                
      RsItem("FisString")=FisString                                
      RsItem("FioString")=FioString                                
   End If                                                
   RsItem("FilterRep")=FilterRep        
   RsItem("Flag")=Flag   
   RsItem("PublicTf")=PublicTf                        
   RsItem.Update                                
   RsItem.Close                                
   Set RsItem=Nothing                                
   Response.Redirect "cj_Settings.asp"                                
Else                                
   Call WriteErrMsg(ErrMsg)                                
End If                  
End Sub

Sub Test
Set RsItem=server.createobject("adodb.recordset")         
SqlItem="select * from DNJY_Article_Filters Where FilterID=" & FilterID         
RsItem.open SqlItem,Conn,1,1      
If Not RsItem.Eof Then
   ItemID=RsItem("ItemID")
   FilterName=RsItem("FilterName")
   FilterObject=RsItem("FilterObject")
   FilterType=RsItem("FilterType")
   FilterContent=RsItem("FilterContent")
   FisString=RsItem("FisString")
   FioString=RsItem("FioString")
   FilterRep=RsItem("FilterRep")
   Flag=RsItem("Flag")
   PublicTf=RsItem("PublicTf")
Else
   FoundErr=True
   ErrMsg=ErrMsg & "<br></li>参数错误，找不到该项目</li>"
End If
RsItem.Close
Set RsItem=Nothing
End Sub
%>
