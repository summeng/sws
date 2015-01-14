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

<%call checkmanage("10")
'option explicit
'Response.Buffer = True
Call OpenConn
Dim SqlItem,RsItem
Dim Action,FoundErr,ErrMsg
Dim FilterID,ItemID,FilterName,FilterObject,FilterType,Flag,PublicTf,FlagName
Dim CurrentPage,AllPage,iItem,ItemNum
Const MaxPerPage=20
Action=Request("Action")



If Action="SetFlag" Then
   Call SetFlag()
End If
If FoundErr=True Then
   Call WriteErrMsg(ErrMsg)
Else
   Call Main()
End If
Conn.Close:Set Conn=Nothing
%>
<%Sub Main%>
<html>
<head>
<title>采集系统</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="../css/manage.css" type="text/css">
<style type="text/css">
.ButtonList {
	BORDER-RIGHT: #000000 2px solid; BORDER-TOP: #ffffff 2px solid; BORDER-LEFT: #ffffff 2px solid; CURSOR: default; BORDER-BOTTOM: #999999 2px solid; BACKGROUND-COLOR: #e6e6e6
}
</style>
<SCRIPT language=javascript>
    function unselectall(thisform){
        if(thisform.chkAll.checked){
            thisform.chkAll.checked = thisform.chkAll.checked&0;
        }   
    }
    function CheckAll(thisform){
        for (var i=0;i<thisform.elements.length;i++){
            var e = thisform.elements[i];
            if (e.Name !="chkAll"&&e.disabled!=true)
                e.checked = thisform.chkAll.checked;
        }
    }
</script>
</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" class="border">
  <tr class='topbg'> 
    <td height="22" colspan="2" align="center" ><strong>新闻采集系统过滤管理</strong></td>
  </tr>
</table>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" class="border">
  <tr class="tdbg"> 
    <td width="65" height="30"><strong>管理导航：</strong></td>
    <td height="30"><a href=cj_Settings.asp>管理首页</a> | <a href="cj_filter.asp">添加新过滤</a></td>
  </tr>
</table>
<br>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" class="border" >
    <tr> 
      <td height="22" colspan="2" class="title"> <div align="center"><b>过&nbsp;滤&nbsp;管&nbsp;理</b></div></td>
    </tr>
</table>
<table class="border" border="0" cellspacing="1" width="100%" cellpadding="0">
<form name="myform" method="POST" action="cj_Settings.asp">
        <TBODY>
        <TR>
          <TD class=ButtonList width="5%" height=22>
            <DIV align=center>选择</DIV></TD>
          <TD class=ButtonList width="17%" height=22>
            <DIV align=center>过滤名称</DIV></TD>
          <TD class=ButtonList width="15%" height=22>
            <DIV align=center>所属项目</DIV></TD>
          <TD class=ButtonList width="12%" height=22>
            <DIV align=center>过滤对象</DIV></TD>
          <TD class=ButtonList width="12%" height=22>
            <DIV align=center>过滤类型</div></TD>
          <TD class=ButtonList width="15%" height=22>
            <DIV align=center>状态</div></TD>
          <TD class=ButtonList height=22>
            <DIV align=center>操作</div></TD>
           </TR>
<%
If Request("page")<>"" then
    CurrentPage=Cint(Request("Page"))
Else
    CurrentPage=1
End if      
set RsItem=server.createobject("adodb.recordset")         
SqlItem="select * from DNJY_Article_Filters order by FilterID "&DN_OrderType&""         
RsItem.open SqlItem,Conn,1,1
If (Not RsItem.Eof) and (Not RsItem.Bof) Then

 RsItem.PageSize=MaxPerPage
 Allpage=RsItem.PageCount
 If Currentpage>Allpage Then Currentpage=1
 ItemNum=RsItem.RecordCount
 RsItem.MoveFirst
 RsItem.AbsolutePage=CurrentPage
 iItem=0
   Do while Not RsItem.Eof
      FilterID=RsItem("FilterID")
      ItemID=RsItem("ItemID")
      FilterName=RsItem("FilterName")
      FilterObject=RsItem("FilterObject")
      FilterType=RsItem("FilterType")
      Flag=RsItem("Flag")
      PublicTf=RsItem("PublicTf")
%>
        <TR class="tdbg">
          <TD class="tdbg" style="WIDTH: 6%" align=center>
            <input type="checkbox" name="FilterID" value="<%=FilterID%>" onClick="unselectall(this.form)">
          </TD>
          <TD class="tdbg" align="center"><%=FilterName%>
          </TD>
          <TD class="tdbg" align="center"><%Call Admin_ShowItem_Name(ItemID)%>
          </TD>
          <TD class="tdbg"  align="center">
          <%
          If FilterObject=1 Then
             Response.Write "标题过滤" 
          ElseIf FilterObject=2 Then 
             Response.Write "正文过滤" 
          Else
             Response.Write "<font color=red>没有选择！</font>" 
          End If
          %>           
          </TD>
          <TD class="tdbg" align="center">
          <%
          If FilterType=1 Then
             Response.Write "简单替换"
          ElseIf FilterType=2 Then
             Response.Write "高级过滤"
          Else
             Response.Write "<font color=red>没有选择！</font>"
          End If
          %>
          </TD>
          <TD class="tdbg" align="center">
          <%If Flag=True Then
               Response.Write "启用"
            Else
               Response.Write "<font color=red>禁用</font>"
            End If
          %>&nbsp;
          <%If PublicTf=False Then
               Response.Write "私有"
            Else
               Response.Write "<font color=red>公有</font>"
            End If
          %>
          </TD>
          <TD class="tdbg" align="center">
            <%If Flag=True Then 
                Response.Write "<a href=cj_Settings.asp?Action=SetFlag&FlagName=a0&FilterID=" & FilterID & ">禁用</a>&nbsp;" & vbCrLf
            Else
                Response.Write "<a href=cj_Settings.asp?Action=SetFlag&FlagName=a1&FilterID=" & FilterID & ">启用</a>&nbsp;" & vbCrLf
            End If

            If PublicTf=True Then 
                Response.Write "&nbsp;&nbsp;<a href=cj_Settings.asp?Action=SetFlag&FlagName=b0&FilterID=" & FilterID & ">私有</a>&nbsp;" & vbCrLf
            Else
                Response.Write "&nbsp;&nbsp;<a href=cj_Settings.asp?Action=SetFlag&FlagName=b1&FilterID=" & FilterID & ">公有</a>&nbsp;" & vbCrLf
            End If			
			%>
           &nbsp;
             <a href="cj_modify.asp?FilterID=<%=RsItem("FilterID")%>">修改</a>
             &nbsp;<a href="cj_Settings.asp?Action=SetFlag&FlagName=Del&FilterID=<%=RsItem("FilterID")%>" onclick='return confirm("确定要删除此记录吗？");'>删除</a>
          </td>
          </TD>
        </TR>
<%    iItem=iItem+1
      If iItem>=MaxPerPage Then Exit Do
      RsItem.MoveNext
   Loop 
%>
    <tr class="tdbg"> 
      <td colspan=9 height="30">  
        <input name="Action" type="hidden"  value="SetFlag">
        &nbsp;<input name="chkAll" type="checkbox" id="chkAll" onclick=CheckAll(this.form) value="checkbox" >全选
      </td>
    </tr>
    <tr class="tdbg"> 
      <td colspan=9 height="30" align="center">  
        <input type="submit" value="删&nbsp;&nbsp;除" name="FlagName" style="cursor: hand;background-color: #cccccc;" onclick='return confirm("确定要执行删除操作吗？该操作不可恢复！");'>&nbsp;&nbsp;
        <input type="submit" value="启&nbsp;&nbsp;用" name="FlagName" style="cursor: hand;background-color: #cccccc;" onclick='return confirm("确定要启用所选择的项目吗？");'>&nbsp;&nbsp;
        <input type="submit" value="禁&nbsp;&nbsp;用" name="FlagName" style="cursor: hand;background-color: #cccccc;" onclick='return confirm("确定要禁用所选择的项目吗？");'>&nbsp;&nbsp;
        <input type="submit" value="公&nbsp;&nbsp;有" name="FlagName" style="cursor: hand;background-color: #cccccc;" onclick='return confirm("确定要将所选择的项目设为公有吗？");'>&nbsp;&nbsp;
        <input type="submit" value="私&nbsp;&nbsp;有" name="FlagName" style="cursor: hand;background-color: #cccccc;" onclick='return confirm("确定要将所选择的项目设为私有吗？");'>&nbsp;&nbsp;
      </td>
    </tr>
<%Else%>
<tr class="tdbg">
        <td colspan='9' class="tdbg" align="center"><br>系统中暂无过滤记录！</td>
</tr> 
<%End If
 RsItem.Close
Set RsItem=Nothing
%>
</form>  
</table>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" class="border" >
    <tr> 
      <td height="22" colspan="2" class="tdbg">
<%
Response.Write ShowPage("cj_Settings.asp",ItemNum,MaxPerPage,True,True," 个记录")
%>

      </td>
    </tr>
</table>
</body>
</html>
<%end sub%>
<%
Sub  SetFlag
   FilterID=Trim(Request("FilterID"))
   FlagName=Trim(Request("FlagName"))
   If FilterID<>"" Then
      FilterID=Replace(FilterID," ","")
   Else
      FoundErr=True
      ErrMsg=ErrMsg & "<br><li>请选择要执行操作的记录！</li>"
   End if
   If FoundErr<>True Then
      Select Case FlagName
      Case "Del","删&nbsp;&nbsp;除"
         SqlItem="Delete from [DNJY_Article_Filters] Where FilterID In(" & FilterID & ")"
      Case "b0"
         SqlItem="Update [DNJY_Article_Filters] set PublicTf="&DN_False&" Where FilterID In(" & FilterID & ")"
      Case "b1"
         SqlItem="Update [DNJY_Article_Filters] set PublicTf="&DN_True&" Where FilterID In(" & FilterID & ")"
      Case "公&nbsp;&nbsp;有"
         SqlItem="Update [DNJY_Article_Filters] set PublicTf="&DN_True&" Where FilterID In(" & FilterID & ")"
      Case "私&nbsp;&nbsp;有"
         SqlItem="Update [DNJY_Article_Filters] set PublicTf="&DN_False&" Where FilterID In(" & FilterID & ")"
      Case "a0"
         SqlItem="Update [DNJY_Article_Filters] set Flag="&DN_False&" Where FilterID In(" & FilterID & ")"
      Case "a1"
         SqlItem="Update [DNJY_Article_Filters] set Flag="&DN_True&" Where FilterID In(" & FilterID & ")"
      Case "启&nbsp;&nbsp;用"
         SqlItem="Update [DNJY_Article_Filters] set Flag="&DN_True&" Where FilterID In(" & FilterID & ")"
      Case "禁&nbsp;&nbsp;用"
         SqlItem="Update [DNJY_Article_Filters] set Flag="&DN_False&" Where FilterID In(" & FilterID & ")"
      End Select
      Conn.Execute(SqlItem)
   End If
End  Sub
%>
