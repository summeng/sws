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
response.buffer=true	
Call OpenConn
Dim SqlItem,RsItem
Dim Action,FoundErr,ErrMsg
Dim ItemID,ItemName,city_oneid,city_twoid,city_threeid,city_one,city_two,city_three,c_oneid,c_twoid,c_threeid,c_one,c_two,c_three,ListStr,ListPaingType,ListPaingStr2,ListPaingID1,ListPaingID2,ListPaingStr3,Flag,ItemCollecDate
Dim ListUrl
Dim CurrentPage,AllPage,iItem,ItemNum
Const MaxPerPage=10


Call Main()
'关闭数据库链接
conn.close:Set conn=nothing
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
        for (var i=0;i<thisform.elements.length-6;i++){
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
    <td height="22" colspan="2" align="center" ><strong>新闻采集系统项目管理</strong></td>
  </tr>
</table>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" class="border">
  <tr class="tdbg"> 
    <td width="65" height="30"><strong>管理导航：</strong></td>
    <td height="30"><a href=cj_start.asp>管理首页</a> | <a href="cj_add_a.asp">添加新项目</a></td>
  </tr>
</table>
<br>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" class="border" >
    <tr> 
      <td height="22" colspan="2" class="title"> <div align="center"><b>新&nbsp;闻&nbsp;采&nbsp;集</b></div></td>
    </tr>
</table>
<table class="border" border="0" cellspacing="1" width="100%" cellpadding="0">
<form name="myform" method="POST" action="cj_tion.asp">
        <TR>
          <TD class=ButtonList width="5%" height=22 align=center>选择</TD>
          <TD class=ButtonList width="17%" height=22  align=center>项目名称</TD>
          <TD class=ButtonList width="15%" height=22  align=center>采集来源</TD>
          <TD class=ButtonList width="15%" height=22  align=center>所属地区</TD>
            <TD class=ButtonList width="15%" height=22  align=center>所属栏目</TD>
            <TD class=ButtonList width="15%" height=22  align=center>所属专题</TD>
            <TD class=ButtonList height=22  align=center>上次采集</TD>
           </TR>
<%
If Request("page")<>"" then
    CurrentPage=Cint(Request("Page"))
Else
    CurrentPage=1
End if      
set RsItem=server.createobject("adodb.recordset")         
SqlItem="select * from DNJY_Article_Item where Flag="&DN_True&" order by ItemID "&DN_OrderType&""         
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
       
	   
	  city_oneid=rsitem("city_oneid")
	  city_twoid=rsitem("city_twoid")
	  city_threeid=rsitem("city_threeid")
	  city_one=rsitem("city_one")
	  city_two=rsitem("city_two")
	  city_three=rsitem("city_three")
    c_oneid=rsitem("c_oneid")
    c_one=rsitem("c_one")
    c_twoid=rsitem("c_twoid")
    c_two=rsitem("c_two")
    c_threeid=rsitem("c_threeid")
    c_three=rsitem("c_three")
      ItemID=RsItem("ItemID")
      ItemName=RsItem("ItemName")
      WebName=RsItem("WebName")
            ListPaingType=RsItem("ListPaingType")
            If  ListPaingType=0  Or ListPaingType=1  Then
            ListUrl=RsItem("ListStr")
      ElseIf  ListPaingType=2  Then
            ListUrl=Replace(RsItem("ListPaingStr2"),"{$ID}",CStr(RsItem("ListPaingID1")))
      ElseIf  ListPaingType=3  Then
            If  Instr(RsItem("ListPaingStr3"),"|")>0  Then
            ListUrl=Left(RsItem("ListPaingStr3"),Instr(RsItem("ListPaingStr3"),"|")-1)
         Else
               ListUrl=RsItem("ListPaingStr3")
         End  If
      Else
            ListUrl="异常"
      End  If
   %>
        <TR class="tdbg">
          <TD class="tdbg" style="WIDTH: 6%" align=center>
            <input type="checkbox" name="ItemID" value="<%=ItemID%>" onClick="unselectall(this.form)">
          </TD>
          <TD class="tdbg" align="center"><%=ItemName%>
          </TD>
          <TD class="tdbg"  align="center"><a href="<%=ListUrl%>" title="点击访问" target=_blank><%=WebName%></a>
          </TD>
          <TD class="tdbg" align="center"><%=city_one%><%If city_two<>"" Then Response.Write " / "%><%=city_two%><%If city_three<>"" Then Response.Write " / "%><%=city_three%></TD>
          <TD class="tdbg" align="center"><%=c_one%><%If c_two<>"" Then Response.Write " / "%><%=c_two%><%If c_three<>"" Then Response.Write " / "%><%=c_three%></TD>
          <TD class="tdbg" align="center">&nbsp;</TD>
          <TD class="tdbg" align="center">
          <%
           set rs=conn.execute("select Top 1 CollecDate From DNJY_Article_Histroly Where ItemID=" & ItemID & " Order by HistrolyID "&DN_OrderType&"")
           If not rs.eof Then
              ItemCollecDate=rs("CollecDate")
           else
              ItemCollecDate=""
           end if
           Set rs=nothing
           if ItemCollecDate<>"" then
              Response.Write ItemCollecDate
           Else
              Response.Write "尚无记录"
           End If
          %>
          </TD>
        </TR>
<%    iItem=iItem+1
      If iItem>=MaxPerPage Then Exit Do
      RsItem.MoveNext
   Loop 
%>
    <tr class="tdbg"> 
      <td colspan=9 height="30">  
        <input name="Action" type="hidden"  value="Start">
        &nbsp;<input name="chkAll" type="checkbox" id="chkAll" onclick=CheckAll(this.form) value="checkbox" >全选
      </td>
    </tr>
    <tr class="tdbg"> 
      <td colspan=9 height="30">  
        &nbsp;&nbsp;&nbsp;采集模式：<input name="CollecType" type="radio" id="CollecType" value="1" checked onClick="javascript:document.myform.Content_View.checked=false">快速模式&nbsp;&nbsp;
        <input name="CollecType" type="radio" id="CollecType" value="0" onClick="javascript:document.myform.Content_View.checked=true" disabled>稳定模式&nbsp;&nbsp;
        <input name="CollecType" type="radio" id="CollecType" value="2" disabled>筛选模式&nbsp;&nbsp;
        <input name="CollecTest" type="checkbox" id="CollecTest" value="yes" onClick="javascript:document.myform.Content_View.checked=true">采集测试&nbsp;&nbsp;<input name="Content_View" type="checkbox" id="Content_View" value="yes">正文预览&nbsp;&nbsp;<input type="submit" value="开始采集" name="StartMe" style="cursor: hand;background-color: #cccccc;">&nbsp;&nbsp;
      </td>
    </tr>
<%Else%>
<tr class="tdbg">
        <td colspan='9' class="tdbg" align="center"><br>系统中暂无可用采集项目！</td>
</tr>
<%End If
 RsItem.Close
Set RsItem=Nothing
%>
</form>  
</TABLE>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" class="border" >
    <tr> 
      <td height="22" colspan="2" class="tdbg">
<%
Response.Write ShowPage("cj_start.asp",ItemNum,MaxPerPage,True,True," 个项目")
%>

      </td>
    </tr>
</table>
</body>
</html>
<%end sub%>
