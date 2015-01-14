<!--#include file="../conn.asp"-->
<!--#include file="../config.asp"-->
<!--#include file=cookies.asp-->
<!--#include file="inc/Function.asp"-->
<%
'=====================================
'天天供求信息网站管理系统
'天天网络科技工作室出品
'客服中心http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>
<%call checkmanage("10")%>
<html>
<head>
<title>信息采集系统规则复制</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="../css/manage.css" type="text/css">
</head>
<body leftmargin="2" topmargin="0" marginwidth="0" marginheight="0">
<%
Dim Action,ItemID,Arr_Item,Arr_i,ItemName,Arr_Filter,FoundErr
Action=Trim(Request("Action"))
ItemID=Trim(Request("ItemID"))
FoundErr=False

If Action<>"Copy" Then
   FoundErr=True
   ErrMsg=ErrMsg & "<br><li>参数错误，请从有效链接进入</li>"
End if   
If ItemID=""  Then
   FoundErr=True
   ErrMsg=ErrMsg & "<br><li>参数错误，请从有效链接进入</li>"
Else
   ItemID=Clng(ItemID)
End If
Call OpenConn
If FoundErr<>True Then
   Set Rs=conn.execute("Select * from DNJY_xx_Item Where ItemID=" & ItemID)
   If Not Rs.Eof And Not Rs.Bof Then
      Arr_Item=Rs.GetRows()
   Else
      Arr_Item=""
   End if
   Set Rs=Nothing
   
   If IsArray(Arr_Item)=True Then
      set rs=server.createobject("adodb.recordset")
      sql="select top 1 * from DNJY_xx_Item" 
      rs.open sql,conn,1,3
      rs.AddNew
         rs(1)=Arr_Item(1,0)&"复件"
         ItemName=Arr_Item(1,0)
         For Arr_i=2 To Ubound(Arr_Item,1)
            rs(Arr_i)=Arr_Item(Arr_i,0)
         Next
         ItemID=rs("ItemID")
      rs.Update
      rs.close
      set rs=Nothing

      If IsArray(Arr_Filter)=True Then
         set rs=server.createobject("adodb.recordset")
         sql="select top 1 * from DNJY_xx_Filters" 
         rs.open sql,conn,1,3
         rs.AddNew
            rs(1)=ItemID
            For Arr_i=2 To Ubound(Arr_Filter,1)
               rs(Arr_i)=Arr_Filter(Arr_i,0)
            Next
         rs.Update
         rs.close
         set rs=Nothing
      End if      
   else
      FoundErr=True
      ErrMsg=ErrMsg & "参数错误，没有找到该项目"
   end if
   
End if

Call Main  
Call CloseDB(conn)'关闭数据库链接
Sub Main%>
<table width="100%" border="0" align="center" cellpadding="3" cellspacing="2" class="admintable">
  <tr>
    <td height="30" class="b1_1" style="text-align:center;padding:50px;"><%=ItemName%> 项目复制完成，新的项目保存为：<%=ItemName%>复件<br><br><a href="collect_start.asp">返回</a></td>
  </tr>
</table>
<%End Sub%>
</body>         
</html>