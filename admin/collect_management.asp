<!--#include file="../conn.asp"-->
<!--#include file="../config.asp"-->
<!--#include file=cookies.asp-->
<!--#include file="inc/Function.asp"-->
<%
'=====================================
'���칩����Ϣ��վ����ϵͳ
'��������Ƽ������ҳ�Ʒ
'�ͷ�����http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>
<%call checkmanage("10")
response.buffer=true
Call OpenConn
Dim SqlItem,RsItem
Dim Action,FoundErr,ErrMsg
Dim ItemID,ItemName,city_oneid,city_twoid,city_threeid,city_one,city_two,city_three,type_oneid,type_twoid,type_threeid,type_one,type_two,type_three,ListStr,ListPaingType,ListPaingStr2,ListPaingID1,ListPaingID2,ListPaingStr3,Flag
Dim ListUrl,ItemCollecDate
Dim CurrentPage,AllPage,iItem,ItemNum
Const MaxPerPage=10
Action=Request("Action")
If Action="Del" Then
   Call Del()
End If
If FoundErr=True Then
   Call WriteErrMsg(ErrMsg)
Else
   Call Main()
End If
'�ر����ݿ�����
conn.close:Set conn=nothing
%>
<%Sub Main%>
<html>
<head>
<title>�ɼ�ϵͳ</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="../css/manage.css" type="text/css">
<BODY background="img/back.gif">
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
            if (e.Name != "chkAll"&&e.disabled!=true)
                e.checked = thisform.chkAll.checked;
        }
    }
</script>
</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" class="border">
  <tr class='topbg'> 
    <td height="22" colspan="2" align="center" ><strong>��Ϣ�ɼ�ϵͳ��Ŀ����</strong></td>
  </tr>
</table>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" class="border">
  <tr class="tdbg"> 
    <td width="65" height="30"><strong>��������</strong></td>
    <td height="30"><a href=collect_management.asp>������ҳ</a> | <a href="collect_add_a.asp">�������Ŀ</a></td>
  </tr>
</table>
<br>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" class="border" >
    <tr> 
      <td height="22" colspan="2" class="title"> <div align="center"><strong>��  Ŀ  ��  ��</strong></div></td>
    </tr>
</table>
<table class="border" border="0" cellspacing="1" width="100%" cellpadding="0">
  <form name="myform" method="POST" action="collect_management.asp">
    <tr class="tdbg" style="padding: 0px 2px;">
      <td width="38" height="22" align="center" class=ButtonList>ѡ��</td>
      <td width="141" align="center" class=ButtonList>��Ŀ����</td>
      <td width="130" align="center" class=ButtonList>�ɼ���ַ</td>
      <td width="120" height="22" align="center" class=ButtonList>��������</td>
      <td width="120" height="22" align="center" class=ButtonList>������Ŀ</td>      
      <td width="43" align="center" class=ButtonList>״̬</td>
      <td width="157" height="22" align="center" class=ButtonList>�ϴβɼ�</td>
      <td width="148" height="22" align="center" class=ButtonList>����</td>
    </tr>
    <%            
If Request("page")<>"" then
    CurrentPage=Cint(Request("Page"))
Else
    CurrentPage=1
End if                 
Set RsItem=server.createobject("adodb.recordset")         
SqlItem="select ItemID,ItemName,WebName,ListStr,ListPaingType,ListPaingStr2,ListPaingID1,ListPaingID2,ListPaingStr3,city_oneid,city_twoid,city_threeid,city_one,city_two,city_three,type_oneid,type_twoid,type_threeid,type_one,type_two,type_three,Flag from DNJY_xx_Item order by ItemID "&DN_OrderType&""
RsItem.open SqlItem,conn,1,1

if Not RsItem.Eof then
   RsItem.PageSize=MaxPerPage
   Allpage=RsItem.PageCount
   If Currentpage>Allpage Then Currentpage=1
   ItemNum=RsItem.RecordCount
   RsItem.MoveFirst
   RsItem.AbsolutePage=CurrentPage
   iItem=0
   Do While Not RsItem.Eof
      ItemID=RsItem("ItemID")
      ItemName=RsItem("ItemName")
      WebName=RsItem("WebName")
	  city_oneid=rsitem("city_oneid")
	  city_twoid=rsitem("city_twoid")
	  city_threeid=rsitem("city_threeid")
	  city_one=rsitem("city_one")
	  city_two=rsitem("city_two")
	  city_three=rsitem("city_three")
	  type_oneid=rsitem("type_oneid")
	  type_twoid=rsitem("type_twoid")
	  type_threeid=rsitem("type_threeid")
	  type_one=rsitem("type_one")
	  type_two=rsitem("type_two")
	  type_three=rsitem("type_three")
      ListStr=RsItem("ListStr")
      ListPaingType=RsItem("ListPaingType")
      ListPaingStr2=RsItem("ListPaingStr2")
      ListPaingID1=RsItem("ListPaingID1")
      ListPaingID2=RsItem("ListPaingID2")
      ListPaingStr3=RsItem("ListPaingStr3")
      Flag=RsItem("Flag")
      If  ListPaingType=0  Or ListPaingType=1  Then
            ListUrl=ListStr
      ElseIf  ListPaingType=2  Then
            ListUrl=Replace(ListPaingStr2,"{$ID}",CStr(ListPaingID1))
      ElseIf  ListPaingType=3  Then
            If  Instr(ListPaingStr3,"|")>0  Then
            ListUrl=Left(ListPaingStr3,Instr(ListPaingStr3,"|")-1)
         Else
               ListUrl=ListPaingStr3
         End  If
      End  If

%>
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'" style="padding: 0px 2px;" align="center" height="22">
      <td width="38"><input type="checkbox" value="<%=ItemID%>" name="ItemID" onClick="unselectall(this.form)"></td>
      <td width="141"><%=ItemName%></td>
      <td width="130"><a href="<%=ListUrl%>" target="_bank"><%=WebName%></a></td>
      <td width="120"><%=city_one%><%If city_two<>"" Then Response.Write " / "%><%=city_two%><%If city_three<>"" Then Response.Write " / "%><%=city_three%></td>
      <td width="120"><%=type_one%><%If type_two<>"" Then Response.Write " / "%><%=type_two%><%If type_three<>"" Then Response.Write " / "%><%=type_three%></td>
      
      <td width="43"> <b>
        <%If Flag=True then
                    Response.write "��"
          Else
                 Response.write "<font color=red>��</font>"
          End If        %>
      </b> </td>
      <td width="157">
        <%       Set Rs=conn.execute("select Top 1 CollecDate From DNJY_xx_Histroly Where ItemID=" & ItemID & " Order by HistrolyID "&DN_OrderType&"")
       If Not Rs.Eof Then
          ItemCollecDate=rs("CollecDate")
       Else
          ItemCollecDate=""
       End if
       Set Rs=Nothing
       if ItemCollecDate<>"" then
          Response.Write ItemCollecDate
       Else
          Response.Write "���޼�¼"
       End If
       %>      </td>
      <td width="148"><a href=collect_editor_a.asp?ItemID=<%=ItemID%>>�༭</a> <a href=collect_attribute.asp?ItemID=<%=ItemID%>>����</a> <a href=collect_editor_e.asp?ItemID=<%=ItemID%>>����</a> <a href="collect_copy.asp?Action=Copy&ItemID=<%=ItemID%>">����</a> <a href=collect_management.asp?Action=Del&ItemID=<%=ItemID%> onClick='return confirm("ȷ��Ҫɾ������Ŀ����������ѡ���⽫ɾ������Ŀ����Ŀ��Ϣ����ʷ��¼��������Ϣ 3 ����Ŀ�������ݡ�");'>ɾ��</a></td>
    </tr>
    <%iItem=iItem+1
      If iItem>=MaxPerPage Then  Exit  Do
      RsItem.MoveNext
   Loop%>
    <tr class="tdbg">
      <td colspan=9 height="30">
        <input name="Action" type="hidden"  value="Del">&nbsp;
        <input name="chkAll" type="checkbox" id="chkAll" onClick=CheckAll(this.form) value="checkbox" >
        ȫѡ </td>
    </tr>
    <tr class="tdbg">
      <td colspan=9 height="30" align=center> &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
          <input type="submit" value=" ɾ&nbsp;&nbsp;�� " name="Del" onClick='return confirm("ȷ��Ҫɾ��ѡ�е���Ŀ����������ѡ���⽫ɾ������Ŀ����Ŀ��Ϣ����ʷ��¼��������Ϣ 3 ����Ŀ�������ݡ�");' style="cursor: hand;background-color: #cccccc;">
  &nbsp;&nbsp;
          <input type="submit" value="������м�¼" name="Del" onClick='return confirm("�����Ҫȷ��Ҫ���������Ŀ���⽫���׸�ʽ���ɼ����ݿ��������Ϣ�������ȱ�����ѡ�񣡣���");' style="cursor: hand;background-color: #cccccc;">
  &nbsp;&nbsp; </td>
    </tr>
    <%Else%>
    <tr class="tdbg">
      <td colspan='9' class="tdbg" align="center"><br>
        ϵͳ�����޲ɼ���Ŀ��</td>
    </tr>
    <%End If
RsItem.Close
Set  RsItem=Nothing
%>
  </form>
</table>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" class="border" >
    <tr> 
      <td height="22" colspan="2" class="tdbg">
<%
Response.Write ShowPage("collect_management.asp",ItemNum,MaxPerPage,True,True," ����Ŀ")
%>

      </td>
    </tr>
</table>
</body>
</html>
<%end sub%>
<%Sub Del
ItemID=Trim(Request("ItemID"))
If Request("Del")="������м�¼" Then
   conn.Execute("Delete From DNJY_xx_Item")
   conn.Execute("Delete From DNJY_xx_Filters")
   conn.Execute("Delete From DNJY_xx_Histroly")
Else
   If ItemID="" Then
      FoundErr=True
      ErrMsg=ErrMsg & "<br><li>��ѡ��Ҫɾ������Ŀ��</li>"
   Else
      ItemID=Replace(ItemID," ","")   
      conn.Execute("Delete From [DNJY_xx_Item] Where ItemID In(" & ItemID & ")")
      conn.Execute("Delete From [DNJY_xx_Filters] Where ItemID In(" & ItemID & ")")
      conn.Execute("Delete From [DNJY_xx_Histroly] Where ItemID In(" & ItemID & ")")
   End If
End  If
End Sub
%>
