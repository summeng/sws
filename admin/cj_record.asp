<!--#include file="../conn.asp"-->
<!--#include file="../config.asp"-->
<!--#include file=cookies.asp-->
<!--#include file="inc/Article_Function.asp"-->
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
Dim HistrolyID,ItemID,ItemName,ArticleID,NewsUrl,Result,city_oneid,city_twoid,city_threeid,city_one,city_two,city_three,c_oneid,c_twoid,c_threeid,c_one,c_two,c_three
Dim ListUrl,ItemCollecDate
Dim CurrentPage,AllPage,iItem,ItemNum
Const MaxPerPage=50
Action=Request("Action")
If Action="Del" Then
   Call DelHistroly()
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
<link rel="stylesheet" type="text/css" href="../css/manage.css">
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
    <td height="22" colspan="2" align="center" ><strong>���Ųɼ�ϵͳ��ʷ��¼����</strong></td>
  </tr>
</table>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" class="border">
  <tr class="tdbg"> 
    <td width="65" height="30"><strong>��������</strong></td>
    <td height="30"><a href=cj_record.asp>������ҳ</a> </td>
  </tr>
</table>
<br>
<table width="100%" border="0" align="center" cellpadding="3" cellspacing="2" class="admintable">
  <tr>
    <td height="30" class="b1_1"><a href="cj_record.asp">������ҳ</a>&nbsp;&nbsp;|&nbsp;&nbsp;<a href="cj_record.asp?Action=Succeed">�ɹ���¼</a>&nbsp;&nbsp;|&nbsp;&nbsp;<a href="cj_record.asp?Action=Failure">ʧ�ܼ�¼</a></td>
  </tr>
</table>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" class="border" >
    <tr> 
      <td height="22" colspan="2" class="title"> <div align="center"><strong>�� ʷ �� ¼  ��  ��</strong></div></td>
    </tr>
</table>
<table class="border" border="0" cellspacing="1" width="100%" cellpadding="0">
  <form name="myform" method="POST" action="cj_record.asp">
    <tr class="tdbg" style="padding: 0px 2px;">
      <td width="30" height="22" align="center" class=ButtonList>ѡ��</td>
      <td width="100" align="center" class=ButtonList>��¼����</td>
      <td width="300" align="center" class=ButtonList>����</td>
      <td width="170" height="22" align="center" class=ButtonList>��������</td>
      <td width="170" height="22" align="center" class=ButtonList>������Ŀ</td>
      <td width="50" align="center" class=ButtonList>�ɼ�ҳ��</td>
      <td width="30" align="center" class=ButtonList>�� ��</td>     
      <td width="30" height="22" align="center" class=ButtonList>����</td>
    </tr>
    <%            
If Request("page")<>"" then
    CurrentPage=Cint(Request("Page"))
Else
    CurrentPage=1
End if                 
Set RsItem=server.createobject("adodb.recordset")         
SqlItem="select * from DNJY_Article_Histroly"
If Action="Succeed"  Then'�ɹ���¼   
   SqlItem=SqlItem  &  " Where  Result="&DN_True&""
ElseIf Action="Failure"  Then'ʧ�ܼ�¼
   SqlItem=SqlItem  &  " Where  Result="&DN_False&""
ElseIf  Action="search"  Then'������¼
   SqlItem=SqlItem  &  " Where  ItemID = "&request("ItemID")&""

Else'�� �� �� ¼
End  If
SqlItem=SqlItem&" order by HistrolyID"&DN_OrderType&""
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
      HistrolyID=RsItem("HistrolyID")
	  ItemID=RsItem("ItemID")
      Title=RsItem("Title")
	  ArticleID=RsItem("ArticleID")
      NewsUrl=RsItem("NewsUrl")
      Result=RsItem("Result")    
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

%>
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'" style="padding: 0px 2px;"  height="22">
      <td width="30"><input type="checkbox" value="<%=HistrolyID%>" name="HistrolyID" onClick="unselectall(this.form)"></td>
      <td width="100"><%set rs=server.CreateObject("adodb.recordset")
	  sql="select ItemName From DNJY_Article_Item Where ItemID="&ItemID&"" 
	  Rs.open Sql,conn,1,1
       ItemName=rs("ItemName")
		Response.Write ItemName
       Set Rs=Nothing%></td>
      <td width="300"><a href="http://<%=web%>/news_show.asp?id=<%=ArticleID%>" target="_bank"><%=Title%></a></td>
      <td width="170"><%=city_one%><%If city_two<>"" Then Response.Write " / "%><%=city_two%><%If city_three<>"" Then Response.Write " / "%><%=city_three%></td>
      <td width="170"><%=c_one%><%If c_two<>"" Then Response.Write " / "%><%=c_two%><%If c_three<>"" Then Response.Write " / "%><%=c_three%></td>
      <td width="50"><a href="<%=NewsUrl%>" target="_bank">������</a></td>
      <td width="30"> <b>
        <%If Result=True then
                    Response.write "�ɹ�"
          Else
                 Response.write "<font color=red>ʧ��</font>"
          End If        %>
      </b> </td>
      
      <td width="30"> <a href=cj_record.asp?Action=Del&HistrolyID=<%=HistrolyID%> onClick='return confirm("ȷ��Ҫɾ���˼�¼����������ѡ��");'>ɾ��</a></td>
    </tr>
    <%iItem=iItem+1
      If iItem>=MaxPerPage Then  Exit  Do
      RsItem.MoveNext
   Loop%>
    <tr class="tdbg">
      <td colspan=9 height="30">
        <input name="Action" type="hidden"  value="Del">&nbsp;
        <input name="chkAll" type="checkbox" id="chkAll" onClick=CheckAll(this.form) value="checkbox" >
        ȫѡ
        <input type="submit" value="���ѡ�м�¼" name="DelFlag" onClick='return confirm("ȷ��Ҫ�����ѡ��¼��");'>
        <input type="submit" value="���ʧ�ܼ�¼" name="DelFlag" onClick='return confirm("ȷ��Ҫ�������ʧ�ܼ�¼��");'>
      <input type="submit" value="������м�¼" name="DelFlag" onClick='return confirm("ȷ��Ҫ������м�¼��");'>	</td>
    </tr>

    <%Else%>
    <tr class="tdbg">
      <td colspan='9' class="tdbg" align="center"><br>
        ϵͳ�����޲ɼ���¼��</td>
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
Response.Write ShowPage("cj_record.asp",ItemNum,MaxPerPage,True,True," ����¼")
%>

      </td>
    </tr>
</table>
</body>
</html>
<%end sub%>
<%Sub DelHistroly
Dim DelFlag
DelFlag=Trim(Request("DelFlag"))
HistrolyID=Trim(Request("HistrolyID"))
If HistrolyID<>"" Then
   HistrolyID=Replace(HistrolyID," ","")
End If
If DelFlag="���ѡ�м�¼" Then
   If HistrolyID="" Then
      FoundErr=True
      ErrMsg=ErrMsg & "<br><li>��ѡ��Ҫɾ���ļ�¼</li>"
   Else
      HistrolyID=Replace(HistrolyID," ","")
      SqlItem="Delete From [DNJY_Article_Histroly] Where HistrolyID in(" & HistrolyID & ")"
   End If
ElseIf DelFlag="���ʧ�ܼ�¼" Then
   SqlItem="Delete From [DNJY_Article_Histroly] Where Result="&DN_False&""
ElseIf DelFlag="������м�¼" Then
   SqlItem="Delete From [DNJY_Article_Histroly]"
Else
   If HistrolyID="" Then
      FoundErr=True
      ErrMsg=ErrMsg & "<br><li>��ѡ��Ҫɾ���ļ�¼</li>"
   Else
      HistrolyID=Replace(HistrolyID," ","")
      SqlItem="Delete From [DNJY_Article_Histroly] Where HistrolyID In(" & HistrolyID & ")"
   End If
End if

If FoundErr<>True Then
   Conn.Execute(SqlItem)
End If
End Sub
%>
