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
   ErrMsg=ErrMsg & "<br><li>����������ĿID����Ϊ�գ�</li>"
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

'�ر����ݿ�����
conn.close:SEt conn=nothing


Sub Main
%>
<html>
<head>
<title>�ɼ�ϵͳ</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="../css/manage.css" type="text/css">
</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" class="border">
  <tr> 
    <td height="22" colspan="2" align="center" class="topbg"><strong>���Ųɼ�ϵͳģ�����</td>
  </tr>
  <tr class="tdbg"> 
    <td width="65" height="30"><strong>��������</strong></td>
    <td height="30">��Ŀ�༭ >> <a href="cj_editor_a.asp?ItemID=<%=ItemID%>">��������</a> >> <a href="cj_editor_b.asp?ItemID=<%=ItemID%>">�б�����</a> >> <a href="cj_editor_c.asp?ItemID=<%=ItemID%>">��������</a> >> <a href="cj_editor_d.asp?ItemID=<%=ItemID%>">��������</a> >>  
	<a href="cj_editor_e.asp?ItemID=<%=ItemID%>">��������</a> >> <a href="cj_attribute.asp?ItemID=<%=ItemID%>"><font color=red>��������</font></a> >> ���</td>         
  </tr>         
</table>
<br>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" class="border">
    <tr> 
      <td height="22" colspan="2" class="title"> <div align="center"><strong>�� �� �� Ŀ--�� �� �� ��</strong></div></td>
    </tr>
</table>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" class="border,tdbg">
<form method="post" action="cj_success.asp" name="myform">
    <tr class="tdbg">
      <td height="30" width="20%" align="right"><b>�������ʼֵ��</b></td>
      <td><input name="Hits" type="text" id="Hits" value="<%=Hits%>" size="10" maxlength="10">
        <font color='#0000FF'>�⹦�����ṩ������Ա�����õġ�����������Ҫ��ѽ��^_^</font>      </td>
    </tr>
    <tr class="tdbg">
      <td height="30" width="20%" align="right"><b>�������ʣ�</b></td>
      <td> 
        <input name="IncludePicYn" type="checkbox" value="yes" <%If IncludePicYn=-1 Then Response.write "checked"%> style="border: 0px;background-color: #eeeeee;">
        ����ͼƬ
        <input name="DefaultPicYn" type="checkbox" value="yes" <%If DefaultPicYn=-1 Then Response.write "checked"%> style="border: 0px;background-color: #eeeeee;">
        �Ƽ�����
        <input name="OnTop" type="checkbox" value="yes" <%If OnTop=-1 Then Response.write "checked"%> style="border: 0px;background-color: #eeeeee;">
        �̶�����
		<input name="Hot" type="checkbox" value="yes" <%If Hot=-1 Then Response.write "checked"%> style="border: 0px;background-color: #eeeeee;">
        �ȵ�����</td>
    </tr>  

    <tr class="tdbg">
      <td height="30" width="20%" align="right"><b>��ǩ���ˣ�</b></td>
      <td>
        <input name="Script_Iframe" type="checkbox" value="yes" <%If Script_Iframe=-1 Then response.write "checked"%> style="border: 0px;background-color: #eeeeee;">
        Iframe
        <input name="Script_Object" type="checkbox" value="yes" <%If Script_Object=-1 Then response.write "checked"%> onclick='return confirm("ȷ��Ҫѡ��ñ�����⽫ɾ�������е�����Object��ǣ���������¸������е�����Flash������ɾ����");' style="border: 0px;background-color: #eeeeee;">
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
        <input name="Script_Html" type="checkbox" value="yes" <%If Script_Html=-1 Then response.write "checked"%> onclick='return confirm("ȷ��Ҫѡ��ñ�����⽫ɾ�������е�����Html��ǣ���������¸����µĿ��Ķ��Խ��ͣ�");' style="border: 0px;background-color: #eeeeee;">
        Html&nbsp;
        <input name="Script_Td" type="checkbox"  value="yes" <%If Script_Td=-1 Then response.write "checked"%> style="border: 0px;background-color: #eeeeee;">
        Td        </td>
    </tr>
    <tr class="tdbg">
      <td height="30" width="20%" align="right"> <b>�б���ȣ�</b></td>
      <td>
        <input name="CollecListNum" type="text" id="CollecListNum" value="<%=CollecListNum%>" size="10" maxlength="10">&nbsp;&nbsp;&nbsp;
		<font color='#0000FF'>0Ϊ���е��б�</font></td>
    </tr>
    <tr class="tdbg">
      <td height="30" width="20%" align="right"> <b>����������</b></td>
      <td>
        <input name="CollecNewsNum" type="text" id="CollecNewsNum" value="<%=CollecNewsNum%>" size="10" maxlength="10">&nbsp;&nbsp;&nbsp; 
		<font color='#0000FF'>0Ϊ���е�����<span lang="en-us">(</span>ÿһ�б��������������<span lang="en-us">)</span></font></td>
    </tr>
    <tr class="tdbg">
      <td height="30" width="20%" align="right"> <b>�ɼ�ѡ�</b></td>
      <td><input name="Passed" type="checkbox" value="yes" <%if Passed=-1 then response.write "checked"%> style="border: 0px;background-color: #eeeeee;">
        �Զ����		
        <input name="SaveFiles" type="checkbox" value="yes" <%if SaveFiles=-1 then response.write "checked"%> <%If IsObjInstalled("Scripting.FileSystemObject")=False Then Response.Write "disabled"%> style="border: 0px;background-color: #eeeeee;">
        ����ͼƬ
		<input name="cj_watermark" type="checkbox" value="yes" <%if cj_watermark=-1 then response.write "checked"%> <%If IsObjInstalled("Scripting.FileSystemObject")=False Then Response.Write "disabled"%> style="border: 0px;background-color: #eeeeee;">
        ͼƬˮӡ
        <input name="CollecOrder" type="checkbox" value="yes" <%if CollecOrder=-1 then response.write "checked"%> style="border: 0px;background-color: #eeeeee;">
        ����ɼ�
        <input name="LinkUrlYn" type="checkbox" value="yes" <%if LinkUrlYn=-1 then response.write "checked"%> style="border: 0px;background-color: #eeeeee;">
        URL��ת</td>
    </tr>
    <tr class="tdbg">
    <td height="30" width="20%" align="right"></td>
    <td align-"center"><center>
       <input type="hidden" value="<%=ItemID%>" name="ItemID">        
       <input type="submit" value=" ��&nbsp;&nbsp;�� " name="submit" style="cursor: hand;background-color: #cccccc;">  </center>       </td> 
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
      ErrMsg=ErrMsg & "<br><li>���������Ҳ�������Ŀ</li>"
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
