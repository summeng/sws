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
<%
response.buffer=true	
%>
<%Call OpenConn
dim FoundErr,ErrMsg,Action
Dim SqlItem,RsItem
Dim ItemID,ItemName,WebUrl,ChannelID,strChannelDir,ClassID,SpecialID,city_oneid,city_twoid,city_threeid,type_oneid,type_twoid,type_threeid,city_one,city_two,city_three,type_one,type_two,type_three,ItemDemo,LoginType,LoginUrl,LoginPostUrl,LoginUser,LoginPass,LoginFalse
Dim ListUrl,LsString,LoString,ListPaingType,LPsString,LPoString,ListStr,ListPaingStr1,ListPaingStr2,ListPaingID1,ListPaingID2,ListPaingStr3
Dim tClass,tSpecial
FoundErr=False

Action=Trim(Request("Action"))
ItemID=Trim(Request("ItemID"))


If ItemID="" Then
   FoundErr=True
   ErrMsg=ErrMsg & "<br><li>����������ĿID����Ϊ�գ�</li>"
Else
   ItemID=Clng(ItemID)
End If

If Action="SaveEdit" And FoundErr<>True Then
   Call SaveEdit()
End If

If FoundErr=True Then
   Call WriteErrMsg(ErrMsg)
Else
   Call Main()
End If
'�ر����ݿ�����
conn.close:SEt conn=nothing
%>
<%Sub Main
   SqlItem="Select * from DNJY_xx_Item Where ItemID=" & ItemID
   Set RsItem=server.CreateObject("adodb.recordset")
   RsItem.Open SqlItem,Conn,1,1
   If RsItem.Eof And RsItem.Bof Then
      FoundErr=True
      ErrMsg=ErrMsg & "<br><li>û���ҵ�����Ŀ!</li>"
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
      If ListPaingStr3<>"" Then
         ListPaingStr3=Replace(ListPaingStr3,"|",CHR(13))
      End If
   End If
   RsItem.Close
   Set RsItem=Nothing
   If FoundErr=True Then
      Call WriteErrMsg(ErrMsg)

   Else
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
    <td height="22" colspan="2" align="center" class="topbg"><strong>��&nbsp;&nbsp;��&nbsp;&nbsp;ϵ&nbsp;&nbsp;ͳ&nbsp;&nbsp;��&nbsp;&nbsp;Ŀ&nbsp;&nbsp;��&nbsp;&nbsp;��</td>
  </tr>
  <tr class="tdbg"> 
    <td width="65" height="30"><strong>��������</strong></td>
    <td height="30">��Ŀ�༭ >> <a href="collect_editor_a.asp?ItemID=<%=ItemID%>">��������</a> >> <a href="collect_editor_b.asp?ItemID=<%=ItemID%>"><font color=red>�б�����</font></a> >> <a href="collect_editor_c.asp?ItemID=<%=ItemID%>">��������</a> >> <a href="collect_editor_d.asp?ItemID=<%=ItemID%>">��������</a> >>  
	<a href="collect_editor_e.asp?ItemID=<%=ItemID%>">��������</a> >> <a href="collect_attribute.asp?ItemID=<%=ItemID%>">��������</a> >> ���</td>         
  </tr>         
</table> 
<br>        
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" class="border" >
<form method="post" action="collect_editor_c.asp" name="form1">
    <tr> 
      <td height="22" colspan="2" class="title"> <div align="center"><strong>�� �� �� Ŀ--��  �� �� ��</strong></div></td>
    </tr>
    <tr class="tdbg"> 
      <td width="20%" class="tdbg"><strong>�б�����ҳ�棺</strong></td>
      <td class="tdbg" width="75%">
		<input name="ListStr" type="text" size="58" maxlength="200" value="<%=ListStr%>">&nbsp;&nbsp;�б�ĵ�һҳ 
      </td>
    </tr>
    <tr class="tdbg"> 
      <td width="20%" class="tdbg" ><strong>�б�ʼ��ǣ�</strong></td>
      <td class="tdbg" width="75%">
      <textarea name="LsString" cols="49" rows="7"><%=LsString%></textarea><br>
      </td>
    </tr>
    <tr class="tdbg"> 
      <td width="20%" class="tdbg" ><strong>�б������ǣ�</strong></td>
      <td class="tdbg" width="75%">
      <textarea name="LoString" cols="49" rows="7"><%=LoString%></textarea><br>
      </td>
    </tr>

    <tr>
      <td width="20%" class="tdbg"><strong> �б�������ҳ��</strong></td>
      <td class="tdbg" width="75%">
		<input type="radio" value="0" name="ListPaingType" <%If ListPaingType=0 Then Response.Write "checked"%> onClick="ListPaing1.style.display='none';ListPaing12.style.display='none';ListPaing2.style.display='none';ListPaing3.style.display='none'">��������&nbsp;
		<input type="radio" value="1" name="ListPaingType" <%If ListPaingType=1 Then Response.Write "checked"%> onClick="ListPaing1.style.display='';ListPaing12.style.display='';ListPaing2.style.display='none';ListPaing3.style.display='none'">���ñ�ǩ&nbsp;
		<input type="radio" value="2" name="ListPaingType" <%If ListPaingType=2 Then Response.Write "checked"%> onClick="ListPaing1.style.display='none';ListPaing12.style.display='none';ListPaing2.style.display='';ListPaing3.style.display='none'">��������&nbsp;
		<input type="radio" value="3" name="ListPaingType" <%If ListPaingType=3 Then Response.Write "checked"%> onClick="ListPaing1.style.display='none';ListPaing12.style.display='none';ListPaing2.style.display='none';ListPaing3.style.display=''">�ֶ����
	   </td>
    </tr>
	<tr class="tdbg" id="ListPaing1" style="display:'<%If ListPaingType<>1 Then Response.Write "none"%>'">
      <td width="20%" class="tdbg"><strong><font color=blue>��ҳ��ʼ��ǣ�</font></strong><p>��</p><p>��</p>
      <strong><font color=blue>��ҳ������ǣ�</font></strong>
      </td>
      <td class="tdbg" width="75%">
		<textarea name="LPsString" cols="49" rows="7"><%=LPsString%></textarea><br>
		<textarea name="LPoString" cols="49" rows="7"><%=LPoString%></textarea>
	   </td>
    </tr>
	<tr class="tdbg" id="ListPaing12" style="display:'<%If ListPaingType<>1 Then Response.Write "none"%>'">
      <td width="20%" class="tdbg"><strong><font color=blue>������ҳ�ض���</font></strong>
      </td>
      <td class="tdbg" width="75%">
		<input name="ListPaingStr1" type="text" size="58" maxlength="200" value="<%=ListPaingStr1%>">
	   </td>
    </tr>
    <tr class="tdbg" id="ListPaing2" style="display:'<%If ListPaingType<>2 Then Response.Write "none"%>'"> 
      <td width="20%" class="tdbg"><strong><font color=blue>�������ɣ�</font></strong></td>
      <td class="tdbg" width="75%">
        ԭ�ַ�����<br>
		<input name="ListPaingStr2" type="text" size="58" maxlength="200" value="<%=ListPaingStr2%>"><br>
                ��ʽ��http://www.scuta.net/list.asp?page={$ID}<br><br>
	    ���ɷ�Χ��<br>
		<input name="ListPaingID1" type="text" size="8" maxlength="200" value="<%=ListPaingID1%>"><span lang="en-us"> To </span><input name="ListPaingID2" type="text" size="8" maxlength="200" value="<%=ListPaingID2%>"><br>
               ��ʽ��ֻ�������֣���������߽���
      </td>
    </tr>
    <tr class="tdbg" id="ListPaing3" style="display:'<%If ListPaingType<>3 Then Response.Write "none"%>'"> 
      <td width="20%" class="tdbg"><strong><font color=blue>�ֶ���ӣ�</font></strong>
      </td>
      <td class="tdbg" width="75%">
      <textarea name="ListPaingStr3" cols="49" rows="7"><%=ListPaingStr3%></textarea><br>
      ��ʽ������һ����ַ�󰴻س�����������һ����
      </td>
    </tr>

    <tr class="tdbg"> 
      <td colspan="2" align="center" class="tdbg">
        <input name="ItemID" type="hidden" id="ItemID" value="<%=ItemID%>">
        <input name="Action" type="hidden" id="Action" value="SaveEdit">
        <input  type="submit" name="Submit" value="��&nbsp;һ&nbsp;��" style="cursor: hand;background-color: #cccccc;"></td>
    </tr>
</form>
</table>    
</body>         
</html>
<%End If%>
<%End Sub%>
<%
Sub SaveEdit
   ItemName=Trim(Request.Form("ItemName"))
   WebName=Trim(Request.Form("WebName"))
   WebUrl=Trim(Request.Form("WebUrl"))
city_oneid=strint(Trim(Request.Form("city_one")))
city_twoid=strint(Trim(Request.Form("city_two")))
city_threeid=strint(Trim(Request.Form("city_three")))

type_oneid=strint(Trim(Request.Form("type_one")))
type_twoid=strint(Trim(Request.Form("type_two")))
type_threeid=strint(Trim(Request.Form("type_three")))

   LoginType=Trim(Request.Form("LoginType"))
   LoginUrl=Trim(Request.Form("LoginUrl"))
   LoginPostUrl=Trim(Request.Form("LoginPostUrl"))
   LoginUser=Trim(Request.Form("LoginUser"))
   LoginPass=Trim(Request.Form("LoginPass"))
   LoginFalse=Trim(Request.Form("LoginFalse"))
   ItemDemo=Request.Form("ItemDemo")
      If ItemName="" Then
         FoundErr=True
         ErrMsg=ErrMsg & "<br><li>��Ŀ���Ʋ���Ϊ��</li>"
      End If
      If WebName="" Then
         FoundErr=True
         ErrMsg=ErrMsg & "<br><li>��վ���Ʋ���Ϊ��</li>"
      End If
      If WebUrl="" Then
         FoundErr=True
         ErrMsg=ErrMsg & "<br><li>��վ��ַ����Ϊ��</li>" 
      End If
	  
	  
If city_oneid=0  Then
   FoundErr=True
   ErrMsg=ErrMsg & "<br><li>δָ������</li>"
Else
   city_one=conn.execute("select city From DNJY_city Where id="&city_oneid&" and twoid=0")(0)
   if city_twoid>0 Then city_two=conn.execute("select city From DNJY_city Where id="&city_oneid&" and twoid="&city_twoid&" and threeid=0")(0)
   if city_twoid>0 and city_threeid>0 Then city_three=conn.execute("select city From DNJY_city Where id="&city_oneid&" and twoid="&city_twoid&" and threeid="&city_threeid)(0)

End if

If type_oneid=0  Then
   FoundErr=True
   ErrMsg=ErrMsg & "<br><li>δָ����Ŀ</li>"
Else
   type_one=conn.execute("select [name] From DNJY_type Where id="  & type_oneid&" and twoid=0")(0)
   if type_twoid>0 Then type_two=conn.execute("select [name] From DNJY_type Where id="  & type_oneid&" and twoid="&type_twoid&" and threeid=0")(0)
   if type_twoid>0 and city_threeid>0 Then type_three=conn.execute("select [name] From DNJY_type Where id="  & type_oneid&" and twoid="&type_twoid&" and threeid="&type_threeid)(0)
End if

      If LoginType="" Then
         FoundErr=True
         ErrMsg=ErrMsg & "<br><li>��ѡ����վ��¼����</li>"
      Else
         LoginType=Clng(LoginType)
         If LoginType=1 Then
            If LoginUrl="" Or LoginPostUrl="" Or LoginUser="" Or LoginPass="" Or LoginFalse="" Then
               FoundErr=True
               ErrMsg=ErrMsg & "<br><li>��վ��¼��Ϣ������</li>"
            End If
         End If
      End If
   If FoundErr<>True Then
      SqlItem="Select top 1 *  from DNJY_xx_Item Where ItemID=" & ItemID
      Set RsItem=server.CreateObject("adodb.recordset")
      RsItem.Open SqlItem,conn,2,3
      RsItem("ItemName")=ItemName
      RsItem("WebName")=WebName
      RsItem("WebUrl")=WebUrl
	rsitem("city_oneid") = city_oneid
	  rsitem("city_twoid")=city_twoid
	rsitem("city_threeid") = city_threeid
	rsitem("city_one") = city_one
	rsitem("city_two") = city_two
	rsitem("city_three") = city_three
	rsitem("type_oneid") = type_oneid
	rsitem("type_twoid") = type_twoid
	rsitem("type_threeid") = type_threeid
	rsitem("type_one") = type_one
	rsitem("type_two") = type_two
	rsitem("type_three") = type_three
      RsItem("LoginType")=LoginType
      If Logintype=1 Then
         RsItem("LoginUrl")=LoginUrl
         RsItem("LoginPostUrl")=LoginPostUrl
         RsItem("LoginUser")=LoginUser
         RsItem("LoginPass")=LoginPass
         RsItem("LoginFalse")=LoginFalse
      End If
      RsItem("ItemDemo")=ItemDemo
      RsItem.UpDate
      RsItem.Close
      Set RsItem=Nothing 
   End If
End Sub
%>
