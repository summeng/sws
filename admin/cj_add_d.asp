<!--#include file="../conn.asp"-->
<!--#include file=cookies.asp-->
<!--#include file=../config.asp-->
<!--#include file="../Include/head.asp"-->
<!--#include file="../Include/Function.asp"-->
<!--#include file="../Include/label.asp"-->
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
Call OpenConn
Dim ItemID
Dim RsItem,SqlItem,FoundErr,ErrMsg
Dim ListStr,LsString,LoString
Dim  ListPaingType,LPsString,LPoString,ListPaingStr1,ListPaingStr2,ListPaingID1,ListPaingID2,ListPaingStr3,HsString,HoString,HttpUrlType,HttpUrlStr
Dim  ListUrl,ListCode,NewsArrayCode,NewsArray,UrlTest,NewsCode
dim Testi
ItemID=Trim(Request("ItemID"))
HsString=Request.Form("HsString")
HoString=Request.Form("HoString")
HttpUrlType=Trim(Request.Form("HttpUrlType"))
HttpUrlStr=Trim(Request.Form("HttpUrlStr"))

If ItemID="" Then
   FoundErr=True
   ErrMsg=ErrMsg & "<br><li>�������������Ч���ӽ���</li>"
Else
   ItemID=Clng(ItemID)
End If
If HsString="" Then
   FoundErr=True
   ErrMsg=ErrMsg & "<br><li>���ӿ�ʼ��ǲ���Ϊ��</li>"
End If
If HoString="" Then
   FoundErr=True
   ErrMsg=ErrMsg & "<br><li>���ӽ�����ǲ���Ϊ��</li>" 
End If
If HttpUrlType="" Then
   FoundErr=True
   ErrMsg=ErrMsg & "<br><li>��ѡ�����Ӵ�������</li>"
Else
   HttpUrlType=Clng(HttpUrlType)
   If HttpUrlType=1 Then
      If HttpUrlStr="" Then
         FoundErr=True
         ErrMsg=ErrMsg & "<br><li>���������ַ����ò���Ϊ��</li>"
      Else
            If  Len(HttpUrlStr)<15  Then
               FoundErr=True
               ErrMsg=ErrMsg & "<br><li>���������ַ����ò���ȷ(15���ַ�����)</li>"
            End  If      
      End If
   End If
End If

If FoundErr<>True Then
   SqlItem="Select ListStr,LsString,LoString,ListPaingType,LPsString,LPoString,ListPaingStr1,ListPaingStr2,ListPaingID1,ListPaingID2,ListPaingStr3,HsString,HoString,HttpUrlType,HttpUrlStr from DNJY_Article_Item Where ItemID=" & ItemID
   Set RsItem=server.CreateObject("adodb.recordset")
   RsItem.Open SqlItem,Conn,2,3
   RsItem("HsString")=HsString
   RsItem("HoString")=HoString
   RsItem("HttpUrlType")=HttpUrlType
   If HttpUrlType=1 Then
   RsItem("HttpUrlStr")=HttpUrlStr
   End If
   ListStr=RsItem("ListStr")
   LsString=RsItem("LsString")
   LoString=RsItem("LoString")
   ListPaingType=RsItem("ListPaingType")
   LPsString=RsItem("LPsString")
   ListPaingStr1=RsItem("ListPaingStr1")
   ListPaingStr2=RsItem("ListPaingStr2")
   ListPaingID1=RsItem("ListPaingID1")
   ListPaingID2=RsItem("ListPaingID2")
   ListPaingStr3=RsItem("ListPaingStr3")
   RsItem.UpDate
   RsItem.Close
   Set RsItem=Nothing
   
   Select  Case  ListPaingType
   Case  0,1
         ListUrl=ListStr
   Case  2
         ListUrl=Replace(ListPaingStr2,"{$ID}",CStr(ListPaingID1))
   Case  3
         If  Instr(ListPaingStr3,"|")>0  Then
         ListUrl=Left(ListPaingStr3,Instr(ListPaingStr3,"|")-1)
      Else
         ListUrl=ListPaingStr3
      End  If
   End  Select

   ListCode=GetHttpPage(ListUrl)
   If ListCode<>"$False$" Then
      ListCode=GetBody(ListCode,LsString,LoString,False,False)
      If ListCode<>"$False$" Then 
         NewsArrayCode=GetArray(ListCode,HsString,HoString,False,False)
         If NewsArrayCode<>"$False$" Then
               NewsArray=Split(NewsArrayCode,"$Array$")
               For Testi=0 To Ubound(NewsArray)
                  If HttpUrlType=1 Then
                     NewsArray(Testi)=Replace(HttpUrlStr,"{$ID}",NewsArray(Testi))
                  Else
                     NewsArray(Testi)=DefiniteUrl(NewsArray(Testi),ListUrl)
                  End If
               Next
               UrlTest=NewsArray(0)
               NewsCode=GetHttpPage(UrlTest)
         Else
            FoundErr=True
            ErrMsg=ErrMsg & "<br><li>�ڻ�ȡ�����б�ʱ����</li>"
         End If   
      Else
         FoundErr=True
         ErrMsg=ErrMsg & "<br><li>�ڽ�ȡ�б�ʱ��������</li>"
      End If
   Else
      FoundErr=True
      ErrMsg=ErrMsg & "<br><li>�ڻ�ȡ:" & ListUrl & "��ҳԴ��ʱ��������</li>"
   End If
End If

If FoundErr=True Then
   Call WriteErrMsg(ErrMsg)
Else
   Call Main()
End If 
'�ر����ݿ�����
conn.close:Set conn=nothing
%>
<%Sub Main()%>
<html>
<head>
<title>�ɼ�ϵͳ</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="../css/manage.css" type="text/css">
</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" class="border">
  <tr> 
    <td height="22" colspan="2" align="center" class="topbg"><strong>���Ųɼ�ϵͳ��Ŀ����</td>
  </tr>
  <tr class="tdbg"> 
    <td width="65" height="30"><strong>��������</strong></td>
    <td height="30"><a href="cj_add_a.asp">�����Ŀ</a> >> <a href="cj_editor_a.asp?ItemID=<%=ItemID%>">��������</a> >> <a href="cj_editor_b.asp?ItemID=<%=ItemID%>">�б�����</a> >> <a href="cj_editor_c.asp?ItemID=<%=ItemID%>">��������</a> >> <font color=red>��������</font> >> �������� >> �������� >> ���</td>         
  </tr>         
</table> 
<br>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" class="border" >        
    <tr> 
      <td height="22" colspan="2" class="title"> <div align="center"><strong>�� �� �� 
		�� Ŀ--�� �� �� �� �� �� �� ��</strong></div></td>
    </tr>
</table>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" class="border" >
    <tr> 
      <td height="22" colspan="2" class="tdbg">�����Ƿ��������õ������ž������ӵ�ַ����鿴�Ƿ���ȷ��<br>
<%
For Testi=0 To Ubound(NewsArray)
   Response.Write "<a href='" & NewsArray(Testi) & "' target=_blank>" & NewsArray(Testi) & "</a><br>"
Next
%>
<br>
��һ������ȡ��һ�����Ž��в��ԣ�����д���±��ʱ������Ҫʹ�õ�һ�����š�
      </td>
    </tr>
</table>
<form method="post" action="cj_add_e.asp" name="form1">
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" class="border" >
    <tr> 
      <td height="22" colspan="2" class="title"> <div align="center"><strong>�� �� �� 
		�� Ŀ--�� �� �� ��</strong></div>      </td>
    </tr>
    <tr class="tdbg"> 
      <td width="20%" class="tdbg"  align="right"><strong>���⿪ʼ��ǣ�</strong><p>��</p><p>��</p>
      <strong>���������ǣ�</strong></td>
      <td class="tdbg" width="75%">
      <textarea name="TsString" cols="49" rows="7"></textarea><br>
      <textarea name="ToString" cols="49" rows="7"></textarea></td>
    </tr>
    <tr class="tdbg"> 
      <td width="20%" class="tdbg" align="right" ><strong>���Ŀ�ʼ��ǣ�</strong><p>��</p><p>��</p>
      <strong>���Ľ�����ǣ�</strong></td>
      <td class="tdbg" width="75%">
      <textarea name="CsString" cols="49" rows="7"></textarea><br>
      <textarea name="CoString" cols="49" rows="7"></textarea></td>
    </tr>
    <tr class="tdbg"> 
      <td width="20%" class="tdbg"  align="right"><b>����ʱ�����ã�</b></td>
      <td class="tdbg" width="75%">
      	<input type="radio" value="0" name="DateType" checked onClick="Date1.style.display='none'">��ǰʱ��&nbsp;
	    <input type="radio" value="1" name="DateType" onClick="Date1.style.display=''">���ñ�ǩ&nbsp;    ����������ʱȡ��ǰʱ�䣩</tr>
    <tr class="tdbg" id="Date1" style="display:none"> 
      <td width="20%" class="tdbg" align="right" ><strong><font color=blue>����ʱ�俪ʼ��ǣ�</font></strong><p>��</p>
		<p>��</p>
      <strong><font color=blue>����ʱ�������ǣ�</font></strong></td>
      <td class="tdbg" width="75%">
      <textarea name="DsString" cols="49" rows="7"></textarea><br>
      <textarea name="DoString" cols="49" rows="7"></textarea></td>
    </tr>

    <tr class="tdbg"> 
      <td width="20%" class="tdbg"  align="right"><b><span lang="en-us">�������ã�</span></b></td>
      <td class="tdbg" width="75%">
      	<input type="radio" value="0" name="AuthorType" checked onClick="Author1.style.display='none';Author2.style.display='none'">��������&nbsp;
		<input type="radio" value="1" name="AuthorType" onClick="Author1.style.display='';Author2.style.display='none'">���ñ�ǩ&nbsp;
		<input type="radio" value="2" name="AuthorType" onClick="Author1.style.display='none';Author2.style.display=''">ָ�����ߣ������ã�����ɼ�����Ϣ�༭ʱҲҪ�</td>
    </tr>
    <tr class="tdbg" id="Author1" style="display:none"> 
      <td width="20%" class="tdbg" align="right" ><strong><font color=blue>���߿�ʼ��ǣ�</font></strong><p>��</p>
		<p>��</p>
      <strong><font color=blue>���߽�����ǣ�</font></strong></td>
      <td class="tdbg" width="75%">
      <textarea name="AsString" cols="49" rows="7"></textarea><br>
      <textarea name="AoString" cols="49" rows="7"></textarea></td>
    </tr>
    <tr class="tdbg" id="Author2" style="display:none"> 
      <td width="20%" class="tdbg" align="right" ><strong><font color=blue>��ָ�����ߣ�</font></strong></td>
      <td class="tdbg" width="75%">
      <input name="AuthorStr" type="text" id="AuthorStr" value="">      </td>
    </tr>

    <tr class="tdbg"> 
      <td width="20%" class="tdbg"  align="right"><b>��&nbsp; Դ&nbsp;
		��&nbsp; �ã�</b></td>
      <td class="tdbg" width="75%">
      	<input type="radio" value="0" name="CopyFromType" checked onClick="CopyFrom1.style.display='none';CopyFrom2.style.display='none'">��������&nbsp;
		<input type="radio" value="1" name="CopyFromType" onClick="CopyFrom1.style.display='';CopyFrom2.style.display='none'">���ñ�ǩ&nbsp;
		<input type="radio" value="2" name="CopyFromType" onClick="CopyFrom1.style.display='none';CopyFrom2.style.display=''">ָ����Դ
		<input type="radio" value="3" name="CopyFromType" onClick="CopyFrom1.style.display='none';CopyFrom2.style.display='none'">
		���ɼ�ҳ</td>
    </tr>
    <tr class="tdbg" id="CopyFrom1" style="display:none"> 
      <td width="20%" class="tdbg" align="right" ><strong><font color=blue>��Դ��ʼ��ǣ�</font></strong><p>��</p>
		<p>��</p>
      <strong><font color=blue>��Դ������ǣ�</font></strong></td>
      <td class="tdbg" width="75%">
      <textarea name="FsString" cols="49" rows="7"></textarea><br>
      <textarea name="FoString" cols="49" rows="7"></textarea></td>
    </tr>
    <tr class="tdbg" id="CopyFrom2" style="display:none"> 
      <td width="20%" class="tdbg"  align="right"><strong><font color=blue>��ָ����Դ��</font></strong></td>
      <td class="tdbg" width="75%">
      <input name="CopyFromStr" type="text" id="CopyFromStr" value="">      </td>
    </tr>

    <tr class="tdbg"> 
      <td width="20%" class="tdbg" align="right" ><b><span lang="en-us">�绰���ã�</span></b></td>
      <td class="tdbg" width="75%">
      	<input type="radio" value="0" name="TelType" checked onClick="Tel1.style.display='none';Tel2.style.display='none'">��������&nbsp;
		<input type="radio" value="1" name="TelType" onClick="Tel1.style.display='';Tel2.style.display='none'">���ñ�ǩ&nbsp;
		<input type="radio" value="2" name="TelType" onClick="Tel1.style.display='none';Tel2.style.display=''">
		ָ���绰</td>
    </tr>
    <tr class="tdbg" id="Tel1" style="display:none"> 
      <td width="20%" class="tdbg"  align="right"><strong><font color=blue>�绰��ʼ��ǣ�</font></strong>
        <p>��</p>
		<p>��</p>
      <strong><font color=blue>�绰������ǣ�</font></strong></td>
      <td class="tdbg" width="75%">
      <textarea name="TelsString" cols="49" rows="7"></textarea><br>
      <textarea name="TeloString" cols="49" rows="7"></textarea></td>
    </tr>
    <tr class="tdbg" id="Tel2" style="display:none"> 
      <td width="20%" class="tdbg"  align="right"><strong><font color=blue>��ָ���绰��</font></strong></td>
      <td class="tdbg" width="75%">
      <input name="TelStr" type="text" id="TelStr" value="">      </td>
    </tr>

	 <tr class="tdbg">
<td width="20%" class="tdbg"  align="right"><b><span lang="en-us">�������䣺</span></b></td>
      <td class="tdbg" width="75%">
      	<input type="radio" value="0" name="mailType" checked onClick="mail1.style.display='none';mail2.style.display='none'">��������&nbsp;
		<input type="radio" value="1" name="mailType" onClick="mail1.style.display='';mail2.style.display='none'">���ñ�ǩ&nbsp;
		<input type="radio" value="2" name="mailType" onClick="mail1.style.display='none';mail2.style.display=''">ָ�����䣨�����ã�����ɼ�����Ϣ�༭ʱҲҪ�</td>
    </tr>
    <tr class="tdbg" id="mail1" style="display:none"> 
      <td width="20%" class="tdbg"  align="right"><strong><font color=blue>���俪ʼ��ǣ�</font></strong>
        <p>��</p>
		<p>��</p>
      <strong><font color=blue>���������ǣ�</font></strong></td>
      <td class="tdbg" width="75%">
      <textarea name="mailsString" cols="49" rows="7"></textarea><br>
      <textarea name="mailoString" cols="49" rows="7"></textarea></td>
    </tr>
    <tr class="tdbg" id="mail2" style="display:none"> 
      <td width="20%" class="tdbg"  align="right"><strong><font color=blue>��ָ�����䣺</font></strong></td>
      <td class="tdbg" width="75%">
      <input name="mailStr" type="text" id="mailStr" value=""></td>
    </tr>

   <tr class="tdbg"> 
      <td width="20%" class="tdbg" align="right" ><b><span lang="en-us">QQ���ã�</span></b></td>
      <td class="tdbg" width="75%">
      	<input type="radio" value="0" name="QQType" checked onClick="QQ1.style.display='none';QQ2.style.display='none'">��������&nbsp;
		<input type="radio" value="1" name="QQType" onClick="QQ1.style.display='';QQ2.style.display='none'">���ñ�ǩ&nbsp;
		<input type="radio" value="2" name="QQType" onClick="QQ1.style.display='none';QQ2.style.display=''">
		ָ��QQ</td>
    </tr>
    <tr class="tdbg" id="QQ1" style="display:none"> 
      <td width="20%" class="tdbg"  align="right"><strong><font color=blue>QQ��ʼ��ǣ�</font></strong>
        <p>��</p>
		<p>��</p>
      <strong><font color=blue>QQ������ǣ�</font></strong></td>
      <td class="tdbg" width="75%">
      <textarea name="QsString" cols="49" rows="7"></textarea><br>
      <textarea name="QoString" cols="49" rows="7"></textarea></td>
    </tr>
    <tr class="tdbg" id="QQ2" style="display:none"> 
      <td width="20%" class="tdbg" align="right" ><strong><font color=blue>��ָ��QQ��</font></strong></td>
      <td class="tdbg" width="75%">
      <input name="QQStr" type="text" id="QQStr" value="">      </td>
    </tr>
    <tr class="tdbg"> 
      <td colspan="2" align="center" class="tdbg"><br>
          <input name="ItemID" type="hidden" value="<%=ItemID%>"> 
        <input  type="button" name="button1" value="��&nbsp;һ&nbsp;��" onClick="window.location.href='javascript:history.go(-1)'"  style="cursor: hand;background-color: #cccccc;">
        &nbsp;&nbsp;&nbsp;&nbsp; 
        <input  type="submit" name="Submit" value="��&nbsp;һ&nbsp;��" style="cursor: hand;background-color: #cccccc;"></td>
        <input type="hidden" name="UrlTest" id="UrlTest" value="<%=UrlTest%>">
    </tr>
</table>
</form>     
</body>         
</html>
<%End Sub%>
