<!--#include file="../conn.asp"-->
<!--#include file="../config.asp"-->
<!--#include file=cookies.asp-->
<!--#include file="inc/Function.asp"-->
<!--#include file="inc/ubbcode.asp"-->

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
Dim RsItem,SqlItem,FoundErr,ErrMsg,Action,ItemID
Dim LoginType,LoginUrl,LoginPostUrl,LoginUser,LoginPass,LoginFalse,LoginResult,LoginData
Dim ListStr,LsString,LoString,ListPaingType,LPsString,LPoString,ListPaingStr1,ListPaingStr2,ListPaingID1,ListPaingID2,ListPaingStr3,HsString,HoString,HttpUrlType,HttpUrlStr
Dim TsString,ToString,CsString,CoString,DateType,DsString,DoString,StopDateType,SDsString,SDoString,StopDate,AuthorType,AsString,AoString,AuthorStr,TelType,TelsString,TeloString,TelStr,sjType,sjsString,sjoString,sjStr,mailType,mailsString,mailoString,mailStr,QQType,QsString,QoString,QQStr,dzType,dzsString,dzoString,dzStr,ybType,ybsString,yboString,ybStr,jgType,jgsString,jgoString,jgStr,CopyFromType,FsString,FoString,CopyFromStr,KeyType,KsString,KoString,KeyStr,NewsPaingType,NPsString,NPoString,NewsPaingStr,NewsPaingHtml,leixingType,LxsString,LxoString,leixingStr
Dim ListUrl,ListCode,NewsArrayCode,NewsArray,UrlTest,NewsCode
Dim Testi
Action=Trim(Request("Action"))
ItemID=strint(Request("ItemID"))
FoundErr=False

If ItemID=0 Then
   FoundErr=True
   ErrMsg=ErrMsg & "<br><li>����������ĿID����Ϊ��</li>"
End If


Call OpenConn
If Action="SaveEdit" And FoundErr<>True Then
   Call SaveEdit()
End If

If FoundErr<>True Then
   Call GetTest()
End If



If FoundErr=True Then
   Call WriteErrMsg(ErrMsg)
Else
   Call Main()
End If
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
<TR>
    <TD height="20" bgcolor="#799AE1" align="center" colspan="13"><FONT color="#FFFFFF" style="font-size:14px">�� �� ϵ ͳ �� Ŀ �� ��</FONT></TD>
 </TR>
  <tr class="tdbg"> 
    <td width="65" height="30"><strong>��������</strong></td>
    <td height="30">��Ŀ�༭ >> <a href="collect_editor_a.asp?ItemID=<%=ItemID%>">��������</a> >> <a href="collect_editor_b.asp?ItemID=<%=ItemID%>">�б�����</a> >> <a href="collect_editor_c.asp?ItemID=<%=ItemID%>">��������</a> >> <a href="collect_editor_d.asp?ItemID=<%=ItemID%>"><font color=red>��������</font></a> >>  
	<a href="collect_editor_e.asp?ItemID=<%=ItemID%>">��������</a> >> <a href="collect_attribute.asp?ItemID=<%=ItemID%>">��������</a> >> ���</td>         
  </tr>         
</table> 
<br>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" class="border" >        
    <tr> 
      <td height="22" colspan="2" class="title"> <div align="center"><strong>�� �� �� Ŀ--�� �� �� �� �� �� �� ��</strong></div></td>
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

<form method="post" action="collect_editor_e.asp" name="form1">
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" class="border" >
    <tr> 
      <td height="22" colspan="2" class="title"> <div align="center"><strong>�� �� �� Ŀ--�� �� �� ��</strong></div>      </td>
    </tr>
    <tr class="tdbg"> 
      <td width="20%" class="tdbg" ><strong>���⿪ʼ��ǣ�</strong><p>��</p><p>��</p>
      <strong>���������ǣ�</strong></td>
      <td class="tdbg" width="75%">
      <textarea name="TsString" cols="49" rows="7"><%=TsString%></textarea><br>
      <textarea name="ToString" cols="49" rows="7"><%=ToString%></textarea></td>
    </tr>
    <tr class="tdbg"> 
      <td width="20%" class="tdbg" ><strong>���Ŀ�ʼ��ǣ�</strong><p>��</p><p>��</p>
      <strong>���Ľ�����ǣ�</strong></td>
      <td class="tdbg" width="75%">
      <textarea name="CsString" cols="49" rows="7"><%=CsString%></textarea><br>
      <textarea name="CoString" cols="49" rows="7"><%=CoString%></textarea></td>
    </tr>
    <tr class="tdbg"> 
      <td width="20%" class="tdbg" ><b>����ʱ�����ã�</b></td>
      <td class="tdbg" width="75%">
      	<input type="radio" value="0" name="DateType" <%If DateType=0 Then%>checked<%End If%> onClick="Date1.style.display='none'">��������&nbsp;
	<input type="radio" value="1" name="DateType" <%If DateType=1 Then%>checked<%End If%> onClick="Date1.style.display=''">���ñ�ǩ&nbsp; ����������ʱȡ��ǰʱ�䣩</td>   
    </tr>
    <tr class="tdbg" id="Date1" style="display:'<%If DateType<>1 Then Response.Write "none" End If%>'"> 
      <td width="20%" class="tdbg" ><strong><font color=blue>ʱ�俪ʼ��ǣ�</font></strong><p>��</p>
		<p>��</p>
      <strong><font color=blue>ʱ�������ǣ�</font></strong></td>
      <td class="tdbg" width="75%">
      <textarea name="DsString" cols="49" rows="7"><%=DsString%></textarea><br>
      <textarea name="DoString" cols="49" rows="7"><%=DoString%></textarea></td>
    </tr>
    <tr class="tdbg"> 
      <td width="20%" class="tdbg" ><b>����ʱ�����ã�</b></td>
      <td class="tdbg" width="75%">
      	<input name="StopDateType" type="radio" onClick="Date2.style.display='none';Date3.style.display=''" value="0" <%If StopDateType=0 Then%>checked<%End If%>>&nbsp;��ǰʱ��+&nbsp;��
   	<input type="radio" value="1" name="StopDateType" onClick="Date2.style.display='';Date3.style.display='none'" <%If StopDateType=1 Then%>checked<%End If%>>���ñ�ǩ&nbsp;</td>   
    </tr>
	
	<tr class="tdbg" id="Date2" <%If StopDateType<>1 Then%>style="display:none"<%End If%>> 
      <td width="20%" class="tdbg" ><strong><font color=blue>����ʱ�俪ʼ��ǣ�</font></strong>
        <p>��</p>
		<p>��</p>
      <strong><font color=blue>����ʱ�������ǣ�</font></strong></td>
      <td class="tdbg" width="75%">
      <textarea name="SDsString" cols="49" rows="7"><%=SDsString%></textarea><br>
      <textarea name="SDoString" cols="49" rows="7"><%=SDoString%></textarea></td>
    </tr>
	
    <tr class="tdbg" id="Date3" <%If StopDateType<>0 Then%>style="display:none"<%End If%>> 
      <td width="20%" class="tdbg" ><strong><font color=blue>��ѡ����ʱ��:</font></strong></td>
      <td class="tdbg" width="75%">
      	<SELECT size="1" name="StopDate" id="StopDate">
          <OPTION value="7" <%If StopDate=7 Then%>selected<%End If%>>һ����</OPTION>
          <OPTION value="15" <%If StopDate=15 Then%>selected<%End If%>>�����</OPTION>
          <OPTION value="30" <%If StopDate=30 Then%>selected<%End If%>>һ����</OPTION>
          <OPTION value="90" <%If StopDate=90 Then%>selected<%End If%>>������</OPTION>
          <OPTION value="180" <%If StopDate=180 Then%>selected<%End If%>>����</OPTION>
          <OPTION value="365" <%If StopDate=365 Then%>selected<%End If%>>һ��</OPTION>
		  <OPTION value="1100" <%If StopDate=1100 Then%>selected<%End If%>>����</OPTION>
        </SELECT></td>
    </tr>
    <tr class="tdbg">
<td width="20%" class="tdbg" ><b><span lang="en-us">���׼۸����ã�</span></b></td>
      <td class="tdbg" width="75%">
      	<input type="radio" value="0" name="jgType" <%If jgType=0 Then%>checked<%End If%> onClick="jg1.style.display='none';jg2.style.display='none'">��������&nbsp;
		<input type="radio" value="1" name="jgType" <%If jgType=1 Then%>checked<%End If%> onClick="jg1.style.display='';jg2.style.display='none'">���ñ�ǩ&nbsp;
		<input type="radio" value="2" name="jgType" <%If jgType=2 Then%>checked<%End If%> onClick="jg1.style.display='none';jg2.style.display=''">ָ�����׼۸�</td>
    </tr>
    <tr class="tdbg" id="jg1" <%If jgType<>1 Then%>style="display:none"<%End If%>> 
      <td width="20%" class="tdbg" ><strong><font color=blue>���׼۸�ʼ��ǣ�</font></strong>
        <p>��</p>
		<p>��</p>
      <strong><font color=blue>���׼۸������ǣ�</font></strong></td>
      <td class="tdbg" width="75%">
      <textarea name="jgsString" cols="49" rows="7"><%=jgsString%></textarea><br>
      <textarea name="jgoString" cols="49" rows="7"><%=jgoString%></textarea></td>
    </tr>
    <tr class="tdbg" id="jg2" <%If jgType<>2 Then%>style="display:none"<%End If%>> 
      <td width="20%" class="tdbg" ><strong><font color=blue>��ָ�����׼۸�</font></strong></td>
      <td class="tdbg" width="75%">
      <input name="jgStr" type="text" id="jgStr" value="<%=jgStr%>" onKeyUp="if(isNaN(value)){alert('���׼۸�ֻ������������');value='';}"   onbeforepaste="clipboardData.setData('text',clipboardData.getData('text').replace(/[^\d|^\.]/g,''))">������Ϊ���֣���</td>
    </tr>
    <tr class="tdbg"> 
      <td width="20%" class="tdbg" ><b>����(��ϵ��)���ã�</b></td>
      <td class="tdbg" width="75%">
      	<input type="radio" value="0" name="AuthorType" <%If AuthorType=0 Then Response.Write "checked"%> onClick="Author1.style.display='none';Author2.style.display='none'">��������&nbsp;
		<input type="radio" value="1" name="AuthorType" <%If AuthorType=1 Then Response.Write "checked"%> onClick="Author1.style.display='';Author2.style.display='none'">���ñ�ǩ&nbsp;
		<input type="radio" value="2" name="AuthorType" <%If AuthorType=2 Then Response.Write "checked"%> onClick="Author1.style.display='none';Author2.style.display=''">ָ������ �������ã�����ɼ�����Ϣ�༭ʱҲҪ�</td>
    </tr>
    <tr class="tdbg" id="Author1" style="display:'<%If AuthorType<>1 Then Response.Write "none"%>'"> 
      <td width="20%" class="tdbg" ><strong><font color=blue>���߿�ʼ��ǣ�</font></strong><p>��</p>
		<p>��</p>
      <strong><font color=blue>���߽�����ǣ�</font></strong></td>
      <td class="tdbg" width="75%">
      <textarea name="AsString" cols="49" rows="7"><%=AsString%></textarea><br>
      <textarea name="AoString" cols="49" rows="7"><%=AoString%></textarea></td>
    </tr>
    <tr class="tdbg" id="Author2" style="display:'<%If AuthorType<>2 Then Response.Write "none"%>'"> 
      <td width="20%" class="tdbg" ><strong><font color=blue>��ָ�����ߣ�</font></strong></td>
      <td class="tdbg" width="75%">
      <input name="AuthorStr" type="text" id="AuthorStr" value="<%=AuthorStr%>">      </td>
    </tr>
    <tr class="tdbg">
<td width="20%" class="tdbg" ><b><span lang="en-us">�̶��绰���ã�</span></b></td>
      <td class="tdbg" width="75%">
      	<input type="radio" value="0" name="TelType" <%If TelType=0 Then%>checked<%End If%> onClick="Tel1.style.display='none';Tel2.style.display='none'">��������&nbsp;
		<input type="radio" value="1" name="TelType" <%If TelType=1 Then%>checked<%End If%> onClick="Tel1.style.display='';Tel2.style.display='none'">���ñ�ǩ&nbsp;
		<input type="radio" value="2" name="TelType" <%If TelType=2 Then%>checked<%End If%> onClick="Tel1.style.display='none';Tel2.style.display=''">ָ���绰�������ã�����ɼ�����Ϣ�༭ʱҲҪ�</td>
    </tr>
    <tr class="tdbg" id="Tel1" <%If TelType<>1 Then%>style="display:none"<%End If%>> 
      <td width="20%" class="tdbg" ><strong><font color=blue>�绰��ʼ��ǣ�</font></strong>
        <p>��</p>
		<p>��</p>
      <strong><font color=blue>�绰������ǣ�</font></strong></td>
      <td class="tdbg" width="75%">
      <textarea name="TelsString" cols="49" rows="7"><%=TelsString%></textarea><br>
      <textarea name="TeloString" cols="49" rows="7"><%=TeloString%></textarea></td>
    </tr>
    <tr class="tdbg" id="Tel2" <%If TelType<>2 Then%>style="display:none"<%End If%>> 
      <td width="20%" class="tdbg" ><strong><font color=blue>��ָ���绰��</font></strong></td>
      <td class="tdbg" width="75%">
      <input name="TelStr" type="text" id="TelStr" value="<%=TelStr%>"></td>
    </tr>

    <tr class="tdbg">
<td width="20%" class="tdbg" ><b><span lang="en-us">�ֻ��������ã�</span></b></td>
      <td class="tdbg" width="75%">
      	<input type="radio" value="0" name="sjType" <%If sjType=0 Then%>checked<%End If%> onClick="sj1.style.display='none';sj2.style.display='none'">��������&nbsp;
		<input type="radio" value="1" name="sjType" <%If sjType=1 Then%>checked<%End If%> onClick="sj1.style.display='';sj2.style.display='none'">���ñ�ǩ&nbsp;
		<input type="radio" value="2" name="sjType" <%If sjType=2 Then%>checked<%End If%> onClick="sj1.style.display='none';sj2.style.display=''">ָ�����루�����ã�����ɼ�����Ϣ�༭ʱҲҪ�</td>
    </tr>
    <tr class="tdbg" id="sj1" <%If sjType<>1 Then%>style="display:none"<%End If%>> 
      <td width="20%" class="tdbg" ><strong><font color=blue>���뿪ʼ��ǣ�</font></strong>
        <p>��</p>
		<p>��</p>
      <strong><font color=blue>���������ǣ�</font></strong></td>
      <td class="tdbg" width="75%">
      <textarea name="sjsString" cols="49" rows="7"><%=sjsString%></textarea><br>
      <textarea name="sjoString" cols="49" rows="7"><%=sjoString%></textarea></td>
    </tr>
    <tr class="tdbg" id="sj2" <%If sjType<>2 Then%>style="display:none"<%End If%>> 
      <td width="20%" class="tdbg" ><strong><font color=blue>��ָ�����룺</font></strong></td>
      <td class="tdbg" width="75%">
      <input name="sjStr" type="text" id="sjStr" value="<%=sjStr%>"></td>
    </tr>

	    <tr class="tdbg">
<td width="20%" class="tdbg" ><b><span lang="en-us">��&nbsp;&nbsp;��&nbsp;&nbsp;��&nbsp;&nbsp;�䣺</span></b></td>
      <td class="tdbg" width="75%">
      	<input type="radio" value="0" name="mailType" <%If mailType=0 Then%>checked<%End If%> onClick="mail1.style.display='none';mail2.style.display='none'">��������&nbsp;
		<input type="radio" value="1" name="mailType" <%If mailType=1 Then%>checked<%End If%> onClick="mail1.style.display='';mail2.style.display='none'">���ñ�ǩ&nbsp;
		<input type="radio" value="2" name="mailType" <%If mailType=2 Then%>checked<%End If%> onClick="mail1.style.display='none';mail2.style.display=''">ָ������</td>
    </tr>
    <tr class="tdbg" id="mail1" <%If mailType<>1 Then%>style="display:none"<%End If%>> 
      <td width="20%" class="tdbg" ><strong><font color=blue>���俪ʼ��ǣ�</font></strong>
        <p>��</p>
		<p>��</p>
      <strong><font color=blue>���������ǣ�</font></strong></td>
      <td class="tdbg" width="75%">
      <textarea name="mailsString" cols="49" rows="7"><%=mailsString%></textarea><br>
      <textarea name="mailoString" cols="49" rows="7"><%=mailoString%></textarea></td>
    </tr>
    <tr class="tdbg" id="mail2" <%If mailType<>2 Then%>style="display:none"<%End If%>> 
      <td width="20%" class="tdbg" ><strong><font color=blue>��ָ�����䣺</font></strong></td>
      <td class="tdbg" width="75%">
      <input name="mailStr" type="text" id="mailStr" value="<%=mailStr%>"></td>
    </tr>

   <tr class="tdbg"> 
      <td width="20%" class="tdbg" ><b><span lang="en-us">QQ&nbsp;&nbsp;��&nbsp;&nbsp;�ã�</span></b></td>
      <td class="tdbg" width="75%">
      	<input type="radio" value="0" name="QQType" <%If QQType=0 Then%>checked<%End If%> onClick="QQ1.style.display='none';QQ2.style.display='none'">��������&nbsp;
		<input type="radio" value="1" name="QQType" <%If QQType=1 Then%>checked<%End If%> onClick="QQ1.style.display='';QQ2.style.display='none'">���ñ�ǩ&nbsp;
		<input type="radio" value="2" name="QQType" <%If QQType=2 Then%>checked<%End If%> onClick="QQ1.style.display='none';QQ2.style.display=''">
		ָ��QQ</td>
    </tr>
    <tr class="tdbg" id="QQ1" <%If QQType<>1 Then%>style="display:none"<%End If%>> 
      <td width="20%" class="tdbg" ><strong><font color=blue>QQ��ʼ��ǣ�</font></strong>
        <p>��</p>
		<p>��</p>
      <strong><font color=blue>QQ������ǣ�</font></strong></td>
      <td class="tdbg" width="75%">
      <textarea name="QsString" cols="49" rows="7"><%=QsString%></textarea><br>
      <textarea name="QoString" cols="49" rows="7"><%=QoString%></textarea></td>
    </tr>
    <tr class="tdbg" id="QQ2" <%If QQType<>2 Then%>style="display:none"<%End If%>> 
      <td width="20%" class="tdbg" ><strong><font color=blue>��ָ��QQ��</font></strong></td>
      <td class="tdbg" width="75%">
      <input name="QQStr" type="text" id="QQStr" value="<%=QQStr%>"></td>
    </tr>
    <tr class="tdbg">
<td width="20%" class="tdbg" ><b><span lang="en-us">ͨѶ��ַ���ã�</span></b></td>
      <td class="tdbg" width="75%">
      	<input type="radio" value="0" name="dzType" <%If dzType=0 Then%>checked<%End If%> onClick="dz1.style.display='none';dz2.style.display='none'">��������&nbsp;
		<input type="radio" value="1" name="dzType" <%If dzType=1 Then%>checked<%End If%> onClick="dz1.style.display='';dz2.style.display='none'">���ñ�ǩ&nbsp;
		<input type="radio" value="2" name="dzType" <%If dzType=2 Then%>checked<%End If%> onClick="dz1.style.display='none';dz2.style.display=''">ָ����ַ</td>
    </tr>
    <tr class="tdbg" id="dz1" <%If dzType<>1 Then%>style="display:none"<%End If%>> 
      <td width="20%" class="tdbg" ><strong><font color=blue>��ַ��ʼ��ǣ�</font></strong>
        <p>��</p>
		<p>��</p>
      <strong><font color=blue>��ַ������ǣ�</font></strong></td>
      <td class="tdbg" width="75%">
      <textarea name="dzsString" cols="49" rows="7"><%=dzsString%></textarea><br>
      <textarea name="dzoString" cols="49" rows="7"><%=dzoString%></textarea></td>
    </tr>
    <tr class="tdbg" id="dz2" <%If dzType<>2 Then%>style="display:none"<%End If%>> 
      <td width="20%" class="tdbg" ><strong><font color=blue>��ָ����ַ��</font></strong></td>
      <td class="tdbg" width="75%">
      <input name="dzStr" type="text" id="dzStr" value="<%=dzStr%>">����100�ַ��ڣ�</td>
    </tr>
    <tr class="tdbg">
<td width="20%" class="tdbg" ><b><span lang="en-us">�����������ã�</span></b></td>
      <td class="tdbg" width="75%">
      	<input type="radio" value="0" name="ybType" <%If ybType=0 Then%>checked<%End If%> onClick="yb1.style.display='none';yb2.style.display='none'">��������&nbsp;
		<input type="radio" value="1" name="ybType" <%If ybType=1 Then%>checked<%End If%> onClick="yb1.style.display='';yb2.style.display='none'">���ñ�ǩ&nbsp;
		<input type="radio" value="2" name="ybType" <%If ybType=2 Then%>checked<%End If%> onClick="yb1.style.display='none';yb2.style.display=''">ָ���ʱ�</td>
    </tr>
    <tr class="tdbg" id="yb1" <%If ybType<>1 Then%>style="display:none"<%End If%>> 
      <td width="20%" class="tdbg" ><strong><font color=blue>�ʱ࿪ʼ��ǣ�</font></strong>
        <p>��</p>
		<p>��</p>
      <strong><font color=blue>�ʱ������ǣ�</font></strong></td>
      <td class="tdbg" width="75%">
      <textarea name="ybsString" cols="49" rows="7"><%=ybsString%></textarea><br>
      <textarea name="yboString" cols="49" rows="7"><%=yboString%></textarea></td>
    </tr>
    <tr class="tdbg" id="yb2" <%If ybType<>2 Then%>style="display:none"<%End If%>> 
      <td width="20%" class="tdbg" ><strong><font color=blue>��ָ���ʱࣺ</font></strong></td>
      <td class="tdbg" width="75%">
      <input name="ybStr" type="text" id="ybStr" value="<%=ybStr%>"></td>
    </tr>
    <tr class="tdbg"> 
      <td width="20%" class="tdbg" ><b>��&nbsp; Դ&nbsp;
		��&nbsp; �ã�</b></td>
      <td class="tdbg" width="75%">
      	<input type="radio" value="0" name="CopyFromType" <%If CopyFromType=0 Then Response.Write "checked"%> onClick="CopyFrom1.style.display='none';CopyFrom2.style.display='none'">��������&nbsp;
		<input type="radio" value="1" name="CopyFromType" <%If CopyFromType=1 Then Response.Write "checked"%> onClick="CopyFrom1.style.display='';CopyFrom2.style.display='none'">���ñ�ǩ&nbsp;
		<input type="radio" value="2" name="CopyFromType" <%If CopyFromType=2 Then Response.Write "checked"%> onClick="CopyFrom1.style.display='none';CopyFrom2.style.display=''">ָ����Դ&nbsp;
		<input type="radio" value="3" name="CopyFromType" <%If CopyFromType=3 Then Response.Write "checked"%> onClick="CopyFrom1.style.display='none';CopyFrom2.style.display='none'">���ɼ�ҳ</td>
    </tr>
    <tr class="tdbg" id="CopyFrom1" style="display:'<%If CopyFromType<>1 Then Response.Write "none"%>'"> 
      <td width="20%" class="tdbg" ><strong><font color=blue>��Դ��ʼ��ǣ�</font></strong><p>��</p>
		<p>��</p>
      <strong><font color=blue>��Դ������ǣ�</font></strong></td>
      <td class="tdbg" width="75%">
      <textarea name="FsString" cols="49" rows="7"><%=FsString%></textarea><br>
      <textarea name="FoString" cols="49" rows="7"><%=FoString%></textarea></td>
    </tr>
    <tr class="tdbg" id="CopyFrom2" style="display:'<%If CopyFromType<>2 Then Response.Write "none"%>'"> 
      <td width="20%" class="tdbg" ><strong><font color=blue>��ָ����Դ��</font></strong></td>
      <td class="tdbg" width="75%">
      <input name="CopyFromStr" type="text" id="CopyFromStr" value="<%=CopyFromStr%>"></td>
    </tr>

   <tr class="tdbg"> 
      <td width="20%" class="tdbg" ><b><span lang="en-us">�� Ϣ �� �ͣ�</span></b></td>
      <td class="tdbg" width="75%">
      	<input type="radio" value="0" name="leixingType" <%If leixingType=0 Then%>checked<%End If%> onClick="leixing1.style.display='none';leixing2.style.display='none'">��������&nbsp;
		<input type="radio" value="1" name="leixingType" <%If leixingType=1 Then%>checked<%End If%> onClick="leixing1.style.display='';leixing2.style.display='none'">���ñ�ǩ&nbsp;
		<input type="radio" value="2" name="leixingType" <%If leixingType=2 Then%>checked<%End If%> onClick="leixing1.style.display='none';leixing2.style.display=''">
		ָ������</td>
    </tr>
    <tr class="tdbg" id="leixing1" <%If leixingType<>1 Then%>style="display:none"<%End If%>> 
      <td width="20%" class="tdbg" ><strong><font color=blue>���Ϳ�ʼ��ǣ�</font></strong>
        <p>��</p>
		<p>��</p>
      <strong><font color=blue>���ͽ�����ǣ�</font></strong></td>
      <td class="tdbg" width="75%">
      <textarea name="LxsString" cols="49" rows="7"><%=LxsString%></textarea><br>
      <textarea name="LxoString" cols="49" rows="7"><%=LxoString%></textarea></td>
    </tr>
    <tr class="tdbg" id="leixing2" <%If leixingType<>2 Then%>style="display:none"<%End If%>> 
      <td width="20%" class="tdbg" ><strong><font color=blue>��ָ�����ͣ�</font></strong></td>
      <td class="tdbg" width="75%"><%Dim rslx,sqllx,leixing,x,le
set rslx=server.createobject("adodb.recordset")
sqllx = "select leixing from DNJY_config "
rslx.open sqllx,conn,1,1
leixing=split(rslx("leixing"),"|")
response.write "<select name=""leixingStr"" id=""leixingStr""><option value=>����</option>"
for x=0 to ubound(leixing)
le=""
If leixingStr=leixing(x) then le="selected"
response.write "<option value="""&leixing(x)&""" "&le&">"&leixing(x)&"</option>"
next
response.write "</select>"
rslx.close
set rslx=nothing%></td>
    </tr>

    <tr class="tdbg"> 
      <td colspan="2" align="center" class="tdbg"><br>
        <input name="Action" type="hidden" id="Action" value="SaveEdit">
        <input name="ItemID" type="hidden" id="ItemID" value="<%=ItemID%>">
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
<%
Sub SaveEdit
   HsString=Request.Form("HsString")
   HoString=Request.Form("HoString")
   HttpUrlType=strint(Request.Form("HttpUrlType"))
   HttpUrlStr=Trim(Request.Form("HttpUrlStr"))

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
      If HttpUrlType=1 Then
         If HttpUrlStr="" Then
            FoundErr=True
            ErrMsg=ErrMsg & "<br><li>�����þ������ӵ�ַ</li>"
         Else
            If Len(HttpUrlStr)<15 Then
               FoundErr=True
               ErrMsg=ErrMsg & "<br><li>�������ӵ�ַ���ò���ȷ(����15���ַ�)</li>"
            End If
         End If
      End If
   End If

   If FoundErr<>True Then
      SqlItem="Select ItemID,HsString,HoString,HttpUrlType,HttpUrlStr from DNJY_xx_Item Where ItemID=" & ItemID
      Set RsItem=server.CreateObject("adodb.recordset")
      RsItem.Open SqlItem,Conn,2,3
      RsItem("HsString")=HsString
      RsItem("HoString")=HoString
      RsItem("HttpUrlType")=HttpUrlType
      If HttpUrlType=1 Then
         RsItem("HttpUrlStr")=HttpUrlStr
      End If
      RsItem.UpDate
      RsItem.Close:Set RsItem=Nothing
   End If
End Sub

Sub GetTest


   SqlItem="Select * from DNJY_xx_Item Where ItemID=" & ItemID
   Set RsItem=server.CreateObject("adodb.recordset")
   RsItem.Open SqlItem,Conn,1,1
   If RsItem.Eof And RsItem.Bof Then
      FoundErr=True
      ErrMsg=ErrMsg & "<br><li>����������ĿID����Ϊ��</li>"
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
      HsString=RsItem("HsString")
      HoString=RsItem("HoString")
      HttpUrlType=RsItem("HttpUrlType")
      HttpUrlStr=RsItem("HttpUrlStr")
      TsString=RsItem("TsString")
      ToString=RsItem("ToString")
      CsString=RsItem("CsString")
      CoString=RsItem("CoString")
      
      DateType=RsItem("DateType")
      DsString=RsItem("DsString")
      DoString=RsItem("DoString")
      
      StopDateType=RsItem("StopDateType")
      SDsString=RsItem("SDsString")
      SDoString=RsItem("SDoString")
	  StopDate=RsItem("StopDate")
	  
      AuthorType=RsItem("AuthorType")
      AsString=RsItem("AsString")
      AoString=RsItem("AoString")
      AuthorStr=RsItem("AuthorStr")

      TelType=RsItem("TelType")
      TelsString=RsItem("TelsString")
      TeloString=RsItem("TeloString")
      TelStr=RsItem("TelStr")

      sjType=RsItem("sjType")
      sjsString=RsItem("sjsString")
      sjoString=RsItem("sjoString")
      sjStr=RsItem("sjStr")

      mailType=RsItem("mailType")
      mailsString=RsItem("mailsString")
      mailoString=RsItem("mailoString")
      mailStr=RsItem("mailStr")
	  
      QQType=RsItem("QQType")
      QsString=RsItem("QsString")
      QoString=RsItem("QoString")
      QQStr=RsItem("QQStr")

      dzType=RsItem("dzType")
      dzsString=RsItem("dzsString")
      dzoString=RsItem("dzoString")
      dzStr=RsItem("dzStr")

      ybType=RsItem("ybType")
      ybsString=RsItem("ybsString")
      yboString=RsItem("yboString")
      ybStr=RsItem("ybStr")

      jgType=RsItem("jgType")
      jgsString=RsItem("jgsString")
      jgoString=RsItem("jgoString")
      jgStr=RsItem("jgStr")
	  
      CopyFromType=RsItem("CopyFromType")
      FsString=RsItem("FsString")
      FoString=RsItem("FoString")
      CopyFromStr=RsItem("CopyFromStr")

      leixingType=RsItem("leixingType")
      LxsString=RsItem("LxsString")
      LxoString=RsItem("LxoString")
      leixingStr=RsItem("leixingStr")

   End If
   RsItem.Close
   Set RsItem=Nothing

   If LsString="" Then
      FoundErr=True
      ErrMsg=ErrMsg & "<br><li>�б�ʼ��ǲ���Ϊ�գ�</li>"
   End If
   If LoString="" Then
      FoundErr=True
      ErrMsg=ErrMsg & "<br><li>�б������ǲ���Ϊ�գ�</li>"
   End If
   If ListPaingType=0 Or ListPaingType=1 Then
      If ListStr="" Then
         FoundErr=True
         ErrMsg=ErrMsg & "<br><li>�б�����ҳ����Ϊ�գ�</li>"
      End If
      If ListPaingType=1 Then    
         If LPsString="" Or LPoString="" Then
            FoundErr=True
            ErrMsg=ErrMsg & "<br><li>������ҳ��ʼ��������ǲ���Ϊ�գ�</li>"
         End If
      End If      
      If  ListPaingStr1<>""  And  Len(ListPaingStr1)<15  Then
         FoundErr=True
         ErrMsg=ErrMsg & "<br><li>������ҳ�ض������ò���ȷ��</li>"
            End  IF
   ElseIf ListPaingType=2 Then
      If ListPaingStr2="" Then
         FoundErr=True
         ErrMsg=ErrMsg & "<br><li>��������ԭ�ַ�������Ϊ�գ�</li>"
      End If
      If IsNumeric(ListPaingID1)=False or IsNumeric(ListPaingID2)=False Then
         FoundErr=True
         ErrMsg=ErrMsg & "<br><li>�������ɵķ�Χֻ�������֣�</li>"
      Else
         ListPaingID1=Clng(ListPaingID1)
         ListPaingID2=Clng(ListPaingID2)
         If ListPaingID1=0 And ListPaingID2=0 Then
            FoundErr=True
            ErrMsg=ErrMsg & "<br><li>�������ɵķ�Χ����ȷ��</li>"
         End If
      End If 
   ElseIf ListPaingType=3 Then
      If ListPaingStr3="" Then
         FoundErr=True
         ErrMsg=ErrMsg & "<br><li>������ҳ����Ϊ�գ�</li>"
      End If
   Else
      FoundErr=True
      ErrMsg=ErrMsg & "<br><li>��ѡ�񷵻���һ������������ҳ����</li>"
   End If
 
   If LoginType=1 Then
      If LoginUrl="" or LoginPostUrl="" or LoginUser="" Or LoginPass="" Or LoginFalse="" Then
         FoundErr=True
         ErrMsg=ErrMsg & "<br><li>�뽫��¼��Ϣ��д����</li>"
      End If
   End If

   If FoundErr<>True Then
      Select Case ListPaingType
      Case 0,1
            ListUrl=ListStr
      Case 2
         ListUrl=Replace(ListPaingStr2,"{$ID}",CStr(ListPaingID1))
      Case 3
         If Instr(ListPaingStr3,"|")> 0 Then
            ListUrl=Left(ListPaingStr3,Instr(ListPaingStr3,"|")-1)
         Else
            ListUrl=ListPaingStr3
         End If
      End Select
   End If

      If  FoundErr<>True  And Action<>"SaveEdit" And  LoginType=1  Then
      LoginData=UrlEncoding(LoginUser & "&" & LoginPass)
      LoginResult=PostHttpPage(LoginUrl,LoginPostUrl,LoginData)
      If Instr(LoginResult,LoginFalse)>0 Then
         FoundErr=True
         ErrMsg=ErrMsg & "<br><li>��¼��վʱ����������ȷ�ϵ�¼��Ϣ����ȷ�ԣ�</li>"
      End If
      End  If
      
   If FoundErr<>True Then
      ListCode=GetHttpPage(ListUrl)
      If ListCode<>"$False$" Then
         ListCode=GetBody(ListCode,LsString,LoString,False,False)
         If ListCode="$False$" Then
            FoundErr=True
            ErrMsg=ErrMsg & "<br><li>�ڽ�ȡ�б�ʱ��������</li>"
         End If
      Else
         FoundErr=True
         ErrMsg=ErrMsg & "<br><li>�ڻ�ȡ:" & ListUrl & "��ҳԴ��ʱ��������</li>"
      End If
   End If

   If FoundErr<>True Then
      NewsArrayCode=GetArray(ListCode,HsString,HoString,False,False)
      If NewsArrayCode="$False$" Then
         FoundErr=True
         ErrMsg=ErrMsg & "<br><li>�ڷ�����" & ListUrl & "�����б�ʱ��������</li>"
      Else
         NewsArray=Split(NewsArrayCode,"$Array$")
         If IsArray(NewsArray)=True Then
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
            ErrMsg=ErrMsg & "<br><li>�ڷ�����" & ListUrl & "�����б�ʱ��������</li>"
         End If            
      End If
   End If 
End Sub
conn.close:Set conn=Nothing%>
