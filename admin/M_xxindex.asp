<!--#include file="../conn.asp"-->
<!--#include file=cookies.asp-->
<!--#include file="../config.asp"-->
<!--#include file="../Include/Function.asp"-->
<!--#include file="../Include/label.asp"-->
<!--#include file="../Include/Form_board .asp"-->
<%
'=====================================
'���칩����Ϣ��վ����ϵͳ
'��������Ƽ������ҳ�Ʒ
'�ͷ�����http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>
<%
Dim Action,ID,fso,Ts,Data,Moldboard,i,Temp,TextFile,Text,M_logo,mb,M_,M(109)'��ǰ��������ʵ�����õ�109
Moldboard=Request.Form("Moldboard")
ID=strint(request.QueryString("ID"))
Call OpenConn
IF Request.Form("Submit")<>"" Then
For i=0 to 109'��ǰ������
'IF i=72 Then'�������ı�����ɫ��
'm(i)=HtmlEncode2(Request.Form("M"&i))
'Else
m(i)=strint(Request.Form("M"&i))
'End IF
Temp=Temp&m(i)
IF i<>109 Then Temp=Temp&"|"'��ǰ������
Next

Set Rs=server.createobject("adodb.recordset")
Sql="Select ID,M_xxIndex,M_xxIndexstr From DNJY_Template Where ID="&ID
Rs.Open Sql,conn,1,3
Rs(1)=Moldboard
Rs(2)=Temp
Rs.Update

set fso = CreateObject("Scripting.FileSystemObject")
set Ts = Fso.OpenTextFile(Server.Mappath("../xxIndex.asp"),2, True)
Ts.Write M_xxIndex(Moldboard,Temp)
Ts.Close:Set Ts=Nothing
Set Fso=Nothing
If Htmlhome=1 Then Call HomeHtml()'����������ҳ
response.write "<script language='javascript'>	alert('ģ�����óɹ������������µ�ҳ�棡');location.href='javascript:history.go(-1)';</script>"
End IF

Set Rs=Conn.Execute("Select M_xxIndex,M_xxIndexstr,M_Name,M_logo From DNJY_Template Where ID="&ID)
IF Not Rs.Eof Then Data=Rs.GetRows
Rs.close:Set Rs=Nothing
Call CloseDB(conn)
M_logo=Data(3,0)
M_=Split(Data(1,0),"|")
%>
<HTML>
<HEAD>
<LINK rel="stylesheet" type="text/css" href="../css/style.css">
<SCRIPT language="javascript" src="admin.js"></SCRIPT>
<META http-equiv="Content-Type" content="text/html; charset=gb2312">
<META name="GENERATOR" content="Microsoft FrontPage 6.0">
<META name="ProgId" content="FrontPage.Editor.Document">
<TITLE>��ϢƵ��ҳģ��</title>
<style type="text/css">
<!--
body {
	background-color: #D6DFF7;
}
body,td,th {
	color: #135294;
}
-->
</style></HEAD>
<BODY background="images/background.gif">  
                                                              
  <TABLE width="98%" border="0" align="center" cellpadding="4" cellspacing="1" bgcolor="#799AE1">
    <TR>
      <TD height="20" align="center" bgcolor="#799AE1"><FONT color="#FFFFFF" style="font-size:14px">�� Ϣ Ƶ �� ҳ ģ �� �� ��</FONT></TD>
    </TR>
    <TR>
      <TD bgcolor="#FFFFFF">
            <TABLE width="98%" border="0"  align=center cellpadding="1" cellspacing="1" bgcolor="#D6DFF7">
  <FORM name="form1" method="post" action="">
    <TR bgcolor="#FFFFFF">
      <TD width="20%" height="26" align="right">ģ�����</TD>
      <TD width="60%" bgcolor="#FFFFFF" valign="top">��ǰģ�壺��ϢƵ��ҳxxindex.asp����ǰģ�巽����<FONT color="#ff0000"><%=Data(2,0)%></font>
       &nbsp;&nbsp;ģ���ʶ��<%=M_logo%><br> <TEXTAREA name="Moldboard" cols="90" rows="30" id="T2"><%=Data(0,0)%></TEXTAREA></TD>
		<TD width="20%" height="26" align="left" >ģ���ǩ��<br>
      ��ͨ�����±�ǩ����������ݣ�<p>
<FONT color="#ff0000">{$������}<br>
{$������}<br>
{$������Ϣ}<br>
{$������Ϣ}<br>
{$�Ƽ���Ϣ}<br>
{$������Ϣ}<br>
{$�Ƽ�ͼƬ��Ϣ}<br>
{$����ͼƬ��Ϣ}<br>
{$����ͼƬ��Ϣ}<br>
{$����ͼƬ��Ϣ}<br>
{$�ö�ͼƬ��Ϣ}<br>
{$������ʾ��Ϣ}<br>
{$��Ϣ����}<br>
{$��Ϣ�б�}<br>
</font><br> <br>
<a href="#" ONCLICK="window.open('../Include/datetype.htm','Sample','toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=yes,resizable=no,copyhistory=no,width=200,height=400,left=300,top=20')"><b>0-15��������ʽ</b></a> </TD></TR>
    <TR bgcolor="#FFFFFF">
      <TD height="26" align="right"><b>ע�⣺</b></TD>
      <TD colspan="2"><b>���´�[��]Ϊ��ǰģ����Ч,���ò�����Ч����[<FONT color=#ff0000>��</font>]Ϊ��ǰģ����Ч,���ò�����Ч��</b></TD>
    </TR>
	<TR bgcolor="#FFFFFF">
      <TD height="30" align="right"  bgcolor="#BDDEEF">վ����������<%mb="A,B,C,D,E,F"
	  If instr(mb,M_logo)>0 Then
	  response.write "[��]"
	  Else
	  response.write "[<FONT color=#ff0000>��</font>]"
	  End if%></TD>
      <TD colspan="2" bgcolor="#BDDEEF">&nbsp;������ʾ��
        <SELECT name="m0" id="m0">
          <OPTION value="1" <%IF StrInt(m_(0))=1 Then%>selected<%End IF%>>��ʾһ��(�Ƽ�)</OPTION>
          <OPTION value="2" <%IF StrInt(m_(0))=2 Then%>selected<%End IF%>>��ʾһ���Ͷ���</OPTION>
          <OPTION value="3" <%IF StrInt(m_(0))=3 Then%>selected<%End IF%>>��ʾȫ��(���Ƽ�)</OPTION>
        </SELECT>
        ������ʾ��
        <SELECT name="m1" id="m1">
          <OPTION value="1" <%IF StrInt(m_(1))=1 Then%>selected<%End IF%>>��ʾһ��(�Ƽ�)</OPTION>
          <OPTION value="2" <%IF StrInt(m_(1))=2 Then%>selected<%End IF%>>��ʾһ���Ͷ���</OPTION>
          <OPTION value="3" <%IF StrInt(m_(1))=3 Then%>selected<%End IF%>>��ʾȫ��(���Ƽ�)</OPTION>
        </SELECT>����ȫ�������б��ٶ�̫����������ʾĬ����ʾһ��������ѡ����Ч��</TD>
    </TR>
    <TR bgcolor="#FFFFFF">
      <TD height="26" align="right">������Ϣ��ʾ���ã�<%mb="A,B,C,D,E,F"
	  If instr(mb,M_logo)>0 Then
	  response.write "[��]"
	  Else
	  response.write "[<FONT color=#ff0000>��</font>]"
	  End if%></TD>
      <TD colspan="2">
	        ��Ϣ���ͣ�<INPUT name="m2" type="checkbox" id="m2" value="1" <%IF strint(m_(2))=1 Then%>checked<%End IF%>>
			�Ƽ���־��<INPUT name="m3" type="checkbox" id="m3" value="1" <%IF strint(m_(3))=1 Then%>checked<%End IF%>>
           	ͼƬ��־��<INPUT name="m5" type="checkbox" id="m5" value="1" <%IF strint(m_(5))=1 Then%>checked<%End IF%>>
			���ù̶���<INPUT name="m4" type="checkbox" id="m4" value="1" <%IF strint(m_(4))=1 Then%>checked<%End IF%>>
			������ʾ��<INPUT name="m6" type="checkbox" id="m6" value="1" <%IF strint(m_(6))=1 Then%>checked<%End IF%>>
			��ʾ���ࣺ<INPUT name="m7" type="checkbox" id="m7" value="1" <%IF strint(m_(7))=1 Then%>checked<%End IF%>>
			������ɫ��<INPUT name="m8" type="checkbox" id="m8" value="1" <%IF strint(m_(8))=1 Then%>checked<%End IF%>>
			�ż㱳����<INPUT name="m9" type="checkbox" id="m9" value="1" <%IF strint(m_(9))=1 Then%>checked<%End IF%>>
			���������<INPUT name="m10" type="checkbox" id="m10" value="1" <%IF strint(m_(10))=1 Then%>checked<%End IF%>><br>
			������ʽ(0-15)��<a href="#" ONCLICK="window.open('../Include/datetype.htm','Sample','toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=yes,resizable=no,copyhistory=no,width=200,height=400,left=300,top=20')">������</a><INPUT name="m11" type="text" id="m11" value="<%=m_(11)%>" size="5">
����������<INPUT name="m12" type="text" id="m12" value="<%=m_(12)%>" size="5">
��ʾ������<INPUT name="m13" type="text" id="m13" value="<%=m_(13)%>" size="5"> 
��ʾ������<INPUT name="m14" type="text" id="m14" value="<%=m_(14)%>" size="6">
�����С��<INPUT name="m15" type="text" value="<%=m_(15)%>" size="5">px
 �оࣺ<INPUT name="m16" type="text" value="<%=m_(16)%>" size="5">px</TD>
    </TR>
    <TR bgcolor="#FFFFFF">
      <TD height="26" align="right" bgcolor="#BDDEEF">������Ϣ��ʾ���ã�<%mb="A,B,C,D,E,F"
	  If instr(mb,M_logo)>0 Then
	  response.write "[��]"
	  Else
	  response.write "[<FONT color=#ff0000>��</font>]"
	  End if%></TD>
      <TD colspan="2" bgcolor="#BDDEEF">
	        ��Ϣ���ͣ�<INPUT name="m17" type="checkbox" id="m17" value="1" <%IF strint(m_(17))=1 Then%>checked<%End IF%>>
			�Ƽ���־��<INPUT name="m18" type="checkbox" id="m18" value="1" <%IF strint(m_(18))=1 Then%>checked<%End IF%>>
           	ͼƬ��־��<INPUT name="m20" type="checkbox" id="m20" value="1" <%IF strint(m_(20))=1 Then%>checked<%End IF%>>
			���ù̶���<INPUT name="m19" type="checkbox" id="m19" value="1" <%IF strint(m_(19))=1 Then%>checked<%End IF%>>
			������ʾ��<INPUT name="m21" type="checkbox" id="m21" value="1" <%IF strint(m_(21))=1 Then%>checked<%End IF%>>
			��ʾ���ࣺ<INPUT name="m22" type="checkbox" id="m22" value="1" <%IF strint(m_(22))=1 Then%>checked<%End IF%>>
			������ɫ��<INPUT name="m23" type="checkbox" id="m23" value="1" <%IF strint(m_(23))=1 Then%>checked<%End IF%>>
			�ż㱳����<INPUT name="m24" type="checkbox" id="m24" value="1" <%IF strint(m_(24))=1 Then%>checked<%End IF%>>
			���������<INPUT name="m25" type="checkbox" id="m25" value="1" <%IF strint(m_(25))=1 Then%>checked<%End IF%>><br>
			������ʽ(0-15)��<a href="#" ONCLICK="window.open('../Include/datetype.htm','Sample','toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=yes,resizable=no,copyhistory=no,width=200,height=400,left=300,top=20')">������</a><INPUT name="m26" type="text" id="m26" value="<%=m_(26)%>" size="5">
����������<INPUT name="m27" type="text" id="m27" value="<%=m_(27)%>" size="5">
��ʾ������<INPUT name="m28" type="text" id="m28" value="<%=m_(28)%>" size="5"> 
��ʾ������<INPUT name="m29" type="text" id="m29" value="<%=m_(29)%>" size="6">
�����С��<INPUT name="m30" type="text" value="<%=m_(30)%>" size="5">px
 �оࣺ<INPUT name="m31" type="text" value="<%=m_(31)%>" size="5">px
	        </TD>
    </TR>
    <TR bgcolor="#FFFFFF">
      <TD height="26" align="right">�Ƽ���Ϣ��ʾ���ã�<%mb="A,B,C,D,E,F"
	  If instr(mb,M_logo)>0 Then
	  response.write "[��]"
	  Else
	  response.write "[<FONT color=#ff0000>��</font>]"
	  End if%></TD>
      <TD colspan="2">
	        ��Ϣ���ͣ�<INPUT name="m32" type="checkbox" id="m32" value="1" <%IF strint(m_(32))=1 Then%>checked<%End IF%>>			       
			�Ƽ���־��<INPUT name="m33" type="checkbox" id="m33" value="1" <%IF strint(m_(33))=1 Then%>checked<%End IF%>>
			ͼƬ��־��<INPUT name="m35" type="checkbox" id="m35" value="1" <%IF strint(m_(35))=1 Then%>checked<%End IF%>>
			���ù̶���<INPUT name="m34" type="checkbox" id="m34" value="1" <%IF strint(m_(34))=1 Then%>checked<%End IF%>>
			������ʾ��<INPUT name="m36" type="checkbox" id="m36" value="1" <%IF strint(m_(36))=1 Then%>checked<%End IF%>>
			��ʾ���ࣺ<INPUT name="m37" type="checkbox" id="m37" value="1" <%IF strint(m_(37))=1 Then%>checked<%End IF%>>
			������ɫ��<INPUT name="m38" type="checkbox" id="m38" value="1" <%IF strint(m_(38))=1 Then%>checked<%End IF%>>
			�ż㱳����<INPUT name="m39" type="checkbox" id="m39" value="1" <%IF strint(m_(39))=1 Then%>checked<%End IF%>>
			���������<INPUT name="m40" type="checkbox" id="m40" value="1" <%IF strint(m_(40))=1 Then%>checked<%End IF%>>     <br>
			������ʽ(0-15)��<a href="#" ONCLICK="window.open('../Include/datetype.htm','Sample','toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=yes,resizable=no,copyhistory=no,width=200,height=400,left=300,top=20')">������</a><INPUT name="m41" type="text" id="m41" value="<%=m_(41)%>" size="5">
����������<INPUT name="m42" type="text" id="m42" value="<%=m_(42)%>" size="5">
��ʾ������<INPUT name="m43" type="text" id="m43" value="<%=m_(43)%>" size="5"> 
��ʾ������<INPUT name="m44" type="text" id="m44" value="<%=m_(44)%>" size="6">
�����С��<INPUT name="m45" type="text" value="<%=m_(45)%>" size="5">px
 �оࣺ<INPUT name="m46" type="text" value="<%=m_(46)%>" size="5">px
</TD>
    </TR>
    <TR bgcolor="#FFFFFF">
      <TD height="26" align="right" bgcolor="#BDDEEF">������Ϣ��ʾ���ã�<%mb="A,B,C,D,E,F"
	  If instr(mb,M_logo)>0 Then
	  response.write "[��]"
	  Else
	  response.write "[<FONT color=#ff0000>��</font>]"
	  End if%></TD>
      <TD colspan="2" bgcolor="#BDDEEF">
	  	    
			��Ϣ���ͣ�<INPUT name="m47" type="checkbox" id="m47" value="1" <%IF strint(m_(47))=1 Then%>checked<%End IF%>>            
			�Ƽ���־��<INPUT name="m48" type="checkbox" id="m48" value="1" <%IF strint(m_(48))=1 Then%>checked<%End IF%>>
			ͼƬ��־��<INPUT name="m50" type="checkbox" id="m50" value="1" <%IF strint(m_(50))=1 Then%>checked<%End IF%>>
			���ù̶���<INPUT name="m49" type="checkbox" id="m49" value="1" <%IF strint(m_(49))=1 Then%>checked<%End IF%>>
			������ʾ��<INPUT name="m51" type="checkbox" id="m51" value="1" <%IF strint(m_(51))=1 Then%>checked<%End IF%>>
			��ʾ���ࣺ<INPUT name="m52" type="checkbox" id="m52" value="1" <%IF strint(m_(52))=1 Then%>checked<%End IF%>>
			������ɫ��<INPUT name="m53" type="checkbox" id="m53" value="1" <%IF strint(m_(53))=1 Then%>checked<%End IF%>>
			�ż㱳����<INPUT name="m54" type="checkbox" id="m54" value="1" <%IF strint(m_(54))=1 Then%>checked<%End IF%>>
			���������<INPUT name="m55" type="checkbox" id="m55" value="1" <%IF strint(m_(55))=1 Then%>checked<%End IF%>><br>
			������ʽ(0-15)��<a href="#" ONCLICK="window.open('../Include/datetype.htm','Sample','toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=yes,resizable=no,copyhistory=no,width=200,height=400,left=300,top=20')">������</a><INPUT name="m56" type="text" id="m56" value="<%=m_(56)%>" size="5">
����������<INPUT name="m57" type="text" id="m57" value="<%=m_(57)%>" size="5">
��ʾ������<INPUT name="m58" type="text" id="m58" value="<%=m_(58)%>" size="5"> 
��ʾ������<INPUT name="m59" type="text" id="m59" value="<%=m_(59)%>" size="6">
�����С��<INPUT name="m60" type="text" value="<%=m_(60)%>" size="5">px
 �оࣺ<INPUT name="m61" type="text" value="<%=m_(61)%>" size="5">px</TD>
    </TR>
    <TR bgcolor="#FFFFFF">
      <TD height="26" align="right">�Ƽ�ͼƬ��Ϣ��ʾ��<%mb="A,B,C,D,E,F"
	  If instr(mb,M_logo)>0 Then
	  response.write "[��]"
	  Else
	  response.write "[<FONT color=#ff0000>��</font>]"
	  End if%></TD>
      <TD colspan="2">
ÿ����ʾ������<INPUT name="m62" type="text" id="m62" value="<%=m_(62)%>" size="5">
�ܹ���ʾ������<INPUT name="m63" type="text" id="m63" value="<%=m_(63)%>" size="5">
		</TD>
    </TR>
    <TR bgcolor="#FFFFFF">
      <TD height="26" align="right" bgcolor="#BDDEEF">����ͼƬ��Ϣ��ʾ��<%mb="A,B,C,D,E,F"
	  If instr(mb,M_logo)>0 Then
	  response.write "[��]"
	  Else
	  response.write "[<FONT color=#ff0000>��</font>]"
	  End if%></TD>
      <TD colspan="2" bgcolor="#BDDEEF">
ÿ����ʾ������<INPUT name="m64" type="text" id="m64" value="<%=m_(64)%>" size="5">
�ܹ���ʾ������<INPUT name="m65" type="text" id="m65" value="<%=m_(65)%>" size="5">
		</TD>
    </TR>
    <TR bgcolor="#FFFFFF">
      <TD height="26" align="right">����ͼƬ��Ϣ��ʾ��<%mb="A,B,C,D,E,F"
	  If instr(mb,M_logo)>0 Then
	  response.write "[��]"
	  Else
	  response.write "[<FONT color=#ff0000>��</font>]"
	  End if%></TD>
      <TD colspan="2">
ÿ����ʾ������<INPUT name="m66" type="text" id="m66" value="<%=m_(66)%>" size="5">
�ܹ���ʾ������<INPUT name="m67" type="text" id="m67" value="<%=m_(67)%>" size="5">
		</TD>
    </TR>
    <TR bgcolor="#FFFFFF">
      <TD height="26" align="right" bgcolor="#BDDEEF">����ͼƬ��Ϣ��ʾ��<%mb="A,B,C,D,E,F"
	  If instr(mb,M_logo)>0 Then
	  response.write "[��]"
	  Else
	  response.write "[<FONT color=#ff0000>��</font>]"
	  End if%></TD>
      <TD colspan="2" bgcolor="#BDDEEF">
ÿ����ʾ������<INPUT name="m68" type="text" id="m68" value="<%=m_(68)%>" size="5">
�ܹ���ʾ������<INPUT name="m69" type="text" id="m69" value="<%=m_(69)%>" size="5">
		</TD>
    </TR>
    <TR bgcolor="#FFFFFF">
      <TD height="26" align="right">�ö�ͼƬ��Ϣ��ʾ��<%mb="A,B,C,D,E,F"
	  If instr(mb,M_logo)>0 Then
	  response.write "[��]"
	  Else
	  response.write "[<FONT color=#ff0000>��</font>]"
	  End if%></TD>
      <TD colspan="2">
ÿ����ʾ������<INPUT name="m70" type="text" id="m70" value="<%=m_(70)%>" size="5">
�ܹ���ʾ������<INPUT name="m71" type="text" id="m71" value="<%=m_(71)%>" size="5">
		</TD>
    </TR>
    <TR bgcolor="#FFFFFF">
      <TD height="26" align="right" bgcolor="#BDDEEF">������Ϣ�������ã�<%mb="A,B,C,D,E,F"
	  If instr(mb,M_logo)>0 Then
	  response.write "[��]"
	  Else
	  response.write "[<FONT color=#ff0000>��</font>]"
	  End if%></TD>
      <TD colspan="2" bgcolor="#BDDEEF">
	        ������ɫ��<INPUT name="m72" type="checkbox" id="m72" value="1" <%IF strint(m_(72))=1 Then%>checked<%End IF%>>
			��Ϣ���ͣ�<INPUT name="m73" type="checkbox" id="m73" value="1" <%IF strint(m_(73))=1 Then%>checked<%End IF%>>            
			��ϢLOGO��<INPUT name="m74" type="checkbox" id="m74" value="1" <%IF strint(m_(74))=1 Then%>checked<%End IF%>>
           	���ù̶���<INPUT name="m75" type="checkbox" id="m75" value="1" <%IF strint(m_(75))=1 Then%>checked<%End IF%>>
			��ʾ��Ч�ڣ�<INPUT name="m76" type="checkbox" id="m76" value="1" <%IF strint(m_(76))=1 Then%>checked<%End IF%>>
			ͼƬ��־��<INPUT name="m77" type="checkbox" id="m77" value="1" <%IF strint(m_(77))=1 Then%>checked<%End IF%>>
			���������<INPUT name="m78" type="checkbox" id="m78" value="1" <%IF strint(m_(78))=1 Then%>checked<%End IF%>>
			��ʾ��Ա��<INPUT name="m79" type="checkbox" id="m79" value="1" <%IF strint(m_(79))=1 Then%>checked<%End IF%>><br>
			��ʾ���ۣ�<INPUT name="m80" type="checkbox" id="m80" value="1" <%IF strint(m_(80))=1 Then%>checked<%End IF%>>

���ⳤ�ȣ�<INPUT name="m81" type="text" id="m81" value="<%=m_(81)%>" size="5">
��鳤�ȣ�<INPUT name="m82" type="text" id="m82" value="<%=m_(82)%>" size="5"> 
��ʾ������<INPUT name="m83" type="text" id="m83" value="<%=m_(83)%>" size="6">
��ʾ������<INPUT name="m84" type="text" id="m84" value="<%=m_(84)%>" size="5">
			��Ϣ��Χ��
            <SELECT name="m85" id="m85">
              <OPTION value="0" <%IF strint(m_(85))=0 Then%>selected<%End IF%>>��ʾȫ����Ϣ</OPTION>
              <OPTION value="1" <%IF strint(m_(85))=1 Then%>selected<%End IF%>>����ʾ������Ϣ</OPTION>			  
            </SELECT>
</TD>
    </TR>
	<TR bgcolor="#FFFFFF"> 
      <TD height="26" align="right">��Ϣ������ʾ���ã�<%mb="A,B,C,D,E,F"
	  If instr(mb,M_logo)>0 Then
	  response.write "[��]"
	  Else
	  response.write "[<FONT color=#ff0000>��</font>]"
	  End if%></TD>
      <TD colspan="2" >
	        ����ID�ţ�<INPUT name="m86" type="text" id="m86" value="<%=m_(86)%>" size="5">
	        ��ʾһ�����ࣺ<INPUT name="m87" type="checkbox" id="m87" value="1" <%IF strint(m_(87))=1 Then%>checked<%End IF%>>
			��ʾ�������ࣺ<INPUT name="m88" type="checkbox" id="m88" value="1" <%IF strint(m_(88))=1 Then%>checked<%End IF%>>
	  	    һ�����������<INPUT name="m89" type="text" id="m89" value="<%=m_(89)%>" size="5">
	  	    �������������<INPUT name="m90" type="text" id="m90" value="<%=m_(90)%>" size="5">
			��һ��������Ϣ����<INPUT name="m91" type="checkbox" id="m91" value="1" <%IF strint(m_(91))=1 Then%>checked<%End IF%>><br>
			�Զ���������Ϣ����<INPUT name="m92" type="checkbox" id="m92" value="1" <%IF strint(m_(92))=1 Then%>checked<%End IF%>>
			һ������������<INPUT name="m93" type="text" id="m93" value="<%=m_(93)%>" size="5">	  	    
			һ�������ֺţ�<INPUT name="m94" type="text" id="m94" value="<%=m_(94)%>" size="5">
			���������ֺţ�<INPUT name="m95" type="text" id="m95" value="<%=m_(95)%>" size="5">
			һ�������оࣺ<INPUT name="m96" type="text" id="m96" value="<%=m_(96)%>" size="5">
			���������оࣺ<INPUT name="m97" type="text" id="m97" value="<%=m_(97)%>" size="5">	  	    
			</TD>
    </TR>
    <TR bgcolor="#FFFFFF"> 
      <TD height="26" align=right bgcolor="#BDDEEF">������Ϣ�б���ʾ���ã�<%mb="A,B,C,D,E,F"
	  If instr(mb,M_logo)>0 Then
	  response.write "[��]"
	  Else
	  response.write "[<FONT color=#ff0000>��</font>]"
	  End if%></TD>
      <TD width="73%" colspan="2" align=left bgcolor="#BDDEEF">
	  	    ��Ϣ���ͣ�<INPUT name="m98" type="checkbox" value="1" <%IF strint(m_(98))=1 Then%>checked<%End IF%>>
			���������<INPUT name="m99" type="text" value="<%=m_(99)%>" size="5">
			��ʾ������<INPUT name="m100" type="text" value="<%=m_(100)%>" size="5">
            ��ʾ������<INPUT name="m101" type="text" value="<%=m_(101)%>" size="5">
            ���ⳤ�ȣ�<INPUT name="m102" type="text" value="<%=m_(102)%>" size="5">
	  	    ������ɫ��<INPUT name="m103" type="checkbox" value="1" <%IF strint(m_(103))=1 Then%>checked<%End IF%>>
			���ù̶���<INPUT name="m104" type="checkbox"value="1" <%IF strint(m_(104))=1 Then%>checked<%End IF%>>
	  	    ͼƬ��־��<INPUT name="m105" type="checkbox" value="1" <%IF strint(m_(105))=1 Then%>checked<%End IF%>>
	  	    �������<INPUT name="m106" type="checkbox" value="1" <%IF strint(m_(106))=1 Then%>checked<%End IF%>><br>			
	  	    ������ʽ(0-15)��<a href="#" ONCLICK="window.open('../Include/datetype.htm','Sample','toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=yes,resizable=no,copyhistory=no,width=200,height=400,left=300,top=20')">������</a>
			<INPUT name="m107" type="text" value="<%=m_(107)%>" size="5">
			ͷ��ͼ�ģ�<INPUT name="m108" type="checkbox" value="1" <%IF strint(m_(108))=1 Then%>checked<%End IF%>>	  
            �ż㱳����
            <SELECT name="m109" >
              <OPTION value="0" <%IF StrInt(m_(109))=0 Then%>selected<%End IF%>>Ҫ�ż㱳��</OPTION>
              <OPTION value="1" <%IF StrInt(m_(109))=1 Then%>selected<%End IF%>>���ż㱣��ʽ</OPTION>
			  <OPTION value="2" <%IF StrInt(m_(109))=2 Then%>selected<%End IF%>>���ż��޸�ʽ</OPTION>
            </SELECT>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<FONT color="#ff0000">�п�Ҫ��ģ�����ã�CSS������ҵ�shadow1���޸ĸö�width:��ֵ������ҳ����</font></TD>
    </TR>

    <TR bgcolor="#FFFFFF"> 
      <TD height="30" colspan="3" align=center>	  
      <input name="Submit" type="submit" id="Submit" value="ȷ��">
	  <input name="Submit1" type="button" id="Submit1" value="����" onClick="location='admin_Template_Management.asp?id=<%=ID%>'"/>&nbsp;&nbsp;&nbsp;<a href="#" ONCLICK="window.open('../Include/datetype.htm','Sample','toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=yes,resizable=no,copyhistory=no,width=200,height=400,left=300,top=20')"><b>0-15��������ʽ</b></a> </TD>
      </TR>
  </FORM>
</TABLE>
      <BR></TD>
    </TR>
</TABLE>
</BODY>                    
                    
</HTML>                    
