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
Dim Action,ID,fso,Ts,Data,Moldboard,i,Temp,TextFile,Text,M_logo,mb,M_,M(295)'��ǰʵ���õ���������
Moldboard=Request.Form("Moldboard")
ID=strint(request.QueryString("ID"))
Call OpenConn
IF Request.Form("Submit")<>"" Then
For i=0 to 295'��ǰ������
IF i=176 Then'��ҵ������Ƭ������ɫ
m(i)=HtmlEncode2(Request.Form("M"&i))
Else
m(i)=strint(Request.Form("M"&i))
End IF
Temp=Temp&m(i)
IF i<>295 Then Temp=Temp&"|"'��ǰ������
Next

Set Rs=server.createobject("adodb.recordset")
Sql="Select ID,M_Index,M_Indexstr From DNJY_Template Where ID="&ID
Rs.Open Sql,conn,1,3
Rs(1)=Moldboard
Rs(2)=Temp
Rs.Update

set fso = CreateObject("Scripting.FileSystemObject")
set Ts = Fso.OpenTextFile(Server.Mappath("../Index.asp"),2, True)
Ts.Write M_Index(Moldboard,Temp)
Ts.Close:Set Ts=Nothing
Set Fso=Nothing
If Htmlhome=1 Then Call HomeHtml()'����������ҳ
response.write "<script language='javascript'>	alert('ģ�����óɹ������������µ�ҳ��\n\n���������ɾ�̬����ҳ��');location.href='javascript:history.go(-1)';</script>"
End IF

Set Rs=Conn.Execute("Select M_Index,M_Indexstr,M_Name,M_logo From DNJY_Template Where ID="&ID)
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
<TITLE>��ҳģ��</title>
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
<IFRAME width="260" height="165" id="colourPalette" src="nc_selcolor.htm" style="visibility:hidden; position: absolute; left: 0px; top: 0px;border:1px gray solid" frameborder="0" scrolling="no" ></IFRAME>                                                               
  <TABLE width="98%" border="0" align="center" cellpadding="4" cellspacing="1" bgcolor="#799AE1">
    <TR>
      <TD height="20" align="center" bgcolor="#799AE1"><FONT color="#FFFFFF" style="font-size:14px">�� ҳ ģ �� �� ��</FONT></TD>
    </TR>
    <TR>
      <TD bgcolor="#FFFFFF">
            <TABLE width="98%" border="0"  align=center cellpadding="1" cellspacing="1" bgcolor="#D6DFF7">
  <FORM name="form1" method="post" action="">
    <TR bgcolor="#FFFFFF">
      <TD width="20%" height="26" align="right">ģ�����</TD>
      <TD width="60%" bgcolor="#FFFFFF" valign="top">��ǰģ�壺��ҳģ��index.asp����ǰģ�巽����<FONT color="#ff0000"><%=Data(2,0)%></font>
        &nbsp;&nbsp;ģ���ʶ��<%=M_logo%><br><TEXTAREA name="Moldboard" cols="90" rows="30" id="T2"><%=Data(0,0)%></TEXTAREA></TD>
		<TD width="20%" height="26" align="left" >ģ���ǩ��<br>
      ��ͨ�����±�ǩ����������ݣ�<p>
<FONT color="#ff0000">{$������}<br>
{$������Ϣ}<br>
{$������Ϣ}<br>
{$�Ƽ���Ϣ}<br>
{$������Ϣ}<br>
{$������Ϣ}<br>
{$������ʾ��Ϣ}<br>
{$ͼƬ��Ϣ�Ƽ�}<br>
{$������Ϣͼ��}<br>
{$�Ƽ���Ϣͼ��}<br>
{$������Ϣͼ��}<br>
{$�����Ƽ�}<br>
{$������Դ����}<br>
{$��ҵ������Ƭ}<br>
{$��ҵͼƬ��Ƭ}<br>
{$������Ѷһ}<br>
{$������Ѷ��}<br>
{$������Ѷ��}<br>
{$������Ѷ��}<br>
{$������Ѷ���}<br>
{$������־}<br>
{$����114}<br>
{$�������}<br>
{$��������}<br>
{$�û�����}<br>
{$��վ����}<br>
</font><br> <br>
<a href="#" ONCLICK="window.open('../Include/datetype.htm','Sample','toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=yes,resizable=no,copyhistory=no,width=200,height=400,left=300,top=20')"><b>0-15����ʾ����</b></a> 
</TD></TR>
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
      <TD colspan="2" bgcolor="#BDDEEF">&nbsp;��Ϣ������ʾ��
        <SELECT name="m0">
          <OPTION value="1" <%IF StrInt(m_(0))=1 Then%>selected<%End IF%>>��ʾһ��(�Ƽ�)</OPTION>
          <OPTION value="2" <%IF StrInt(m_(0))=2 Then%>selected<%End IF%>>��ʾһ���Ͷ���</OPTION>
          <OPTION value="3" <%IF StrInt(m_(0))=3 Then%>selected<%End IF%>>��ʾȫ��(���Ƽ�)</OPTION>
        </SELECT>
        ����������ʾ��
        <SELECT name="m1">
          <OPTION value="1" <%IF StrInt(m_(1))=1 Then%>selected<%End IF%>>��ʾһ��(�Ƽ�)</OPTION>
          <OPTION value="2" <%IF StrInt(m_(1))=2 Then%>selected<%End IF%>>��ʾһ���Ͷ���</OPTION>
          <OPTION value="3" <%IF StrInt(m_(1))=3 Then%>selected<%End IF%>>��ʾȫ��(���Ƽ�)</OPTION>
        </SELECT>������������̫���б��ٶȻ�����������ѡ����ʾ���ټ���</TD>
    </TR>
    <TR bgcolor="#FFFFFF">
      <TD height="26" align="right">������Ϣ��ʾ���ã�<%mb="A,B,C,D,E"
	  If instr(mb,M_logo)>0 Then
	  response.write "[��]"
	  Else
	  response.write "[<FONT color=#ff0000>��</font>]"
	  End if%></TD>
      <TD colspan="2">
	        ��Ϣ���ͣ�<INPUT name="m2" type="checkbox" value="1" <%IF strint(m_(2))=1 Then%>checked<%End IF%>>
			���ù̶���<INPUT name="m3" type="checkbox" value="1" <%IF strint(m_(3))=1 Then%>checked<%End IF%>>
			�Ƽ���־��<INPUT name="m4" type="checkbox" value="1" <%IF strint(m_(4))=1 Then%>checked<%End IF%>>
           	ͼƬ��־��<INPUT name="m5" type="checkbox" value="1" <%IF strint(m_(5))=1 Then%>checked<%End IF%>>			
			������ʾ��<INPUT name="m6" type="checkbox" value="1" <%IF strint(m_(6))=1 Then%>checked<%End IF%>>
			��ʾ���ࣺ<INPUT name="m7" type="checkbox" value="1" <%IF strint(m_(7))=1 Then%>checked<%End IF%>>
			������ɫ��<INPUT name="m8" type="checkbox" value="1" <%IF strint(m_(8))=1 Then%>checked<%End IF%>>
			�ż㱳����<INPUT name="m9" type="checkbox" value="1" <%IF strint(m_(9))=1 Then%>checked<%End IF%>>
			���������<INPUT name="m10" type="checkbox" value="1" <%IF strint(m_(10))=1 Then%>checked<%End IF%>><br>
			��ʾ����(0-15)��<a href="#" ONCLICK="window.open('../Include/datetype.htm','Sample','toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=yes,resizable=no,copyhistory=no,width=200,height=400,left=300,top=20')">������</a>
			<INPUT name="m11" type="text" value="<%=m_(11)%>" size="5">
����������<INPUT name="m12" type="text" value="<%=m_(12)%>" size="5">
��ʾ������<INPUT name="m13" type="text" value="<%=m_(13)%>" size="5"> 
��ʾ������<INPUT name="m14" type="text" value="<%=m_(14)%>" size="6">
�����С��<INPUT name="m15" type="text" value="<%=m_(15)%>" size="5">px
 �оࣺ<INPUT name="m16" type="text" value="<%=m_(16)%>" size="5">px</TD>
    </TR>

    <TR bgcolor="#FFFFFF">
      <TD height="26" align="right" bgcolor="#BDDEEF">������Ϣ��ʾ���ã�<%mb="A,B,C,D,E"
	  If instr(mb,M_logo)>0 Then
	  response.write "[��]"
	  Else
	  response.write "[<FONT color=#ff0000>��</font>]"
	  End if%></TD>
      <TD colspan="2" bgcolor="#BDDEEF">
	        ��Ϣ���ͣ�<INPUT name="m17" type="checkbox" value="1" <%IF strint(m_(17))=1 Then%>checked<%End IF%>>
			���ù̶���<INPUT name="m18" type="checkbox" value="1" <%IF strint(m_(18))=1 Then%>checked<%End IF%>>
			�Ƽ���־��<INPUT name="m19" type="checkbox" value="1" <%IF strint(m_(19))=1 Then%>checked<%End IF%>> 
           	ͼƬ��־��<INPUT name="m20" type="checkbox" value="1" <%IF strint(m_(20))=1 Then%>checked<%End IF%>>			
			������ʾ��<INPUT name="m21" type="checkbox" value="1" <%IF strint(m_(21))=1 Then%>checked<%End IF%>>
			��ʾ���ࣺ<INPUT name="m22" type="checkbox" value="1" <%IF strint(m_(22))=1 Then%>checked<%End IF%>>
			������ɫ��<INPUT name="m23" type="checkbox" value="1" <%IF strint(m_(23))=1 Then%>checked<%End IF%>>
			�ż㱳����<INPUT name="m24" type="checkbox" value="1" <%IF strint(m_(24))=1 Then%>checked<%End IF%>>
			���������<INPUT name="m25" type="checkbox" value="1" <%IF strint(m_(25))=1 Then%>checked<%End IF%>><br>
			��ʾ����(0-15)��<a href="#" ONCLICK="window.open('../Include/datetype.htm','Sample','toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=yes,resizable=no,copyhistory=no,width=200,height=400,left=300,top=20')">������</a>
			<INPUT name="m26" type="text" value="<%=m_(26)%>" size="5">
����������<INPUT name="m27" type="text" value="<%=m_(27)%>" size="5">
��ʾ������<INPUT name="m28" type="text" value="<%=m_(28)%>" size="5"> 
��ʾ������<INPUT name="m29" type="text" value="<%=m_(29)%>" size="6">
�����С��<INPUT name="m30" type="text" value="<%=m_(30)%>" size="5">px
 �оࣺ<INPUT name="m31" type="text" value="<%=m_(31)%>" size="5">px
	        </TD>
    </TR>
    <TR bgcolor="#FFFFFF">
      <TD height="26" align="right">�Ƽ���Ϣ��ʾ���ã�<%mb="A,B,C,D,E"
	  If instr(mb,M_logo)>0 Then
	  response.write "[��]"
	  Else
	  response.write "[<FONT color=#ff0000>��</font>]"
	  End if%></TD>
      <TD colspan="2">
	        ��Ϣ���ͣ�<INPUT name="m32" type="checkbox" value="1" <%IF strint(m_(32))=1 Then%>checked<%End IF%>>
			���ù̶���<INPUT name="m33" type="checkbox" value="1" <%IF strint(m_(33))=1 Then%>checked<%End IF%>>
			�Ƽ���־��<INPUT name="m34" type="checkbox" value="1" <%IF strint(m_(34))=1 Then%>checked<%End IF%>>
           	ͼƬ��־��<INPUT name="m35" type="checkbox" value="1" <%IF strint(m_(35))=1 Then%>checked<%End IF%>>			
			������ʾ��<INPUT name="m36" type="checkbox" value="1" <%IF strint(m_(36))=1 Then%>checked<%End IF%>>
			��ʾ���ࣺ<INPUT name="m37" type="checkbox" value="1" <%IF strint(m_(37))=1 Then%>checked<%End IF%>>
			������ɫ��<INPUT name="m38" type="checkbox" value="1" <%IF strint(m_(38))=1 Then%>checked<%End IF%>>
			�ż㱳����<INPUT name="m39" type="checkbox" value="1" <%IF strint(m_(39))=1 Then%>checked<%End IF%>>
			���������<INPUT name="m40" type="checkbox" value="1" <%IF strint(m_(40))=1 Then%>checked<%End IF%>><br>
			��ʾ����(0-15)��<a href="#" ONCLICK="window.open('../Include/datetype.htm','Sample','toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=yes,resizable=no,copyhistory=no,width=200,height=400,left=300,top=20')">������</a>
			<INPUT name="m41" type="text" value="<%=m_(41)%>" size="5">
����������<INPUT name="m42" type="text" value="<%=m_(42)%>" size="5">
��ʾ������<INPUT name="m43" type="text" value="<%=m_(43)%>" size="5"> 
��ʾ������<INPUT name="m44" type="text" value="<%=m_(44)%>" size="6">
�����С��<INPUT name="m45" type="text" value="<%=m_(45)%>" size="5">px
 �оࣺ<INPUT name="m46" type="text" value="<%=m_(46)%>" size="5">px
</TD>
    </TR>
    <TR bgcolor="#FFFFFF">
      <TD height="26" align="right" bgcolor="#BDDEEF">������Ϣ��ʾ���ã�<%mb="A,B,C,D,E"
	  If instr(mb,M_logo)>0 Then
	  response.write "[��]"
	  Else
	  response.write "[<FONT color=#ff0000>��</font>]"
	  End if%></TD>
      <TD colspan="2" bgcolor="#BDDEEF">
	  	    ��Ϣ���ͣ�<INPUT name="m47" type="checkbox" value="1" <%IF strint(m_(47))=1 Then%>checked<%End IF%>>
			���ù̶���<INPUT name="m48" type="checkbox" value="1" <%IF strint(m_(48))=1 Then%>checked<%End IF%>>
			�Ƽ���־��<INPUT name="m49" type="checkbox" value="1" <%IF strint(m_(49))=1 Then%>checked<%End IF%>> 
           	ͼƬ��־��<INPUT name="m50" type="checkbox" value="1" <%IF strint(m_(50))=1 Then%>checked<%End IF%>>			
			������ʾ��<INPUT name="m51" type="checkbox" value="1" <%IF strint(m_(51))=1 Then%>checked<%End IF%>>
			��ʾ���ࣺ<INPUT name="m52" type="checkbox" value="1" <%IF strint(m_(52))=1 Then%>checked<%End IF%>>
			������ɫ��<INPUT name="m53" type="checkbox" value="1" <%IF strint(m_(53))=1 Then%>checked<%End IF%>>
			�ż㱳����<INPUT name="m54" type="checkbox" value="1" <%IF strint(m_(54))=1 Then%>checked<%End IF%>>
			���������<INPUT name="m55" type="checkbox" value="1" <%IF strint(m_(55))=1 Then%>checked<%End IF%>><br>
			��ʾ����(0-15)��<a href="#" ONCLICK="window.open('../Include/datetype.htm','Sample','toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=yes,resizable=no,copyhistory=no,width=200,height=400,left=300,top=20')">������</a>
			<INPUT name="m56" type="text" value="<%=m_(56)%>" size="5">
����������<INPUT name="m57" type="text" value="<%=m_(57)%>" size="5">
��ʾ������<INPUT name="m58" type="text" value="<%=m_(58)%>" size="5"> 
��ʾ������<INPUT name="m59" type="text" value="<%=m_(59)%>" size="6">
�����С��<INPUT name="m60" type="text" value="<%=m_(60)%>" size="5">px
 �оࣺ<INPUT name="m61" type="text" value="<%=m_(61)%>" size="5">px</TD>
    </TR>
    <TR bgcolor="#FFFFFF"> 
      <TD height="26" align=right>��Ϣ�������б���ʾ���ã�<%mb="A,B,C,D,E"
	  If instr(mb,M_logo)>0 Then
	  response.write "[��]"
	  Else
	  response.write "[<FONT color=#ff0000>��</font>]"
	  End if%></TD>
      <TD width="73%" colspan="2" align=left>
	  	    ��Ϣ���ͣ�<INPUT name="m62" type="checkbox" value="1" <%IF strint(m_(62))=1 Then%>checked<%End IF%>>
			���������<INPUT name="m63" type="text" value="<%=m_(63)%>" size="5">
            ��ʾ������<INPUT name="m64" type="text" value="<%=m_(64)%>" size="5">
            ��ʾ������<INPUT name="m65" type="text" value="<%=m_(65)%>" size="5">
			���ⳤ�ȣ�<INPUT name="m66" type="text" value="<%=m_(66)%>" size="5">
			������ɫ��<INPUT name="m67" type="checkbox" value="1" <%IF strint(m_(67))=1 Then%>checked<%End IF%>>
			���ù̶���<INPUT name="m68" type="checkbox" value="1" <%IF strint(m_(68))=1 Then%>checked<%End IF%>>
	  	    ͼƬ��־��<INPUT name="m69" type="checkbox" value="1" <%IF strint(m_(69))=1 Then%>checked<%End IF%>>	  	    
	  	    �������<INPUT name="m70" type="checkbox" value="1" <%IF strint(m_(70))=1 Then%>checked<%End IF%>><br>			
	  	    ��ʾ����(0-15)��<a href="#" ONCLICK="window.open('../Include/datetype.htm','Sample','toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=yes,resizable=no,copyhistory=no,width=200,height=400,left=300,top=20')">������</a>
			<INPUT name="m71" type="text" value="<%=m_(71)%>" size="5">
			ͷ��ͼ�ģ�<INPUT name="m72" type="checkbox"  value="1" <%IF strint(m_(72))=1 Then%>checked<%End IF%>>	  
            �ż㱳����
            <SELECT name="m73">
              <OPTION value="0" <%IF StrInt(m_(73))=0 Then%>selected<%End IF%>>Ҫ�ż㱳��</OPTION>
              <OPTION value="1" <%IF StrInt(m_(73))=1 Then%>selected<%End IF%>>���ż㱣��ʽ</OPTION>
			  <OPTION value="2" <%IF StrInt(m_(73))=2 Then%>selected<%End IF%>>���ż��޸�ʽ</OPTION>
            </SELECT><br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<FONT color="#ff0000">�п�Ҫ�ڱ�ģ�巽����CSS������ҵ�shadow1���޸ĸö�width:��ֵ������ϢƵ��ҳ����</font></TD>
    </TR>

    <TR bgcolor="#FFFFFF">
      <TD height="26" align="right" bgcolor="#BDDEEF">������Ϣ������ʾ���ã�<%mb="A,B,C,D,E,F"
	  If instr(mb,M_logo)>0 Then
	  response.write "[��]"
	  Else
	  response.write "[<FONT color=#ff0000>��</font>]"
	  End if%></TD>
      <TD colspan="2" bgcolor="#BDDEEF">
	        ������ɫ��<INPUT name="m74" type="checkbox" value="1" <%IF strint(m_(74))=1 Then%>checked<%End IF%>>
			��Ϣ���ͣ�<INPUT name="m75" type="checkbox" value="1" <%IF strint(m_(75))=1 Then%>checked<%End IF%>>            
			��ϢLOGO��<INPUT name="m76" type="checkbox" value="1" <%IF strint(m_(76))=1 Then%>checked<%End IF%>>
           	���ù̶���<INPUT name="m77" type="checkbox" value="1" <%IF strint(m_(77))=1 Then%>checked<%End IF%>>
			��ʾ��Ч�ڣ�<INPUT name="m78" type="checkbox" value="1" <%IF strint(m_(78))=1 Then%>checked<%End IF%>>
			ͼƬ��־��<INPUT name="m79" type="checkbox" value="1" <%IF strint(m_(79))=1 Then%>checked<%End IF%>>
			���������<INPUT name="m80" type="checkbox" value="1" <%IF strint(m_(80))=1 Then%>checked<%End IF%>>
			��ʾ��Ա��<INPUT name="m81" type="checkbox" value="1" <%IF strint(m_(81))=1 Then%>checked<%End IF%>>
			��ʾ���ۣ�<INPUT name="m82" type="checkbox" value="1" <%IF strint(m_(82))=1 Then%>checked<%End IF%>><br>

���ⳤ�ȣ�<INPUT name="m83" type="text" value="<%=m_(83)%>" size="5">
��鳤�ȣ�<INPUT name="m84" type="text" value="<%=m_(84)%>" size="5"> 
��ʾ������<INPUT name="m85" type="text" value="<%=m_(85)%>" size="6">
��ʾ������<INPUT name="m86" type="text" value="<%=m_(86)%>" size="5">
			��Ϣ��Χ��
            <SELECT name="m87">
              <OPTION value="0" <%IF strint(m_(87))=0 Then%>selected<%End IF%>>��ʾȫ����Ϣ</OPTION>
              <OPTION value="1" <%IF strint(m_(87))=1 Then%>selected<%End IF%>>����ʾ������Ϣ</OPTION>			  
            </SELECT>
</TD>
    </TR>

    <TR bgcolor="#FFFFFF">
      <TD height="26" align="right">ͼƬ��Ϣ�Ƽ���ʾ��<%mb=""
	  If instr(mb,M_logo)>0 Then
	  response.write "[��]"
	  Else
	  response.write "[<FONT color=#ff0000>��</font>]"
	  End if%></TD>
      <TD colspan="2">
ÿ����ʾ������<INPUT name="m88" type="text" value="<%=m_(88)%>" size="5">
�ܹ���ʾ������<INPUT name="m89" type="text" value="<%=m_(89)%>" size="5">
		</TD>
    </TR>
    <TR bgcolor="#FFFFFF">
      <TD height="26" align="right" bgcolor="#BDDEEF">������Ϣͼ����ʾ���ã�<%mb="F"
	  If instr(mb,M_logo)>0 Then
	  response.write "[��]"
	  Else
	  response.write "[<FONT color=#ff0000>��</font>]"
	  End if%></TD>
      <TD colspan="2" bgcolor="#BDDEEF">
	        ��Ϣ���ͣ�<INPUT name="m90" type="checkbox" value="1" <%IF strint(m_(90))=1 Then%>checked<%End IF%>>
			���ù̶���<INPUT name="m91" type="checkbox" value="1" <%IF strint(m_(91))=1 Then%>checked<%End IF%>>
			�Ƽ���־��<INPUT name="m92" type="checkbox"  value="1" <%IF strint(m_(92))=1 Then%>checked<%End IF%>>
           	ͼƬ��־��<INPUT name="m93" type="checkbox"  value="1" <%IF strint(m_(93))=1 Then%>checked<%End IF%>>						
			���������<INPUT name="m94" type="checkbox"  value="1" <%IF strint(m_(94))=1 Then%>checked<%End IF%>>
			��ʾ����(0-15)��<a href="#" ONCLICK="window.open('../Include/datetype.htm','Sample','toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=yes,resizable=no,copyhistory=no,width=200,height=400,left=300,top=20')">������</a>
			<INPUT name="m95" type="text" value="<%=m_(95)%>" size="5">
�г��۸�<INPUT name="m96" type="checkbox"  value="1" <%IF strint(m_(96))=1 Then%>checked<%End IF%>>
��վ�۸�<INPUT name="m97" type="checkbox"  value="1" <%IF strint(m_(97))=1 Then%>checked<%End IF%>><br>
���ⳤ�ȣ�<INPUT name="m98" type="text" value="<%=m_(98)%>" size="5">
������ɫ��<INPUT name="m99" type="checkbox" value="1" <%IF strint(m_(99))=1 Then%>checked<%End IF%>>
�����С��<INPUT name="m100" type="text" value="<%=m_(100)%>" size="5">px
ͼƬ��ȣ�<INPUT name="m101" type="text" value="<%=m_(101)%>" size="6">
ͼƬ�߶ȣ�<INPUT name="m102" type="text" value="<%=m_(102)%>" size="6">
��Ԫ��ȣ�<INPUT name="m103" type="text" value="<%=m_(103)%>" size="6">
��Ԫ�߶ȣ�<INPUT name="m104" type="text" value="<%=m_(104)%>" size="6"><br>
ÿҳ������<INPUT name="m105" type="text" value="<%=m_(105)%>" size="5"> 
��ʾ������<INPUT name="m106" type="text" value="<%=m_(106)%>" size="6">
��ʾ��ҳ��<INPUT name="m107" type="checkbox" value="1" <%IF strint(m_(107))=1 Then%>checked<%End IF%>>

    </TR>

    <TR bgcolor="#ffffff">
      <TD height="26" align="right">�Ƽ���Ϣͼ����ʾ���ã�<%mb="F"
	  If instr(mb,M_logo)>0 Then
	  response.write "[��]"
	  Else
	  response.write "[<FONT color=#ff0000>��</font>]"
	  End if%></TD>
      <TD colspan="2">
	        ��Ϣ���ͣ�<INPUT name="m108" type="checkbox" value="1" <%IF strint(m_(108))=1 Then%>checked<%End IF%>>
			���ù̶���<INPUT name="m109" type="checkbox" value="1" <%IF strint(m_(109))=1 Then%>checked<%End IF%>>
			�Ƽ���־��<INPUT name="m110" type="checkbox"  value="1" <%IF strint(m_(110))=1 Then%>checked<%End IF%>>
           	ͼƬ��־��<INPUT name="m111" type="checkbox"  value="1" <%IF strint(m_(111))=1 Then%>checked<%End IF%>>						
			���������<INPUT name="m112" type="checkbox"  value="1" <%IF strint(m_(112))=1 Then%>checked<%End IF%>>
			��ʾ����(0-15)��<a href="#" ONCLICK="window.open('../Include/datetype.htm','Sample','toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=yes,resizable=no,copyhistory=no,width=200,height=400,left=300,top=20')">������</a>
			<INPUT name="m113" type="text" value="<%=m_(113)%>" size="5">
�г��۸�<INPUT name="m114" type="checkbox"  value="1" <%IF strint(m_(114))=1 Then%>checked<%End IF%>>
��վ�۸�<INPUT name="m115" type="checkbox"  value="1" <%IF strint(m_(115))=1 Then%>checked<%End IF%>><br>
���ⳤ�ȣ�<INPUT name="m116" type="text" value="<%=m_(116)%>" size="5">
������ɫ��<INPUT name="m117" type="checkbox" value="1" <%IF strint(m_(117))=1 Then%>checked<%End IF%>>
�����С��<INPUT name="m118" type="text" value="<%=m_(118)%>" size="5">px
ͼƬ��ȣ�<INPUT name="m119" type="text" value="<%=m_(119)%>" size="6">
ͼƬ�߶ȣ�<INPUT name="m120" type="text" value="<%=m_(120)%>" size="6">
��Ԫ��ȣ�<INPUT name="m121" type="text" value="<%=m_(121)%>" size="6">
��Ԫ�߶ȣ�<INPUT name="m122" type="text" value="<%=m_(122)%>" size="6"><br>
ÿҳ������<INPUT name="m123" type="text" value="<%=m_(123)%>" size="5"> 
��ʾ������<INPUT name="m124" type="text" value="<%=m_(124)%>" size="6">
��ʾ��ҳ��<INPUT name="m125" type="checkbox" value="1" <%IF strint(m_(125))=1 Then%>checked<%End IF%>>

    </TR>
	
	    <TR  bgcolor="#BDDEEF">
      <TD height="26" align="right">������Ϣͼ����ʾ���ã�<%mb="F"
	  If instr(mb,M_logo)>0 Then
	  response.write "[��]"
	  Else
	  response.write "[<FONT color=#ff0000>��</font>]"
	  End if%></TD>
      <TD colspan="2">
	        ��Ϣ���ͣ�<INPUT name="m126" type="checkbox" value="1" <%IF strint(m_(126))=1 Then%>checked<%End IF%>>
			���ù̶���<INPUT name="m127" type="checkbox" value="1" <%IF strint(m_(127))=1 Then%>checked<%End IF%>>
			�Ƽ���־��<INPUT name="m128" type="checkbox"  value="1" <%IF strint(m_(128))=1 Then%>checked<%End IF%>>
           	ͼƬ��־��<INPUT name="m129" type="checkbox"  value="1" <%IF strint(m_(129))=1 Then%>checked<%End IF%>>						
			���������<INPUT name="m130" type="checkbox"  value="1" <%IF strint(m_(130))=1 Then%>checked<%End IF%>>
			��ʾ����(0-15)��<a href="#" ONCLICK="window.open('../Include/datetype.htm','Sample','toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=yes,resizable=no,copyhistory=no,width=200,height=400,left=300,top=20')">������</a>
			<INPUT name="m131" type="text" value="<%=m_(131)%>" size="5">
�г��۸�<INPUT name="m132" type="checkbox"  value="1" <%IF strint(m_(132))=1 Then%>checked<%End IF%>>
��վ�۸�<INPUT name="m133" type="checkbox"  value="1" <%IF strint(m_(133))=1 Then%>checked<%End IF%>><br>
���ⳤ�ȣ�<INPUT name="m134" type="text" value="<%=m_(134)%>" size="5">
������ɫ��<INPUT name="m135" type="checkbox" value="1" <%IF strint(m_(135))=1 Then%>checked<%End IF%>>
�����С��<INPUT name="m136" type="text" value="<%=m_(136)%>" size="5">px
ͼƬ��ȣ�<INPUT name="m137" type="text" value="<%=m_(137)%>" size="6">
ͼƬ�߶ȣ�<INPUT name="m138" type="text" value="<%=m_(138)%>" size="6">
��Ԫ��ȣ�<INPUT name="m139" type="text" value="<%=m_(139)%>" size="6">
��Ԫ�߶ȣ�<INPUT name="m140" type="text" value="<%=m_(140)%>" size="6"><br>
ÿҳ������<INPUT name="m141" type="text" value="<%=m_(141)%>" size="5"> 
��ʾ������<INPUT name="m142" type="text" value="<%=m_(142)%>" size="6">
��ʾ��ҳ��<INPUT name="m143" type="checkbox" value="1" <%IF strint(m_(143))=1 Then%>checked<%End IF%>>

    </TR>
    <TR bgcolor="#FFFFFF">
      <TD height="26" align="right">������ʾ���ã�<%mb="A,B,C,D,E,F"
	  If instr(mb,M_logo)>0 Then
	  response.write "[��]"
	  Else
	  response.write "[<FONT color=#ff0000>��</font>]"
	  End if%></TD>
      <TD colspan="2">
��ʾ������<INPUT name="m144" type="text" value="<%=m_(144)%>" size="5">
��ʾ������<INPUT name="m145" type="text" value="<%=m_(145)%>" size="5">
���ⳤ�ȣ�<INPUT name="m146" type="text" value="<%=m_(146)%>" size="5"> 
�����иߣ�<INPUT name="m147" type="text" value="<%=m_(147)%>" size="5">
�����ֺţ�<INPUT name="m148" type="text" value="<%=m_(148)%>" size="5">
���ù̶���<INPUT name="m149" type="checkbox" value="1" <%IF strint(m_(149))=1 Then%>checked<%End IF%>><br>
����VIP��Ա�ģ�<INPUT name="m150" type="checkbox" value="1" <%IF strint(m_(150))=1 Then%>checked<%End IF%>>
ͼ����ʾ��
<SELECT name="m151">
<OPTION value="0" <%IF strint(m_(151))=0 Then%>selected<%End IF%>>����</OPTION>
<OPTION value="1" <%IF strint(m_(151))=1 Then%>selected<%End IF%>>ͼ��</OPTION>
</SELECT>
ͼƬ��ȣ�<INPUT name="m152" type="text" value="<%=m_(152)%>" size="5"> 
ͼƬ�иߣ�<INPUT name="m153" type="text" value="<%=m_(153)%>" size="5">
</TD>
    </TR>
	   <TR  bgcolor="#BDDEEF">
      <TD height="26" align="right">������Դ������ʾ���ã�<%mb="A,B,C,D,E,F"
	  If instr(mb,M_logo)>0 Then
	  response.write "[��]"
	  Else
	  response.write "[<FONT color=#ff0000>��</font>]"
	  End if%></TD>
      <TD colspan="2">
            ��ʾ��Ŀ��<INPUT name="m154" type="text" value="<%=m_(154)%>" size="5">������ĿID�ţ�0Ϊȫ����
			��ʾ���ͣ�
            <SELECT name="m155">
              <OPTION value="0" <%IF strint(m_(155))=0 Then%>selected<%End IF%>>ȫ��</OPTION>
              <OPTION value="1" <%IF strint(m_(155))=1 Then%>selected<%End IF%>>�Ƽ�</OPTION>			  
			  <OPTION value="2" <%IF strint(m_(155))=2 Then%>selected<%End IF%>>����</OPTION>			  
            </SELECT>
	  	    ���ù̶���<INPUT name="m156" type="checkbox" value="1" <%IF strint(m_(156))=1 Then%>checked<%End IF%>>
	  	    �Ƽ���־��<INPUT name="m157" type="checkbox" value="1" <%IF strint(m_(157))=1 Then%>checked<%End IF%>>
	  	    ͼƬ��־��<INPUT name="m158" type="checkbox" value="1" <%IF strint(m_(158))=1 Then%>checked<%End IF%>>
	  	    �������<INPUT name="m159" type="checkbox" value="1" <%IF strint(m_(159))=1 Then%>checked<%End IF%>>
			��ʾ����(0-15)��<a href="#" ONCLICK="window.open('../Include/datetype.htm','Sample','toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=yes,resizable=no,copyhistory=no,width=200,height=400,left=300,top=20')">������</a>
			<INPUT name="m160" type="text" value="<%=m_(160)%>" size="5"><br>
			���ⳤ�ȣ�<INPUT name="m161" type="text" value="<%=m_(161)%>" size="5">
			��ʾ������<INPUT name="m162" type="text" value="<%=m_(162)%>" size="5">
			��ʾ������<INPUT name="m163" type="text" value="<%=m_(163)%>" size="5">
			�����ֺţ�<INPUT name="m164" type="text" value="<%=m_(164)%>" size="5">
			�����иߣ�<INPUT name="m165" type="text" value="<%=m_(165)%>" size="5">
            ��ʾ���ࣺ<INPUT name="m166" type="checkbox" value="1" <%IF strint(m_(166))=1 Then%>checked<%End IF%>>

</TD>
    </TR>
    <TR bgcolor="#FFFFFF">
      <TD height="26" align="right">��ҵ��Ƭ(����)��ʾ���ã�<%mb="A,B,C,D,E"
	  If instr(mb,M_logo)>0 Then
	  response.write "[��]"
	  Else
	  response.write "[<FONT color=#ff0000>��</font>]"
	  End if%></TD>
      <TD colspan="2">
��ʾ������<INPUT name="m167" type="text" value="<%=m_(167)%>" size="5">
��ʾ������<INPUT name="m168" type="text" value="<%=m_(168)%>" size="5"> 
���ⳤ�ȣ�<INPUT name="m169" type="text" value="<%=m_(169)%>" size="5">
�����иߣ�<INPUT name="m170" type="text" value="<%=m_(170)%>" size="5">
��Ӫ���ȣ�<INPUT name="m171" type="text" value="<%=m_(171)%>" size="5">
��ַ���ȣ�<INPUT name="m172" type="text" value="<%=m_(172)%>" size="5">
����ȣ�<INPUT name="m173" type="text" value="<%=m_(173)%>" size="5"><br>
����Ӵ֣�<INPUT name="m174" type="checkbox" value="1" <%IF strint(m_(174))=1 Then%>checked<%End IF%>>
�ż㱳����
            <SELECT name="m175">
              <OPTION value="0" <%IF strint(m_(175))=0 Then%>selected<%End IF%>>Ҫ�ż㱳��</OPTION>
              <OPTION value="1" <%IF strint(m_(175))=1 Then%>selected<%End IF%>>��Ҫ�ż㱣��ʽ</OPTION>
			  <OPTION value="2" <%IF strint(m_(175))=2 Then%>selected<%End IF%>>��Ҫ�ż��޸�ʽ</OPTION>			 
            </SELECT>
������ɫ��<INPUT name="m176" type="text" style="background:<%=m_(176)%>" onClick="Getcolor(this)" value="<%=m_(176)%>" size="8" maxlength="7">
��ʾ��Ӫ��ַ�绰��<INPUT name="m177" type="checkbox" value="1" <%IF strint(m_(177))=1 Then%>checked<%End IF%>>
��ʾ���ࣺ<INPUT name="m178" type="checkbox" value="1" <%IF strint(m_(178))=1 Then%>checked<%End IF%>>

</TD>
    </TR>

    <TR bgcolor="#FFFFFF">
      <TD height="26" align="right" bgcolor="#BDDEEF">��ҵ��Ƭ(ͼƬ)��ʾ���ã�<%mb="A,B,C,D,E,F"
	  If instr(mb,M_logo)>0 Then
	  response.write "[��]"
	  Else
	  response.write "[<FONT color=#ff0000>��</font>]"
	  End if%></TD>
      <TD colspan="2" bgcolor="#BDDEEF" >
ÿ��������<INPUT name="m179" type="text" value="<%=m_(179)%>" size="5">
��ʾ������<INPUT name="m180" type="text" value="<%=m_(180)%>" size="5">
���ⳤ�ȣ�<INPUT name="m181" type="text" value="<%=m_(181)%>" size="5"> 
��Ӫ���ȣ�<INPUT name="m182" type="text" value="<%=m_(182)%>" size="5">

		</TD>
    </TR>
    <TR bgcolor="#FFFFFF">
      <TD height="26" align="right" >������Ѷһ��ʾ���ã�<%mb="A,B,C,D,E,F"
	  If instr(mb,M_logo)>0 Then
	  response.write "[��]"
	  Else
	  response.write "[<FONT color=#ff0000>��</font>]"
	  End if%><br>��Ĭ���������Ŷ�̬��&nbsp;&nbsp;&nbsp;&nbsp;</TD>
      <TD colspan="2">
			��ʾ���ͣ�
            <SELECT name="m183">
              <OPTION value="0" <%IF strint(m_(183))=0 Then%>selected<%End IF%>>ȫ��</OPTION>
              <OPTION value="1" <%IF strint(m_(183))=1 Then%>selected<%End IF%>>�Ƽ�</OPTION>
			  <OPTION value="2" <%IF strint(m_(183))=2 Then%>selected<%End IF%>>ͼ��</OPTION>
			  <OPTION value="3" <%IF strint(m_(183))=3 Then%>selected<%End IF%>>����</OPTION>			  
            </SELECT>
            һ����ĿID�ţ�<INPUT name="m184" type="text" value="<%=m_(184)%>" size="3">
			������ĿID�ţ�<INPUT name="m185" type="text" value="<%=m_(185)%>" size="3">
			������ĿID�ţ�<INPUT name="m186" type="text" value="<%=m_(186)%>" size="3"><FONT color="#ff0000">������ID����������Ѷ����������Ŀ�����л�ã��������Ϊ0ʱ��ʾȫ����</font><br>
	  	    ���ù̶���<INPUT name="m187" type="checkbox" value="1" <%IF strint(m_(187))=1 Then%>checked<%End IF%>>
	  	    �Ƽ���־��<INPUT name="m188" type="checkbox" value="1" <%IF strint(m_(188))=1 Then%>checked<%End IF%>>
	  	    ͼƬ��־��<INPUT name="m189" type="checkbox" value="1" <%IF strint(m_(189))=1 Then%>checked<%End IF%>>
	  	    �������<INPUT name="m190" type="checkbox" value="1" <%IF strint(m_(190))=1 Then%>checked<%End IF%>>
	  	    ������ʾ��<INPUT name="m191" type="checkbox" value="1" <%IF strint(m_(191))=1 Then%>checked<%End IF%>>
			��ʾ����(0-15)��<a href="#" ONCLICK="window.open('../Include/datetype.htm','Sample','toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=yes,resizable=no,copyhistory=no,width=200,height=400,left=300,top=20')">������</a>
			<INPUT name="m192" type="text" value="<%=m_(192)%>" size="3">			
			��ʾ������<INPUT name="m193" type="text" value="<%=m_(193)%>" size="3">
			��ʾ������<INPUT name="m194" type="text" value="<%=m_(194)%>" size="3"><br>
			���ⳤ�ȣ�<INPUT name="m195" type="text" value="<%=m_(195)%>" size="3">
			�����ֺţ�<INPUT name="m196" type="text" value="<%=m_(196)%>" size="3">
			�����иߣ�<INPUT name="m197" type="text" value="<%=m_(197)%>" size="3">

</TD>
    </TR>
	   <TR bgcolor="#FFFFFF">
      <TD height="26" align="right" bgcolor="#BDDEEF">������Ѷ����ʾ���ã�<%mb="A,B,C,D,E"
	  If instr(mb,M_logo)>0 Then
	  response.write "[��]"
	  Else
	  response.write "[<FONT color=#ff0000>��</font>]"
	  End if%><br>��Ĭ�����ڷ���ָ�ϣ�&nbsp;&nbsp;&nbsp;&nbsp;</TD>
      <TD colspan="2" bgcolor="#BDDEEF">
			��ʾ���ͣ�
            <SELECT name="m198">
              <OPTION value="0" <%IF strint(m_(198))=0 Then%>selected<%End IF%>>ȫ��</OPTION>
              <OPTION value="1" <%IF strint(m_(198))=1 Then%>selected<%End IF%>>�Ƽ�</OPTION>
			  <OPTION value="2" <%IF strint(m_(198))=2 Then%>selected<%End IF%>>ͼ��</OPTION>
			  <OPTION value="3" <%IF strint(m_(198))=3 Then%>selected<%End IF%>>����</OPTION>			  
            </SELECT>
            һ����ĿID�ţ�<INPUT name="m199" type="text" value="<%=m_(199)%>" size="3">
			������ĿID�ţ�<INPUT name="m200" type="text" value="<%=m_(200)%>" size="3">
			������ĿID�ţ�<INPUT name="m201" type="text" value="<%=m_(201)%>" size="3"><FONT color="#ff0000">������ID����������Ѷ����������Ŀ�����л�ã��������Ϊ0ʱ��ʾȫ����</font><br>
	  	    ���ù̶���<INPUT name="m202" type="checkbox" value="1" <%IF strint(m_(202))=1 Then%>checked<%End IF%>>
	  	    �Ƽ���־��<INPUT name="m203" type="checkbox" value="1" <%IF strint(m_(203))=1 Then%>checked<%End IF%>>
	  	    ͼƬ��־��<INPUT name="m204" type="checkbox" value="1" <%IF strint(m_(204))=1 Then%>checked<%End IF%>>
	  	    �������<INPUT name="m205" type="checkbox" value="1" <%IF strint(m_(205))=1 Then%>checked<%End IF%>>
	  	    ������ʾ��<INPUT name="m206" type="checkbox" value="1" <%IF strint(m_(206))=1 Then%>checked<%End IF%>>
			��ʾ����(0-15)��<a href="#" ONCLICK="window.open('../Include/datetype.htm','Sample','toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=yes,resizable=no,copyhistory=no,width=200,height=400,left=300,top=20')">������</a>
			<INPUT name="m207" type="text" value="<%=m_(207)%>" size="3">
			��ʾ������<INPUT name="m208" type="text" value="<%=m_(208)%>" size="3">
			��ʾ������<INPUT name="m209" type="text" value="<%=m_(209)%>" size="3"><br>
			���ⳤ�ȣ�<INPUT name="m210" type="text" value="<%=m_(210)%>" size="3">
			�����ֺţ�<INPUT name="m211" type="text" value="<%=m_(211)%>" size="3">			
			�����иߣ�<INPUT name="m212" type="text" value="<%=m_(212)%>" size="3">

</TD>
    </TR>
   <TR bgcolor="#FFFFFF">
      <TD height="26" align="right">������Ѷ����ʾ���ã�<%mb="A,B,C,D,E"
	  If instr(mb,M_logo)>0 Then
	  response.write "[��]"
	  Else
	  response.write "[<FONT color=#ff0000>��</font>]"
	  End if%></TD>
      <TD colspan="2">
			��ʾ���ͣ�
            <SELECT name="m213">
              <OPTION value="0" <%IF strint(m_(213))=0 Then%>selected<%End IF%>>ȫ��</OPTION>
              <OPTION value="1" <%IF strint(m_(213))=1 Then%>selected<%End IF%>>�Ƽ�</OPTION>
			  <OPTION value="2" <%IF strint(m_(213))=2 Then%>selected<%End IF%>>ͼ��</OPTION>
			  <OPTION value="3" <%IF strint(m_(213))=3 Then%>selected<%End IF%>>����</OPTION>			  
            </SELECT>
            һ����ĿID�ţ�<INPUT name="m214" type="text" value="<%=m_(214)%>" size="3">
			������ĿID�ţ�<INPUT name="m215" type="text" value="<%=m_(215)%>" size="3">
			������ĿID�ţ�<INPUT name="m216" type="text" value="<%=m_(216)%>" size="3"><FONT color="#ff0000">������ID����������Ѷ����������Ŀ�����л�ã��������Ϊ0ʱ��ʾȫ����</font><br>
	  	    ���ù̶���<INPUT name="m217" type="checkbox" value="1" <%IF strint(m_(217))=1 Then%>checked<%End IF%>>
	  	    �Ƽ���־��<INPUT name="m218" type="checkbox" value="1" <%IF strint(m_(218))=1 Then%>checked<%End IF%>>
	  	    ͼƬ��־��<INPUT name="m219" type="checkbox" value="1" <%IF strint(m_(219))=1 Then%>checked<%End IF%>>
	  	    �������<INPUT name="m220" type="checkbox" value="1" <%IF strint(m_(220))=1 Then%>checked<%End IF%>>
	  	    ������ʾ��<INPUT name="m221" type="checkbox" value="1" <%IF strint(m_(221))=1 Then%>checked<%End IF%>>
			��ʾ����(0-15)��<a href="#" ONCLICK="window.open('../Include/datetype.htm','Sample','toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=yes,resizable=no,copyhistory=no,width=200,height=400,left=300,top=20')">������</a>
			<INPUT name="m222" type="text" value="<%=m_(222)%>" size="3">
			��ʾ������<INPUT name="m223" type="text" value="<%=m_(223)%>" size="3">
			��ʾ������<INPUT name="m224" type="text" value="<%=m_(224)%>" size="3"><br>
			���ⳤ�ȣ�<INPUT name="m225" type="text" value="<%=m_(225)%>" size="3">
			�����ֺţ�<INPUT name="m226" type="text" value="<%=m_(226)%>" size="3">
			�����иߣ�<INPUT name="m227" type="text" value="<%=m_(227)%>" size="3">

</TD>
    </TR>
    <TR bgcolor="#FFFFFF">
      <TD height="26" align="right" bgcolor="#BDDEEF">������Ѷ����ʾ���ã�<%mb="A,B,C,D,E"
	  If instr(mb,M_logo)>0 Then
	  response.write "[��]"
	  Else
	  response.write "[<FONT color=#ff0000>��</font>]"
	  End if%></TD>
      <TD colspan="2" bgcolor="#BDDEEF">
	  		��ʾ���ͣ�
            <SELECT name="m228">
              <OPTION value="0" <%IF strint(m_(228))=0 Then%>selected<%End IF%>>ȫ��</OPTION>
              <OPTION value="1" <%IF strint(m_(228))=1 Then%>selected<%End IF%>>�Ƽ�</OPTION>
			  <OPTION value="2" <%IF strint(m_(228))=2 Then%>selected<%End IF%>>ͼ��</OPTION>
			  <OPTION value="3" <%IF strint(m_(228))=3 Then%>selected<%End IF%>>����</OPTION>			  
            </SELECT>
            һ����ĿID�ţ�<INPUT name="m229" type="text" value="<%=m_(229)%>" size="3">
			������ĿID�ţ�<INPUT name="m230" type="text" value="<%=m_(230)%>" size="3">
			������ĿID�ţ�<INPUT name="m231" type="text" value="<%=m_(231)%>" size="3"><FONT color="#ff0000">������ID����������Ѷ����������Ŀ�����л�ã��������Ϊ0ʱ��ʾȫ����</font><br>
	  	    ���ù̶���<INPUT name="m232" type="checkbox" value="1" <%IF strint(m_(232))=1 Then%>checked<%End IF%>>
	  	    �Ƽ���־��<INPUT name="m233" type="checkbox" value="1" <%IF strint(m_(233))=1 Then%>checked<%End IF%>>
	  	    ͼƬ��־��<INPUT name="m234" type="checkbox" value="1" <%IF strint(m_(234))=1 Then%>checked<%End IF%>>
	  	    �������<INPUT name="m235" type="checkbox" value="1" <%IF strint(m_(235))=1 Then%>checked<%End IF%>>
	  	    ������ʾ��<INPUT name="m236" type="checkbox" value="1" <%IF strint(m_(236))=1 Then%>checked<%End IF%>>
			��ʾ����(0-15)��<a href="#" ONCLICK="window.open('../Include/datetype.htm','Sample','toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=yes,resizable=no,copyhistory=no,width=200,height=400,left=300,top=20')">������</a>
			<INPUT name="m237" type="text" value="<%=m_(237)%>" size="3">
			��ʾ������<INPUT name="m238" type="text" value="<%=m_(238)%>" size="3">
			��ʾ������<INPUT name="m239" type="text" value="<%=m_(239)%>" size="3"><br>			
			���ⳤ�ȣ�<INPUT name="m240" type="text" value="<%=m_(240)%>" size="3">
			�����ֺţ�<INPUT name="m241" type="text" value="<%=m_(241)%>" size="3">
			�����иߣ�<INPUT name="m242" type="text" value="<%=m_(242)%>" size="3">

</TD>
    </TR>	
	<TR bgcolor="#FFFFFF">
      <TD height="26" align="right">������Ѷ�����ʾ���ã�<%mb="A,B,C,D,E"
	  If instr(mb,M_logo)>0 Then
	  response.write "[��]"
	  Else
	  response.write "[<FONT color=#ff0000>��</font>]"
	  End if%></TD>
      <TD colspan="2">
            һ����ĿID�ţ�<INPUT name="m243" type="text" value="<%=m_(243)%>" size="3">
			������ĿID�ţ�<INPUT name="m244" type="text" value="<%=m_(244)%>" size="3">
			������ĿID�ţ�<INPUT name="m245" type="text" value="<%=m_(245)%>" size="3"><FONT color="#ff0000">����Ҫ������ID����������Ѷ����������Ŀ�����л�ã��������Ϊ0ʱ��ʾȫ����</font><br>
	  	    ���ù̶���<INPUT name="m246" type="checkbox" value="1" <%IF strint(m_(246))=1 Then%>checked<%End IF%>>
			��ʾ�Ƽ���<INPUT name="m247" type="checkbox" value="1" <%IF strint(m_(247))=1 Then%>checked<%End IF%>>
	  	    ��ͼ������ʾ��<INPUT name="m248" type="checkbox" value="1" <%IF strint(m_(248))=1 Then%>checked<%End IF%>>
            ������ʾ��<INPUT name="m249" type="checkbox" value="1" <%IF strint(m_(249))=1 Then%>checked<%End IF%>>
			���ⳤ�ȣ�<INPUT name="m250" type="text" value="<%=m_(250)%>" size="3">
			ͼƬ��ȣ�<INPUT name="m251" type="text" value="<%=m_(251)%>" size="3">
			ͼƬ�߶ȣ�<INPUT name="m252" type="text" value="<%=m_(252)%>" size="3"><br>
			ÿ��������<INPUT name="m253" type="text" value="<%=m_(253)%>" size="3">
			��ʾ������<INPUT name="m254" type="text" value="<%=m_(254)%>" size="3">
			��ʾ��ҳ��<INPUT name="m255" type="checkbox" value="1" <%IF strint(m_(255))=1 Then%>checked<%End IF%>>

</TD>
    </TR>
<TR  bgcolor="#BDDEEF">
      <TD height="26" align="right" >������־���ã�<%mb="A,B,C,D,E"
	  If instr(mb,M_logo)>0 Then
	  response.write "[��]"
	  Else
	  response.write "[<FONT color=#ff0000>��</font>]"
	  End if%> </TD>
      <TD colspan="2">
ÿ����ʾ������<INPUT name="m256" type="text" value="<%=m_(256)%>" size="5">
�ܹ���ʾ������<INPUT name="m257" type="text" value="<%=m_(257)%>" size="5">
ͼƬ��ȣ�<INPUT name="m258" type="text" value="<%=m_(258)%>" size="5">
ͼƬ�߶ȣ�<INPUT name="m259" type="text" value="<%=m_(259)%>" size="5">
���ⳤ�ȣ�<INPUT name="m260" type="text" value="<%=m_(260)%>" size="5">
		</TD>
    </TR>
<TR bgcolor="#FFFFFF">
      <TD height="26" align="right" >����114��ʾ���ã�<%mb="A,B,C,D,E,F"
	  If instr(mb,M_logo)>0 Then
	  response.write "[��]"
	  Else
	  response.write "[<FONT color=#ff0000>��</font>]"
	  End if%> </TD>
      <TD colspan="2">
ÿ����ʾ������<INPUT name="m261" type="text" value="<%=m_(261)%>" size="5">
�ܹ���ʾ������<INPUT name="m262" type="text" value="<%=m_(262)%>" size="5">
		(<FONT color="#ff0000">�п�Ҫ�ڱ�ģ�巽����CSS������ҵ�.articleline114sy��޸ĸö�width:��ֵ</font>)</TD>
    </TR>
	<TR bgcolor="#FFFFFF"> 
      <TD height="26" align="right" bgcolor="#BDDEEF">���������ʾ���ã�<%mb="A,B,C,D,E"
	  If instr(mb,M_logo)>0 Then
	  response.write "[��]"
	  Else
	  response.write "[<FONT color=#ff0000>��</font>]"
	  End if%></TD>
      <TD colspan="2"  bgcolor="#BDDEEF">
	  	    ��ʾ������<INPUT name="m263" type="text" value="<%=m_(263)%>" size="5">
	  	    ÿ�и�����<INPUT name="m264" type="text" value="<%=m_(264)%>" size="5">
			С���ȣ�<INPUT name="m265" type="text" value="<%=m_(265)%>" size="5">
	  	    �иߣ�<INPUT name="m266" type="text" value="<%=m_(266)%>" size="5">
			�ֺţ�<INPUT name="m267" type="text" value="<%=m_(267)%>" size="5">
	  	    ���ⳤ�ȣ�<INPUT name="m268" type="text" value="<%=m_(268)%>" size="5"><br>
	  	    ����ɫ��<INPUT name="m269" type="checkbox" value="1" <%IF strint(m_(269))=1 Then%>checked<%End IF%>>
			�Ӵ֣�<INPUT name="m270" type="checkbox" value="1" <%IF strint(m_(270))=1 Then%>checked<%End IF%>>
            �»��ߣ�<INPUT name="m271" type="checkbox" value="1" <%IF strint(m_(271))=1 Then%>checked<%End IF%>>
            ��ʾ���ࣺ<INPUT name="m272" type="checkbox" value="1" <%IF strint(m_(272))=1 Then%>checked<%End IF%>>
			</TD>
    </TR>

	<TR bgcolor="#FFFFFF"> 
      <TD height="26" align="right">����������ʾ���ã�<%mb="A,B,C,D,E,F"
	  If instr(mb,M_logo)>0 Then
	  response.write "[��]"
	  Else
	  response.write "[<FONT color=#ff0000>��</font>]"
	  End if%></TD>
      <TD colspan="2">
            ��ʾ��ʽ��
            <SELECT name="m273">
              <OPTION value="0" <%IF StrInt(m_(273))=0 Then%>selected<%End IF%>>ͼ�ķ���</OPTION>
              <OPTION value="1" <%IF StrInt(m_(273))=1 Then%>selected<%End IF%>>�ۺ���ʾ</OPTION>
              <OPTION value="2" <%IF StrInt(m_(273))=2 Then%>selected<%End IF%>>��������</OPTION>  
            </SELECT>
			վ�����ȣ�<INPUT name="m274" type="text" value="<%=m_(274)%>" size="5">
	  	    ÿ�и߶ȣ�<INPUT name="m275" type="text" value="<%=m_(275)%>" size="5">			
            ÿ��������<INPUT name="m276" type="text" value="<%=m_(276)%>" size="5"><br>
			�ۺ���ʾ������<INPUT name="m277" type="text" value="<%=m_(277)%>" size="5">
			�ۺ���ʾLOGOռ������<INPUT name="m278" type="text" value="<%=m_(278)%>" size="5">			
			LOGO����������<INPUT name="m279" type="text" value="<%=m_(279)%>" size="5">
			���ֶ���������<INPUT name="m280" type="text" value="<%=m_(280)%>" size="5">	
			</TD>
    </TR>

	<TR  bgcolor="#BDDEEF"> 
      <TD height="26" align="right">�ͻ�������ʾ���ã�<%mb=""
	  If instr(mb,M_logo)>0 Then
	  response.write "[��]"
	  Else
	  response.write "[<FONT color=#ff0000>��</font>]"
	  End if%></TD>
      <TD colspan="2">
	  	    ����״̬��<INPUT name="m281" type="checkbox" value="1" <%IF strint(m_(281))=1 Then%>checked<%End IF%>>
	  	    ��ʾ���ߣ�<INPUT name="m282" type="checkbox" value="1" <%IF strint(m_(282))=1 Then%>checked<%End IF%>>			
	  	    ��ʾ����(0-15)��<a href="#" ONCLICK="window.open('../Include/datetype.htm','Sample','toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=yes,resizable=no,copyhistory=no,width=200,height=400,left=300,top=20')">������</a>
			<INPUT name="m283" type="text" value="<%=m_(283)%>" size="5"><br>
			��ʾ������<INPUT name="m284" type="text" value="<%=m_(284)%>" size="5">
            ���ⳤ�ȣ�<INPUT name="m285" type="text" value="<%=m_(285)%>" size="5">
            �ż㱳����
            <SELECT name="m286">
              <OPTION value="0" <%IF StrInt(m_(286))=0 Then%>selected<%End IF%>>Ҫ�ż㱳��</OPTION>
              <OPTION value="1" <%IF StrInt(m_(286))=1 Then%>selected<%End IF%>>��Ҫ�ż㱣��ʽ</OPTION>
			  <OPTION value="2" <%IF StrInt(m_(286))=2 Then%>selected<%End IF%>>��Ҫ�ż��޸�ʽ</OPTION> 
            </SELECT> 	
			</TD>
    </TR>
   <TR bgcolor="#FFFFFF">
      <TD height="26" align="right">��վ������ʾ���ã�<%mb="A,B,C,D,E,F"
	  If instr(mb,M_logo)>0 Then
	  response.write "[��]"
	  Else
	  response.write "[<FONT color=#ff0000>��</font>]"
	  End if%></TD>
      <TD colspan="2">
	  	    ���ù̶���<INPUT name="m287" type="checkbox" value="1" <%IF strint(m_(287))=1 Then%>checked<%End IF%>>
	  	    �������<INPUT name="m288" type="checkbox" value="1" <%IF strint(m_(288))=1 Then%>checked<%End IF%>>
	  	    ������ʾ��<INPUT name="m289" type="checkbox" value="1" <%IF strint(m_(289))=1 Then%>checked<%End IF%>>
			��ʾ����(0-15)��<a href="#" ONCLICK="window.open('../Include/datetype.htm','Sample','toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=yes,resizable=no,copyhistory=no,width=200,height=400,left=300,top=20')">������</a>
			<INPUT name="m290" type="text" value="<%=m_(290)%>" size="3">
			��ʾ������<INPUT name="m291" type="text" value="<%=m_(291)%>" size="3">
			��ʾ������<INPUT name="m292" type="text" value="<%=m_(292)%>" size="3">
			���ⳤ�ȣ�<INPUT name="m293" type="text" value="<%=m_(293)%>" size="3">
			�����ֺţ�<INPUT name="m294" type="text" value="<%=m_(294)%>" size="3">
			�����иߣ�<INPUT name="m295" type="text" value="<%=m_(295)%>" size="3">
</TD>
    </TR>
<table>
<tr>
	<td> </td>
</tr>
</table>
	
	<TR bgcolor="#FFFFFF"> 
      <TD height="30" colspan="3" align=center>	  
      <input name="Submit" type="submit" id="Submit" value="ȷ��">
	  <input name="Submit1" type="button" id="Submit1" value="����" onClick="location='admin_Template_Management.asp?id=<%=ID%>'"/>&nbsp;&nbsp;&nbsp;&nbsp; <a href="#" ONCLICK="window.open('../Include/datetype.htm','Sample','toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=yes,resizable=no,copyhistory=no,width=200,height=400,left=300,top=20')"><b>0-15����ʾ����</b></a> </TD>
      </TR>
  </FORM>
</TABLE>
      <BR></TD>
    </TR>
</TABLE>
</BODY>                    
                    
</HTML>                    
