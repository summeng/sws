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
<%Dim Action,ID,fso,Ts,Data,i,Moldboard,Text,TextFile,m_name,temp,M_logo,mb,m_,M(40)'��ǰ��������ʵ�����õ�40
Moldboard=Request.Form("Moldboard")
ID=strint(request.QueryString("ID"))
Call OpenConn
IF Request.Form("Submit")<>"" Then
For i=0 to 40'��ǰ������
IF i=0 Then
m(i)=request.form("M"&i)
Else
m(i)=strint(Request.Form("M"&i))
End IF
Temp=Temp&m(i)
IF i<>40 Then Temp=Temp&"|"'��ǰ������
Next
Set Rs=server.createobject("adodb.recordset")
Sql="Select ID,M_show,M_showstr From DNJY_Template Where ID="&ID
Rs.Open Sql,conn,1,3
Rs(1)=Moldboard
Rs(2)=Temp
Rs.Update

set fso = CreateObject("Scripting.FileSystemObject")
set Ts = Fso.OpenTextFile(Server.Mappath("../show.asp"),2, True)
Ts.Write M_show(Moldboard,Temp)
Ts.Close:Set Ts=Nothing
Set Fso=Nothing
If Htmlhome=1 Then Call HomeHtml()'����������ҳ
response.write "<script language='javascript'>	alert('ģ�����óɹ������������µ�ҳ��\n\n���������ɾ�̬����ҳ��');location.href='javascript:history.go(-1)';</script>"
End IF

Set Rs=Conn.Execute("Select M_show,M_showstr,M_Name,M_logo From DNJY_Template Where ID="&ID)
IF Not Rs.Eof Then Data=Rs.GetRows
Rs.close:Set Rs=Nothing
Call CloseDB(conn)
M_logo=Data(3,0)
M_=Split(Data(1,0),"|")%>
<HTML>
<HEAD>
<LINK rel="stylesheet" type="text/css" href="../css/style.css">
<META http-equiv="Content-Type" content="text/html; charset=gb2312">
<META name="GENERATOR" content="Microsoft FrontPage 6.0">
<META name="ProgId" content="FrontPage.Editor.Document">
<TITLE>��Ϣҳģ��</title>
</HEAD>
<BODY background="images/background.gif">                                                                 
  <TABLE width="98%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#799AE1">
    <TR>
      <TD height="20" bgcolor="#799AE1" align="center"><FONT color="#FFFFFF" style="font-size:14px">�� Ϣ �� ʾ ҳ �� ģ ��
	  </FONT></TD>
    </TR>
    <TR>
      <TD bgcolor="#FFFFFF"> <BR>
            <TABLE width="98%" border="0"  align=center cellpadding="1" cellspacing="1" bgcolor="#D6DFF7">
  <FORM name="form1" method="post">
    <TR bgcolor="#FFFFFF">
      <TD width="15%" height="26" align="right">ģ�����</TD>
      <TD width="85%" valign="top">��ǰģ�壺
        ��Ϣ��ʾҳģ��show.asp����ǰģ�巽����<FONT color="#ff0000"><%=Data(2,0)%></font>&nbsp;&nbsp;ģ���ʶ��<%=M_logo%><br> <TEXTAREA name="moldboard" cols="100" rows="30"><%=server.HTMLEncode(Data(0,0))%></TEXTAREA>
        <BR>
        <BR>
        </P> ģ���ǩ��<br>
      ��ͨ�����±�ǩ����������ݣ�<br>
<FONT color="#ff0000">{$��վ����}
{$��վ�ؼ���}
{$��վ���}
{$��Ϣ���}
{$��Ϣ����}
{$������Ϣ}
{$�Ƽ���Ϣ}
{$������Ϣ}
{$��������}
{$��Ϣ������ʾ}
{$����ʱ��}
{$��Чʱ��}
{$�������}
{$�ظ�����}
{$������¼}
{$�����ٱ�}
{$��Ϣ���}
{$����״̬}
{$���׼۸�}
{$���׵���}{$IP��ַ}{$�̼ҵ���}{$��ԱID��ʾ}{$��ϢͼƬ}{$��ϵ����ʾ}{$�̶��绰}{$�ƶ��绰}{$QQ����}{$��������}{$��˾����}{$��ϵ��ַ}{$��������}{$��ϸ˵��}{$��վ����}{$ת������}<BR>
{$�ظ�����ʼ}{$ֱ�ӻظ�}{$����ظ�}{$�ظ��û�����} {$��������}{$�ظ��ʼ������}{$�ظ������}{$�ظ��������}{$�ظ������ı���}{$��֤��}{$�ظ��ύ��ť}{$�ظ�������} ����������ڡ�{$�ظ�����ʼ}����{$�ظ�������}��֮�ڣ�<BR>

{$���zuo1}�����ݹ��ID</font> </TD>
      </TR>
    <TR bgcolor="#FFFFFF">
      <TD height="26" align="right"><b>ע�⣺</b></TD>
      <TD colspan="2"><b>���´�[��]Ϊ��ǰģ����Ч,���ò�����Ч����[<FONT color=#ff0000>��</font>]Ϊ��ǰģ����Ч,���ò�����Ч��</b></TD>
    </TR>
    <TR bgcolor="#FFFFFF">
      <TD height="26" align="right">��վ������<%mb="A,B,C,D,E,F"
	  If instr(mb,M_logo)>0 Then
	  response.write "[��]"
	  Else
	  response.write "[<FONT color=#ff0000>��</font>]"
	  End if%></TD>
      <TD colspan="2">&nbsp;<TEXTAREA name="m0" cols="80" rows="10"><%=m_(0)%></TEXTAREA></TD></TR>
    <TR bgcolor="#FFFFFF">
      <TD height="26" align="right">������Ϣ��ʾ���ã�<%mb="A,B,C,D,E"
	  If instr(mb,M_logo)>0 Then
	  response.write "[��]"
	  Else
	  response.write "[<FONT color=#ff0000>��</font>]"
	  End if%></TD>
      <TD colspan="2">&nbsp;��ʾ��Ŀ��ʽ��
        <SELECT name="m1" id="m1">
          <OPTION value="0" <%IF strint(m_(1))=0 Then%>selected<%End IF%>>����ʾ</OPTION>
          <OPTION value="1" <%IF strint(m_(1))=1 Then%>selected<%End IF%>>һ����Ŀ</OPTION>
          <OPTION value="2" <%IF strint(m_(1))=2 Then%>selected<%End IF%>>������Ŀ</OPTION>
        </SELECT>
        ��ʾ������<INPUT name="m2" type="text" id="m2" value="<%=m_(2)%>" size="5">
        ��ʾ������<INPUT name="m3" type="text" id="m3" value="<%=m_(3)%>" size="5">
&nbsp;��ʾ����:<INPUT name="m4" type="checkbox" id="m4" value="1" <%IF strint(m_(4))=1 Then%>checked<%End IF%>> ���ù̶�:<INPUT name="m5" type="checkbox" id="m5" value="1" <%IF strint(m_(5))=1 Then%>checked<%End IF%>>  �������:<INPUT name="m6" type="checkbox" id="m6" value="1" <%IF strint(m_(6))=1 Then%>checked<%End IF%>><BR>����������<INPUT name="m7" type="text" id="m7" value="<%=m_(7)%>" size="5">
�����С��<INPUT name="m8" type="text" id="m8" value="<%=m_(8)%>" size="5">px 
�оࣺ<INPUT name="m9" type="text" id="m9" value="<%=m_(9)%>" size="5">px ������ʽ(0-15)��<a href="#" ONCLICK="window.open('../Include/datetype.htm','Sample','toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=yes,resizable=no,copyhistory=no,width=200,height=400,left=300,top=20')">������</a><INPUT name="m10" type="text" id="m10" value="<%=m_(10)%>" size="5">

<b>&nbsp;&nbsp;&nbsp;&nbsp;<a href="#" ONCLICK="window.open('../Include/datetype.htm','Sample','toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=yes,resizable=no,copyhistory=no,width=200,height=400,left=300,top=20')">0-15��������ʽ</b></a></TD>
    </TR>
    <TR bgcolor="#FFFFFF">
      <TD height="26" align="right" bgcolor="#BDDEEF">������Ϣ��ʾ���ã�<%mb="A,B,C,D,E"
	  If instr(mb,M_logo)>0 Then
	  response.write "[��]"
	  Else
	  response.write "[<FONT color=#ff0000>��</font>]"
	  End if%></TD>
      <TD colspan="2" bgcolor="#BDDEEF">&nbsp;��ʾ��Ŀ��ʽ��
        <SELECT name="m11" id="m11">
          <OPTION value="0" <%IF strint(m_(11))=0 Then%>selected<%End IF%>>����ʾ</OPTION>
          <OPTION value="1" <%IF strint(m_(11))=1 Then%>selected<%End IF%>>һ����Ŀ</OPTION>
          <OPTION value="2" <%IF strint(m_(11))=2 Then%>selected<%End IF%>>������Ŀ</OPTION>
        </SELECT>
        ��ʾ������<INPUT name="m12" type="text" id="m12" value="<%=strint(m_(12))%>" size="5">
        ��ʾ������<INPUT name="m13" type="text" id="m13" value="<%=strint(m_(13))%>" size="5">
&nbsp;��ʾ����:<INPUT name="m14" type="checkbox" id="m14" value="1" <%IF strint(m_(14))=1 Then%>checked<%End IF%>> ���ù̶�:<INPUT name="m15" type="checkbox" id="m15" value="1" <%IF strint(m_(15))=1 Then%>checked<%End IF%>>  �������:<INPUT name="m16" type="checkbox" id="m16" value="1" <%IF strint(m_(16))=1 Then%>checked<%End IF%>><BR> ����������<INPUT name="m17" type="text" id="m17" value="<%=strint(m_(17))%>" size="5">
�����С:<INPUT name="m18" type="text" id="m18" value="<%=m_(18)%>" size="5"> px 
�оࣺ<INPUT name="m19" type="text" id="m19" value="<%=m_(19)%>" size="5">px ������ʽ(0-15)��<a href="#" ONCLICK="window.open('../Include/datetype.htm','Sample','toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=yes,resizable=no,copyhistory=no,width=200,height=400,left=300,top=20')">������</a><INPUT name="m20" type="text" id="m20" value="<%=m_(20)%>" size="5">

</TD>
    </TR>
    <TR bgcolor="#FFFFFF">
      <TD height="26" align="right">�Ƽ���Ϣ��ʾ���ã�<%mb="A,B,C,D,E"
	  If instr(mb,M_logo)>0 Then
	  response.write "[��]"
	  Else
	  response.write "[<FONT color=#ff0000>��</font>]"
	  End if%></TD>
      <TD colspan="2">&nbsp;��ʾ��Ŀ��ʽ��
        <SELECT name="m21" id="m21">
          <OPTION value="0" <%IF strint(m_(21))=0 Then%>selected<%End IF%>>����ʾ</OPTION>
          <OPTION value="1" <%IF strint(m_(21))=1 Then%>selected<%End IF%>>һ����Ŀ</OPTION>
          <OPTION value="2" <%IF strint(m_(21))=2 Then%>selected<%End IF%>>������Ŀ</OPTION>
        </SELECT>
��ʾ������<INPUT name="m22" type="text" id="m22" value="<%=strint(m_(22))%>" size="5">
��ʾ������<INPUT name="m23" type="text" id="m23" value="<%=strint(m_(23))%>" size="5">
&nbsp;��ʾ����:<INPUT name="m24" type="checkbox" id="m24" value="1" <%IF strint(m_(24))=1 Then%>checked<%End IF%>> ���ù̶�:<INPUT name="m25" type="checkbox" id="m25" value="1" <%IF strint(m_(25))=1 Then%>checked<%End IF%>>  �������:<INPUT name="m26" type="checkbox" id="m26" value="1" <%IF strint(m_(26))=1 Then%>checked<%End IF%>><BR> ����������<INPUT name="m27" type="text" id="m27" value="<%=strint(m_(27))%>" size="5">
�����С:<INPUT name="m28" type="text" id="m28" value="<%=m_(28)%>" size="5"> px 
�оࣺ<INPUT name="m29" type="text" id="m29" value="<%=m_(29)%>" size="5">px ������ʽ(0-15)��<a href="#" ONCLICK="window.open('../Include/datetype.htm','Sample','toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=yes,resizable=no,copyhistory=no,width=200,height=400,left=300,top=20')">������</a><INPUT name="m30" type="text" id="m30" value="<%=m_(30)%>" size="5">

</TD>
    </TR>
    <TR bgcolor="#FFFFFF">
      <TD height="26" align="right" bgcolor="#BDDEEF">������Ϣ��ʾ���ã�<%mb="A,B,C,D,E"
	  If instr(mb,M_logo)>0 Then
	  response.write "[��]"
	  Else
	  response.write "[<FONT color=#ff0000>��</font>]"
	  End if%></TD>
      <TD colspan="2" bgcolor="#BDDEEF">&nbsp;��ʾ��Ŀ��ʽ��
        <SELECT name="m31" id="m31">
          <OPTION value="0" <%IF strint(m_(31))=0 Then%>selected<%End IF%>>����ʾ</OPTION>
          <OPTION value="1" <%IF strint(m_(31))=1 Then%>selected<%End IF%>>һ����Ŀ</OPTION>
          <OPTION value="2" <%IF strint(m_(31))=2 Then%>selected<%End IF%>>������Ŀ</OPTION>
        </SELECT>
��ʾ������<INPUT name="m32" type="text" id="m32" value="<%=strint(m_(32))%>" size="5">
��ʾ������<INPUT name="m33" type="text" id="m33" value="<%=strint(m_(33))%>" size="5">
&nbsp;��ʾ����:<INPUT name="m34" type="checkbox" id="m34" value="1" <%IF strint(m_(34))=1 Then%>checked<%End IF%>> ���ù̶�:<INPUT name="m35" type="checkbox" id="m35" value="1" <%IF strint(m_(35))=1 Then%>checked<%End IF%>>  �������:<INPUT name="m36" type="checkbox" id="m36" value="1" <%IF strint(m_(36))=1 Then%>checked<%End IF%>><BR> ����������<INPUT name="m37" type="text" id="m37" value="<%=strint(m_(37))%>" size="5">
�����С:<INPUT name="m38" type="text" id="m38" value="<%=m_(38)%>" size="5"> px 
�оࣺ<INPUT name="m39" type="text" id="m39" value="<%=m_(39)%>" size="5">px ������ʽ(0-15)��<a href="#" ONCLICK="window.open('../Include/datetype.htm','Sample','toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=yes,resizable=no,copyhistory=no,width=200,height=400,left=300,top=20')">������</a><INPUT name="m40" type="text" id="m40" value="<%=m_(40)%>" size="5">

</TD>
    </TR>
    <TR bgcolor="#FFFFFF"> 
      <TD height="30" align=left></TD>
      <TD height="30" colspan="2" align=left>
	  <INPUT type="submit" name="Submit" value="  ���沢����  ">
	  <input name="Submit1" type="button" id="Submit1" value="����" onClick="location='admin_Template_Management.asp?id=<%=ID%>'"/>
        </TD>
    </TR>
  </FORM>

</TABLE>
      <BR>      </TD>
    </TR>
  </TABLE>
</BODY>                    
                    
</HTML>                    
