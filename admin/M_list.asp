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
<%Dim Action,ID,Fso,Ts,i,Moldboard,Text,TextFile,m_name,temp,M_logo,mb,m_,data,M(96)'��ǰ��������ʵ�����õ�96
Moldboard=Request.Form("Moldboard")
ID=strint(request.QueryString("ID"))
Call OpenConn
IF Request.Form("Submit")<>"" Then
For i=0 to 96'��ǰ������
IF i=0 Then
m(i)=request.form("M"&i)
Else
m(i)=strint(Request.Form("M"&i))
End IF
Temp=Temp&m(i)
IF i<>96 Then Temp=Temp&"|"'��ǰ������
Next

Set Rs=server.createobject("adodb.recordset")
Sql="Select ID,M_list,M_liststr From DNJY_Template Where ID="&ID
Rs.Open Sql,conn,1,3
Rs(1)=Moldboard
Rs(2)=Temp
Rs.Update

set fso = CreateObject("Scripting.FileSystemObject")
set Ts = Fso.OpenTextFile(Server.Mappath("../list.asp"),2, True)
Ts.Write M_list(Moldboard,Temp)
Ts.Close:Set Ts=Nothing
Set Fso=Nothing
If Htmlhome=1 Then Call HomeHtml()'����������ҳ
response.write "<script language='javascript'>	alert('ģ�����óɹ������������µ�ҳ�棡');location.href='javascript:history.go(-1)';</script>"
End IF

Set Rs=Conn.Execute("Select M_list,M_liststr,M_Name,M_logo From DNJY_Template Where ID="&ID)
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
<TITLE>��Ϣ��Ŀģ��</title>
<style type="text/css">
<!--
.STYLE1 {font-weight: bold}
-->
</style>
</HEAD>
<BODY background="images/background.gif">                                                                 
  <TABLE width="98%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#799AE1">
    <TR>
      <TD height="20" bgcolor="#799AE1" align="center"><FONT color="#FFFFFF" style="font-size:14px">�� Ϣ �� Ŀ ģ �� �� ��
	  </FONT></TD>
    </TR>
    <TR>
      <TD bgcolor="#FFFFFF"> <BR>
            <TABLE width="98%" border="0"  align=center cellpadding="1" cellspacing="1" bgcolor="#D6DFF7">
  <FORM name="form1" method="post">
    <TR bgcolor="#FFFFFF">
      <TD width="20%" height="26" align="right">ģ�����<BR>
       </TD>
      <TD width="60%" >��ǰģ�壺��Ϣ��Ŀģ�壺list.asp,��ǰģ�巽����<FONT color="#ff0000"><%=Data(2,0)%></font>&nbsp;&nbsp;ģ���ʶ��<%=M_logo%><br>
        &nbsp;<TEXTAREA name="moldboard" cols="90" rows="25"><%=server.HTMLEncode(Data(0,0))%></TEXTAREA>
        <BR>
        <BR>
        </P></TD>
	<TD width="20%" height="26" align="left"> 
        ģ�������<BR>
<FONT color="#ff0000">
{$������}<br>
{$��Ϣ�����б�}<br>
{$���������б�}<br>
{$������Ϣ}<br>
{$������Ϣ}<br>
{$�Ƽ���Ϣ}<br>
{$������Ϣ}<br>
{$��Ŀ��Ϣ}<br>
{$��Ϣͼ���б�}<br></FONT>
<br> <br><a href="#" ONCLICK="window.open('../Include/datetype.htm','Sample','toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=yes,resizable=no,copyhistory=no,width=200,height=400,left=300,top=20')"><b>0-15��������ʽ</b></a> </TD>
      </TR>
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
      <TD height="26" align="right">��Ϣ������ʾ���ã�<%mb="A,B,C,D,E,F"
	  If instr(mb,M_logo)>0 Then
	  response.write "[��]"
	  Else
	  response.write "[<FONT color=#ff0000>��</font>]"
	  End if%></TD>
      <TD colspan="2">
	        ����Բ���־��<INPUT name="m2" type="checkbox" value="1" <%IF strint(m_(2))=1 Then%>checked<%End IF%>>
			������Ϣ������<INPUT name="m3" type="checkbox" value="1" <%IF strint(m_(3))=1 Then%>checked<%End IF%>>            
			��ɫ��ʾ��<INPUT name="m4" type="checkbox" value="1" <%IF strint(m_(4))=1 Then%>checked<%End IF%>> 
            ��ʾ������<INPUT name="m5" type="text" value="<%=m_(5)%>" size="6">
            �оࣺ<INPUT name="m6" type="text" value="<%=m_(6)%>" size="5">px
</TD>
    </TR>
    <TR bgcolor="#FFFFFF">
      <TD height="26" align="right" bgcolor="#BDDEEF">����������ʾ���ã�<%mb="A,B,C,D,E,F"
	  If instr(mb,M_logo)>0 Then
	  response.write "[��]"
	  Else
	  response.write "[<FONT color=#ff0000>��</font>]"
	  End if%></TD>
      <TD colspan="2" bgcolor="#BDDEEF">
	        ����Բ���־��<INPUT name="m7" type="checkbox" value="1" <%IF strint(m_(7))=1 Then%>checked<%End IF%>>
			������Ϣ������<INPUT name="m8" type="checkbox" value="1" <%IF strint(m_(8))=1 Then%>checked<%End IF%>>            
			��ɫ��ʾ��<INPUT name="m9" type="checkbox" value="1" <%IF strint(m_(9))=1 Then%>checked<%End IF%>> 
            ��ʾ������<INPUT name="m10" type="text" value="<%=m_(10)%>" size="6">
            �оࣺ<INPUT name="m11" type="text" value="<%=m_(11)%>" size="5">px</TD>
    </TR>

    <TR bgcolor="#FFFFFF">
      <TD height="26" align="right">������Ϣ��ʾ���ã�<%mb="A,B,C,D,E,F"
	  If instr(mb,M_logo)>0 Then
	  response.write "[��]"
	  Else
	  response.write "[<FONT color=#ff0000>��</font>]"
	  End if%></TD>
      <TD colspan="2">
	        ��Ϣ���ͣ�<INPUT name="m12" type="checkbox" value="1" <%IF strint(m_(12))=1 Then%>checked<%End IF%>>
			�Ƽ���־��<INPUT name="m13" type="checkbox" value="1" <%IF strint(m_(13))=1 Then%>checked<%End IF%>>
			���ù̶���<INPUT name="m14" type="checkbox" value="1" <%IF strint(m_(14))=1 Then%>checked<%End IF%>>
           	ͼƬ��־��<INPUT name="m15" type="checkbox" value="1" <%IF strint(m_(15))=1 Then%>checked<%End IF%>>
			������ʽ(0-15)��<a href="#" ONCLICK="window.open('../Include/datetype.htm','Sample','toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=yes,resizable=no,copyhistory=no,width=200,height=400,left=300,top=20')">������</a><INPUT name="m16" type="text" value="<%=m_(16)%>" size="5">			
			������ʾ��<INPUT name="m17" type="checkbox" value="1" <%IF strint(m_(17))=1 Then%>checked<%End IF%>>
			��ʾ���ࣺ<INPUT name="m18" type="checkbox" value="1" <%IF strint(m_(18))=1 Then%>checked<%End IF%>>
			������ɫ��<INPUT name="m19" type="checkbox" value="1" <%IF strint(m_(19))=1 Then%>checked<%End IF%>><br>
			�ż㱳����<INPUT name="m20" type="checkbox" value="1" <%IF strint(m_(20))=1 Then%>checked<%End IF%>>
			���������<INPUT name="m21" type="checkbox" value="1" <%IF strint(m_(21))=1 Then%>checked<%End IF%>>

����������<INPUT name="m22" type="text" value="<%=m_(22)%>" size="5">
��ʾ������<INPUT name="m23" type="text" value="<%=m_(23)%>" size="5"> 
��ʾ������<INPUT name="m24" type="text" value="<%=m_(24)%>" size="6">
�����С��<INPUT name="m25" type="text" value="<%=m_(25)%>" size="5">px
 �оࣺ<INPUT name="m26" type="text" value="<%=m_(26)%>" size="5">px</TD>
    </TR>
    <TR bgcolor="#FFFFFF">
      <TD height="26" align="right" bgcolor="#BDDEEF">������Ϣ��ʾ���ã�<%mb="A,B,C,D,E,F"
	  If instr(mb,M_logo)>0 Then
	  response.write "[��]"
	  Else
	  response.write "[<FONT color=#ff0000>��</font>]"
	  End if%></TD>
      <TD colspan="2" bgcolor="#BDDEEF">
	        ��Ϣ���ͣ�<INPUT name="m27" type="checkbox" value="1" <%IF strint(m_(27))=1 Then%>checked<%End IF%>>
			�Ƽ���־��<INPUT name="m28" type="checkbox" value="1" <%IF strint(m_(28))=1 Then%>checked<%End IF%>>
			���ù̶���<INPUT name="m29" type="checkbox" value="1" <%IF strint(m_(29))=1 Then%>checked<%End IF%>>
           	ͼƬ��־��<INPUT name="m30" type="checkbox" value="1" <%IF strint(m_(30))=1 Then%>checked<%End IF%>>
			������ʽ(0-15)��<a href="#" ONCLICK="window.open('../Include/datetype.htm','Sample','toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=yes,resizable=no,copyhistory=no,width=200,height=400,left=300,top=20')">������</a><INPUT name="m31" type="text" value="<%=m_(31)%>" size="5">
			������ʾ��<INPUT name="m32" type="checkbox" value="1" <%IF strint(m_(32))=1 Then%>checked<%End IF%>>
			��ʾ���ࣺ<INPUT name="m33" type="checkbox" value="1" <%IF strint(m_(33))=1 Then%>checked<%End IF%>>
			������ɫ��<INPUT name="m34" type="checkbox" value="1" <%IF strint(m_(34))=1 Then%>checked<%End IF%>><br>
			�ż㱳����<INPUT name="m35" type="checkbox" value="1" <%IF strint(m_(35))=1 Then%>checked<%End IF%>>
			���������<INPUT name="m36" type="checkbox" value="1" <%IF strint(m_(36))=1 Then%>checked<%End IF%>>
			
����������<INPUT name="m37" type="text" value="<%=m_(37)%>" size="5">
��ʾ������<INPUT name="m38" type="text" value="<%=m_(38)%>" size="5"> 
��ʾ������<INPUT name="m39" type="text" value="<%=m_(39)%>" size="6">
�����С��<INPUT name="m40" type="text" value="<%=m_(40)%>" size="5">px
 �оࣺ<INPUT name="m41" type="text" value="<%=m_(41)%>" size="5">px
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
	        ��Ϣ���ͣ�<INPUT name="m42" type="checkbox" value="1" <%IF strint(m_(42))=1 Then%>checked<%End IF%>>
			�Ƽ���־��<INPUT name="m43" type="checkbox" value="1" <%IF strint(m_(43))=1 Then%>checked<%End IF%>>
			���ù̶���<INPUT name="m44" type="checkbox" value="1" <%IF strint(m_(44))=1 Then%>checked<%End IF%>>
           	ͼƬ��־��<INPUT name="m45" type="checkbox" value="1" <%IF strint(m_(45))=1 Then%>checked<%End IF%>>
			������ʽ(0-15)��<a href="#" ONCLICK="window.open('../Include/datetype.htm','Sample','toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=yes,resizable=no,copyhistory=no,width=200,height=400,left=300,top=20')">������</a><INPUT name="m46" type="text" value="<%=m_(46)%>" size="5">			
			������ʾ��<INPUT name="m47" type="checkbox" value="1" <%IF strint(m_(47))=1 Then%>checked<%End IF%>>
			��ʾ���ࣺ<INPUT name="m48" type="checkbox" value="1" <%IF strint(m_(48))=1 Then%>checked<%End IF%>>
			������ɫ��<INPUT name="m49" type="checkbox" value="1" <%IF strint(m_(49))=1 Then%>checked<%End IF%>><br>
			�ż㱳����<INPUT name="m50" type="checkbox" value="1" <%IF strint(m_(50))=1 Then%>checked<%End IF%>>
			���������<INPUT name="m51" type="checkbox" value="1" <%IF strint(m_(51))=1 Then%>checked<%End IF%>>

����������<INPUT name="m52" type="text" value="<%=m_(52)%>" size="5">
��ʾ������<INPUT name="m53" type="text" value="<%=m_(53)%>" size="5"> 
��ʾ������<INPUT name="m54" type="text" value="<%=m_(54)%>" size="6">
�����С��<INPUT name="m55" type="text" value="<%=m_(55)%>" size="5">px
 �оࣺ<INPUT name="m56" type="text" value="<%=m_(56)%>" size="5">px
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
	        ��Ϣ���ͣ�<INPUT name="m57" type="checkbox" value="1" <%IF strint(m_(57))=1 Then%>checked<%End IF%>>
			�Ƽ���־��<INPUT name="m58" type="checkbox" value="1" <%IF strint(m_(58))=1 Then%>checked<%End IF%>>
			���ù̶���<INPUT name="m59" type="checkbox" value="1" <%IF strint(m_(59))=1 Then%>checked<%End IF%>>
           	ͼƬ��־��<INPUT name="m60" type="checkbox" value="1" <%IF strint(m_(60))=1 Then%>checked<%End IF%>>
			������ʽ(0-15)��<a href="#" ONCLICK="window.open('../Include/datetype.htm','Sample','toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=yes,resizable=no,copyhistory=no,width=200,height=400,left=300,top=20')">������</a><INPUT name="m61" type="text" value="<%=m_(61)%>" size="5">			
			������ʾ��<INPUT name="m62" type="checkbox" value="1" <%IF strint(m_(62))=1 Then%>checked<%End IF%>>
			��ʾ���ࣺ<INPUT name="m63" type="checkbox" value="1" <%IF strint(m_(63))=1 Then%>checked<%End IF%>>
			������ɫ��<INPUT name="m64" type="checkbox" value="1" <%IF strint(m_(64))=1 Then%>checked<%End IF%>><br>
			�ż㱳����<INPUT name="m65" type="checkbox" value="1" <%IF strint(m_(65))=1 Then%>checked<%End IF%>>
			���������<INPUT name="m66" type="checkbox" value="1" <%IF strint(m_(66))=1 Then%>checked<%End IF%>>

����������<INPUT name="m67" type="text" value="<%=m_(67)%>" size="5">
��ʾ������<INPUT name="m68" type="text" value="<%=m_(68)%>" size="5"> 
��ʾ������<INPUT name="m69" type="text" value="<%=m_(69)%>" size="6">
�����С��<INPUT name="m70" type="text" value="<%=m_(70)%>" size="5">px
 �оࣺ<INPUT name="m71" type="text" value="<%=m_(71)%>" size="5">px
</TD>
    </TR>
    <TR bgcolor="#FFFFFF">
      <TD height="26" align="right">������Ϣ�б���ʾ���ã�<%mb="A,B,C,D,E"
	  If instr(mb,M_logo)>0 Then
	  response.write "[��]"
	  Else
	  response.write "[<FONT color=#ff0000>��</font>]"
	  End if%></TD>
      <TD colspan="2">
	        ��Ϣ���ͣ�<INPUT name="m72" type="checkbox" value="1" <%IF strint(m_(72))=1 Then%>checked<%End IF%>>
			ͼ����ʾ��<INPUT name="m73" type="checkbox" value="1" <%IF strint(m_(73))=1 Then%>checked<%End IF%>>            
ÿҳ������<INPUT name="m74" type="text" value="<%=m_(74)%>" size="5">
ͼƬ��ȣ�<INPUT name="m75" type="text" value="<%=m_(75)%>" size="5">
ͼƬ�߶ȣ�<INPUT name="m76" type="text" value="<%=m_(76)%>" size="5">
���ⳤ�ȣ�<INPUT name="m77" type="text" value="<%=m_(77)%>" size="5">
���Ľ�ȡ���ȣ�<INPUT name="m78" type="text" value="<%=m_(78)%>" size="5"> 
</TD>
    </TR>

    <TR bgcolor="#BDDEEF">
      <TD height="26" align="right">��Ϣͼ����ʾ����(M)��<%mb="F"
	  If instr(mb,M_logo)>0 Then
	  response.write "[��]"
	  Else
	  response.write "[<FONT color=#ff0000>��</font>]"
	  End if%></TD>
      <TD colspan="2">
	        ��Ϣ���ͣ�<INPUT name="m79" type="checkbox" value="1" <%IF strint(m_(79))=1 Then%>checked<%End IF%>>
			���ù̶���<INPUT name="m80" type="checkbox" value="1" <%IF strint(m_(80))=1 Then%>checked<%End IF%>>
			�Ƽ���־��<INPUT name="m81" type="checkbox"  value="1" <%IF strint(m_(81))=1 Then%>checked<%End IF%>>
           	ͼƬ��־��<INPUT name="m82" type="checkbox"  value="1" <%IF strint(m_(82))=1 Then%>checked<%End IF%>>						
			���������<INPUT name="m83" type="checkbox"  value="1" <%IF strint(m_(83))=1 Then%>checked<%End IF%>>
			��ʾ����(0-15)��<a href="#" ONCLICK="window.open('../Include/datetype.htm','Sample','toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=yes,resizable=no,copyhistory=no,width=200,height=400,left=300,top=20')">������</a><INPUT name="m84" type="text" value="<%=m_(84)%>" size="5">
�г��۸�<INPUT name="m85" type="checkbox"  value="1" <%IF strint(m_(85))=1 Then%>checked<%End IF%>>
��վ�۸�<INPUT name="m86" type="checkbox"  value="1" <%IF strint(m_(86))=1 Then%>checked<%End IF%>><br>
���ⳤ�ȣ�<INPUT name="m87" type="text" value="<%=m_(87)%>" size="5">
������ɫ��<INPUT name="m88" type="checkbox" value="1" <%IF strint(m_(88))=1 Then%>checked<%End IF%>>
�����С��<INPUT name="m89" type="text" value="<%=m_(89)%>" size="5">px
ͼƬ��ȣ�<INPUT name="m90" type="text" value="<%=m_(90)%>" size="6">
ͼƬ�߶ȣ�<INPUT name="m91" type="text" value="<%=m_(91)%>" size="6">
��Ԫ��ȣ�<INPUT name="m92" type="text" value="<%=m_(92)%>" size="6">
��Ԫ�߶ȣ�<INPUT name="m93" type="text" value="<%=m_(93)%>" size="6"><br>
ÿҳ������<INPUT name="m94" type="text" value="<%=m_(94)%>" size="5"> 
��ʾ������<INPUT name="m95" type="text" value="<%=m_(95)%>" size="6">
��ʾ��ҳ��<INPUT name="m96" type="checkbox" value="1" <%IF strint(m_(96))=1 Then%>checked<%End IF%>>

    </TR>
    <TR bgcolor="#FFFFFF"> 
      <TD height="30" align=left></TD>
      <TD height="30" colspan="2" align=left>
	  <INPUT type="submit" name="Submit" value="  �� ��  ">
        <input name="Submit1" type="button" id="Submit1" value="����" onClick="location='admin_Template_Management.asp?id=<%=ID%>'"/>&nbsp;&nbsp;&nbsp;&nbsp;<a href="#" ONCLICK="window.open('../Include/datetype.htm','Sample','toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=yes,resizable=no,copyhistory=no,width=200,height=400,left=300,top=20')"><b>0-15��������ʽ</b></a> </TD>
    </TR>
  </FORM>
</TABLE>
      <BR>      </TD>
    </TR>
  </TABLE>
</BODY>                    
                    
</HTML>                    
