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

<%Dim Action,ID,fso,Ts,Data,Moldboard,i,FsObject,M_logo
Moldboard=Request.Form("Moldboard")
ID=strint(request.QueryString("ID"))
Call OpenConn
IF Request.Form("Submit")<>"" Then
set rs=server.createobject("adodb.recordset")
Sql="Select ID,M_top,M_Name From DNJY_Template Where ID="&ID
Rs.Open Sql,conn,1,3
Rs(1)=Moldboard
Rs.Update

set fso = CreateObject("Scripting.FileSystemObject")
'set Fso = CreateObject(FsObject)
set Ts = Fso.OpenTextFile(Server.Mappath("../top.asp"),2, True)
Ts.Write Moldboard
Ts.Close:Set Ts=Nothing
Set Fso=Nothing
If Htmlhome=1 Then Call HomeHtml()'����������ҳ
response.write "<script language='javascript'>	alert('ģ�����óɹ������������µ�ҳ��\n\n���������ɾ�̬����ҳ��');location.href='javascript:history.go(-1)';</script>"
End IF

Set Rs=Conn.Execute("Select M_top,M_Name,M_logo From DNJY_Template Where ID="&ID)
IF Not Rs.Eof Then Data=Rs.GetRows
Rs.close:Set Rs=Nothing
Call CloseDB(conn)
M_logo=Data(2,0)
%>
<HEAD>
<LINK rel="stylesheet" type="text/css" href="../css/style.css">
<META http-equiv="Content-Type" content="text/html; charset=gb2312">
<TITLE>ͷ��ģ������</title>
<style type="text/css">
<!--
body,td,th {
	color: #135294;
}
body {
	background-color: #D6DFF7;
}
-->
</style></HEAD>
<BODY background="images/background.gif">                                                                 
  <TABLE width="98%" border="0" align="center" cellpadding="4" cellspacing="1" bgcolor="#799AE1">
    <TR>
      <TD width="100%" height="20" align="center" bgcolor="#799AE1"><FONT color="#FFFFFF" style="font-size:14px">ͷ �� ģ �� �� �� </FONT></TD>
    </TR>
	<FORM name="form2" method="post" action="">
    <TR>
      <TD align="center" bgcolor="#FFFFFF"><br>
        <table width="98%" border="0" cellpadding="0" cellspacing="1" bgcolor="#D6DFF7">
        <tr><TD bgcolor="#FFFFFF">
	  <DIV align="right">ģ����룺  <br>  	</DIV></TD>
          <td bgcolor="#FFFFFF">&nbsp;&nbsp;
		  ��ǰģ�壺ͷ��ģ��top.asp����ǰģ�巽����<FONT color="#ff0000"><%=Data(1,0)%></font>&nbsp;&nbsp;ģ���ʶ��<%=M_logo%><br>
              <TEXTAREA name="Moldboard" cols="100" rows="30" id="T2"><%=server.HTMLEncode(Data(0,0))%></TEXTAREA>
<br>
<br></td>
        </tr>
      </table>
      <br></TD>
    </TR>
	<TR>
	  <TD height="30" align="center" bgcolor="#FFFFFF">
	        <input name="Submit" type="submit" id="Submit" value="ȷ��">
			<input name="Submit1" type="button" id="Submit1" value="����" onClick="location='admin_Template_Management.asp?id=<%=ID%>'"/></TD>
	  </TR>
	  </FORM>
  </TABLE>
</BODY>                    
                    
</HTML>                    
