<!--#include file="../conn.asp"-->
<!--#include file="../SqlIn.Asp"-->
<!--#include file=cookies.asp-->
<!--#include file="../config.asp"-->
<!--#include file="../Include/label.asp"-->
<!--#include file="../Include/Function.asp"-->
<!--#include file="../Include/Form_board .asp"-->
<%
'=====================================
'���칩����Ϣ��վ����ϵͳ
'��������Ƽ������ҳ�Ʒ
'�ͷ�����http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>
<%call checkmanage("11")%>
<%Dim ID,M_Name,Fs,Fso,Ts,Rs1,FsObject,D_name,k,M_Dir,T_dir,M_logo,T_logo
ID=Strint(Request.QueryString("ID"))
M_Name=Request.Form("T_Name")
T_dir=Request.Form("T_dir")
T_logo=Request.Form("T_logo")

Call OpenConn
Sql="Select ID,M_Name From DNJY_Template Where M_Reclaim=0"
Set Rs=Conn.Execute(Sql)
if rs.eof or rs.bof then
response.write "û��ģ�巽��������<a href=""Admin_Template.asp?Action=Import"" TARGET=right><FONT color=""#0000FF"" >����ģ�巽��</font></a>����ע��ģ�巽���ж�������ٷ�����ֻ����ѡ����һ��ģ���н��У����ѡ��ģ�巽����Ҫ�ٸ����������޷��ṩ������Ҫ�����շѡ�����ʱ��ȫ�����룬ѡ���������ΪĬ�ϼ��ɣ������Ŀɷ������վ��ɾ������"
Call CloseRs(rs)
response.End
End If
Call CloseRs(rs)

Sql="Select ID,M_Name From DNJY_Template Where M_Tacitly=1 and M_Reclaim=0"
Set Rs=Conn.Execute(Sql)
if rs.eof or rs.bof then
Set Rs=server.CreateObject("adodb.recordset")
Sql="Select * From DNJY_Template Where M_Reclaim=0"
Rs.Open Sql,conn,1,3
Rs("M_Tacitly")=1
Rs.MoveFirst
Rs.Update
Call CloseRs(rs)
End If

IF ID=0 Then 
	Sql="Select ID,M_Name,M_Tacitly,M_Dir,M_logo From DNJY_Template Where M_Tacitly=1 and M_Reclaim=0"
Else
	Sql="Select ID,M_Name,M_Dir,M_logo From DNJY_Template where M_Reclaim=0 and ID="&ID	
End If
Set Rs=Conn.Execute(Sql)

D_name=Rs("M_Name")
M_Dir=Rs("M_Dir")
M_logo=Rs("M_logo")
IF Not Rs.Eof Then id=Rs(0)


IF Request("action")="Reclaim" Then
Conn.Execute("Update DNJY_Template Set M_Reclaim=1,M_Tacitly=0 Where ID="&ID&"")
Set Rs=server.CreateObject("adodb.recordset")
Sql="Select * From DNJY_Template Where M_Reclaim=0"
Rs.Open Sql,conn,1,3
Rs("M_Tacitly")=1
Rs.MoveFirst
Rs.Update
Call CloseRs(rs)

Response.Write"<script>alert('�ѽ�ѡ���ģ�巽���������վ�����ڻ���վ�н�ģ�巽���һأ�');location.href=""admin_Template_Management.asp"";</script>"
Response.end  
End IF


IF Request("action")="Tacitly" Then

	Conn.Execute("Update DNJY_Template Set M_Tacitly=0")
	Conn.Execute("Update DNJY_Template Set M_Tacitly=1 Where ID="&ID&"")	
	Set Rs=Conn.Execute("Select M_top,M_topstr,M_index,M_indexstr,M_xxindex,M_xxindexstr,M_Footer,M_Footerstr,M_show,M_showstr,M_list,M_liststr,M_css,M_News,M_Newsstr,M_News_More,M_News_Morestr,M_Picw,M_Pich,M_logo From DNJY_Template Where ID="&ID&"")	
	set fso = CreateObject("Scripting.FileSystemObject")
	set Ts = Fso.OpenTextFile(Server.Mappath("../top.asp"),2, True)
	Ts.Write M_top(Rs(0))
	
	set Ts = Fso.OpenTextFile(Server.Mappath("../Index.asp"),2, True)
	Ts.Write M_Index(Rs(2),Rs(3))

	set Ts = Fso.OpenTextFile(Server.Mappath("../xxIndex.asp"),2, True)
	Ts.Write M_xxIndex(Rs(4),Rs(5))
		
	set Ts = Fso.OpenTextFile(Server.Mappath("../Footer.asp"),2, True)
	Ts.Write M_Footer(Rs(6),Rs(7))
	
	set Ts = Fso.OpenTextFile(Server.Mappath("../show.asp"),2, True)
	Ts.Write M_show(Rs(8),Rs(9))
	
	set Ts = Fso.OpenTextFile(Server.Mappath("../list.asp"),2, True)
	Ts.Write M_list(Rs(10),Rs(11))
If Not IsNull(Rs(12)) Then
if fso.folderexists(server.mappath("../skin/"&M_Dir&"")) then
	set Ts = Fso.OpenTextFile(Server.Mappath("../skin/"&M_Dir&"/style.css"),2, True)
	 Ts.Write Rs(12)
	 Else
Response.Write"<script>alert('ģ����Ŀ¼../skin/"&M_Dir&"�����ڣ�������ٲ�����');location.href=""admin_Template_Management.asp"";</script>"
Response.end 
	 End if
End if	
	'set Ts = Fso.OpenTextFile(Server.Mappath("../list.asp"),2, True)
	'Ts.Write News(Rs(10),Rs(11))
	
	'set Ts = Fso.OpenTextFile(Server.Mappath("../News_list.asp"),2, True)
	'Ts.Write News_More(Rs(12),Rs(13))	
	Ts.Close:Set Ts=Nothing
	Set Fso=Nothing
If Htmlhome=1 Then Call HomeHtml()'����������ҳ
response.write "<script language='javascript'>	alert('Ĭ��ģ�巽�����óɹ�\n\n����������ɾ�̬ҳ���������ɾ�̬����ҳ��');location.href='javascript:history.go(-1)';</script>"
End IF

IF Request.Form("submit")<>"" Then
Set Rs=Conn.Execute("Select M_Name From DNJY_Template Where M_Name='"&M_Name&"' or M_dir='"&T_dir&"' or M_logo='"&T_logo&"'")
	IF Not Rs.Eof Then
	response.write "<script language='javascript'>	alert('ģ�淽������ģ��Ŀ¼����ģ���ʶ���Ѵ���,�����ظ�������ѡ');location.href='javascript:history.go(-1)';</script>"		
	Response.End
	Else

response.write "��Ѱ�û�д˹��ܣ�"

	End IF
End IF
%>
<HTML>
<HEAD>
<META http-equiv="Content-Type" content="text/html; charset=gb2312">
<TITLE>ģ�����</title>
<LINK rel="stylesheet" type="text/css" href="../css/style.css">
<style type="text/css">
<!--
body {
	background-color: #D6DFF7;
}
body,td,th {
	color: #135294;
}
-->
</style><BODY>
<table width="98%"  border="0" align="center" cellpadding="4" cellspacing="1" bgcolor="#799AE1">

<TR> 
    <TD height="5" colspan="12" bgcolor="#799AE1" align="center"><FONT color="#FFFFFF" style="font-size:14px">ģ �� �� ��</FONT></TD>
  </TR>
		<tr bgcolor="#FFFFFF">
		  <td>ѡ��ģ�巽����
		      <SELECT name="select" onChange="location=this.options[this.selectedIndex].value">
			<%Set Rs=Conn.Execute("Select ID,M_Name,M_Tacitly From DNJY_Template Where M_Reclaim=0 Order By ID")
			k=0
			do while not Rs.Eof
			Response.write "<OPTION value=""?id="&Rs(0)&""""
			If ID=Rs(0) Then Response.write " selected "
			Response.write ">"&Rs(1)&"</OPTION>"
			Rs.movenext
			k=k+1
			loop		
			rs.close
            set rs=nothing
            Call CloseDB(conn)%></select>
		      <input name="Submit1" type="submit" id="Submit1" value="��ΪĬ�Ϸ���" onClick="location='?action=Tacitly&ID=<%=ID%>'"/>
			  <%If k=1 Then
response.write "  <FONT color=""#ff0000"">(Ŀǰֻ��һ��ģ�巽�����Բ���ɾ����)</FONT>"
Else%>
		      <input name="Submit2" type="button" id="Submit2" onClick="{if(confirm('ȷ��Ҫ�����ģ�巽���������վ��\n�������վ��ģ�巽�����һأ�')){location='?action=Reclaim&id=<%=ID%>';}return false;}" value="�������վ"/>
<%End if%>
<br><FONT color="#ff0000">��ע��ģ�巽���ж�������ٷ�����ֻ����ѡ����һ��ģ���н��У����ѡ��ģ�巽����Ҫ�ٸ����������޷��ṩ������Ҫ�����շѡ�ѡ���������ΪĬ�ϼ��ɣ������Ŀɷ������վ��ɾ������</FONT>
			<form action="" method="post" onSubmit="if(T_Name.value=='' || T_Dir.value==''|| M_logo.value==''){alert('ģ�巽������ģ��Ŀ¼����ģ���ʶ������Ϊ��!');return false;};">
		      �½�ģ�巽������<input name="T_Name" type="text" id="T_Name" size="15"/> (���麺��)<br>�½�ģ��Ŀ¼����<input name="T_Dir" type="text" id="T_Dir" size="15" onkeyup="value=value.replace(/[^\w\.\/]/ig,'')"> (������Ӣ�������»���)<br>�½�ģ���ʶ����<input name="T_logo" type="text" id="T_logo" size="15"/> (��дӢ��)<br>
		      <input name="submit" type="submit" value="�½�ģ�淽��" /> ���½�ģ�巽��ʱ���ӵ�ǰĬ��ģ�巽���и����йش���Ͳ������½�������û����޸ļ��ɡ���
		  </form></td>
		</tr>
		<tr bgcolor="#FFFFFF">
		  <td><br>
		    <table width="98%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#D6DFF7">
			<tr>
              <td height="28" colspan="5" align="center" bgcolor="#FFFFFF"><b>��ǰ����ģ�巽������:<FONT color="#ff0000"><%=D_name%></FONT></b>&nbsp;&nbsp;<b>ģ���ʶ:</b><%=M_logo%></td>			  
            </tr>
            <tr>
              <td height="28" align="center" bgcolor="#FFFFFF"><b>ģ������</b></td>
			  <td height="28" align="center" bgcolor="#FFFFFF"><b>���ɵ��ļ�</b></td>
              <td align="center" bgcolor="#FFFFFF"><b>�޸Ĳ�����</b></td>
            </tr>

            <tr>
              <td height="28" align="center" bgcolor="#FFFFFF">ͷ��ģ��</td>
			  <td height="28" align="center" bgcolor="#FFFFFF">/<%=strInstallDir%>top.asp</td>
              <td align="center" bgcolor="#FFFFFF"><a href="M_top.asp?id=<%=ID%>">�޸Ĳ�����</a></td>
            </tr>
            <tr>
              <td height="28" align="center" bgcolor="#FFFFFF">β��ģ��</td>
			  <td height="28" align="center" bgcolor="#FFFFFF">/<%=strInstallDir%>Footer.asp</td>
              <td align="center" bgcolor="#FFFFFF"><a href="M_Footer.asp?id=<%=ID%>">�޸Ĳ�����</a></td>
            </tr>
            <tr>
              <td height="28" align="center" bgcolor="#FFFFFF">��ҳģ��</td>
			  <td height="28" align="center" bgcolor="#FFFFFF">/<%=strInstallDir%>Index.asp</td>
              <td align="center" bgcolor="#FFFFFF"><a href="M_Index.asp?id=<%=ID%>">�޸Ĳ�����</a></td>
            </tr>
            <tr>
              <td height="28" align="center" bgcolor="#FFFFFF">��ϢƵ��ҳģ��</td>
			  <td height="28" align="center" bgcolor="#FFFFFF">/<%=strInstallDir%>xxIndex.asp</td>
              <td align="center" bgcolor="#FFFFFF"><a href="M_xxIndex.asp?id=<%=ID%>">�޸Ĳ�����</a></td>
            </tr>
            <tr>
              <td height="28" align="center" bgcolor="#FFFFFF">��Ϣ��ʾģ��</td>
			  <td height="28" align="center" bgcolor="#FFFFFF">/<%=strInstallDir%>show.asp</td>
              <td align="center" bgcolor="#FFFFFF"><a href="M_show.asp?id=<%=ID%>">�޸Ĳ�����</a></td>
            </tr>
            <tr>
              <td height="28" align="center" bgcolor="#FFFFFF">��Ϣ��Ŀģ��</td>
			  <td height="28" align="center" bgcolor="#FFFFFF">/<%=strInstallDir%>list.asp</td>
              <td align="center" bgcolor="#FFFFFF"><a href="M_list.asp?id=<%=ID%>">�޸Ĳ�����</a></td>
            </tr>
            <tr>
              <td height="28" align="center" bgcolor="#FFFFFF">CSS���</td>
			  <td height="28" align="center" bgcolor="#FFFFFF">/<%=strInstallDir%>skin/<%=M_Dir%>/style.css</td>
              <td align="center" bgcolor="#FFFFFF"><a href="M_css.asp?id=<%=ID%>">�޸Ĳ�����</a></td>
            </tr>
            <!--<tr>
              <td height="28" align="center" bgcolor="#FFFFFF">������ʾģ��</td>
			  <td height="28" align="center" bgcolor="#FFFFFF">list.asp</td>
              <td align="center" bgcolor="#FFFFFF"><a href="M_News.asp?id=<%=ID%>">�޸�</a></td>
            </tr>
            <tr>
              <td height="28" align="center" bgcolor="#FFFFFF">������Ŀģ��</td>
			  <td height="28" align="center" bgcolor="#FFFFFF">News_list.asp</td>
              <td align="center" bgcolor="#FFFFFF"><a href="M_News_list.asp?id=<%=ID%>">�޸�</a></td>
            </tr>
            <tr>
              <td height="28" align="center" bgcolor="#FFFFFF">����ͼƬ���ߴ�</td>
			  <td height="28" align="center" bgcolor="#FFFFFF"> </td>
              <td align="center" bgcolor="#FFFFFF"><a href="adwh.asp?id=<%=ID%>">�޸�</a></td>
            </tr>--->
          </table>
	      <br></td>
  </tr>
</table>
</BODY>                                                                        
</HTML> 
