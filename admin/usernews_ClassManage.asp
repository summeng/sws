<!--#include file="../conn.asp"-->
<!--#include file="../SqlIn.Asp"-->
<!--#include file="../config.asp"-->
<!--#include file=cookies.asp-->
<%
'=====================================
'���칩����Ϣ��վ����ϵͳ
'��������Ƽ������ҳ�Ʒ
'�ͷ�����http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>
<%call checkmanage("04")%>
<html>
<head>
<link rel="stylesheet" type="text/css" href="../css/css.css">
<script language="JavaScript" type="text/JavaScript">

function ConfirmDelBig()
{
   if(confirm("ȷ��Ҫɾ������Ŀ��ɾ������Ŀͬʱ��ɾ�������µ����м�¼�����Ҳ��ָܻ���"))
     return true;
   else
     return false;
	 
}


</script>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>��Ŀ����</title>
<style type="text/css">
<!--
body {
	background-color: #E3EBF9;
}
.style1 {color: #FFFFFF;
	font-weight: bold;
}
.style2 {color: #FFFFFF}
-->
</style>
</head>
<%

dim rsClass,rsSmallClass,ErrMsg
set rsClass=server.CreateObject("adodb.recordset")
rsClass.open "Select * From DNJY_userNewsClass",conn,1,3

%>
<body>
<div align="center">
      <br>
<table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#ADAED6" width="100%">

  <TR> 
    <TD height="5" colspan="12" bgcolor="#799AE1" align="center"><FONT color="#FFFFFF" style="font-size:14px">���е��̵Ļ�Ա���·������</FONT></TD>
  </TR>
      <a href="usernews_ClassAdd.asp"><strong><font color="#FF0000" size="2"><u>�����Ŀ</u></font></strong></a>
      </p>
      
     <table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#ADAED6" width="100%" height="30">
        <tr >
          <td width="25%" class="style1"><div align="center">��Ŀ����</div></td>
          <td width="12%" class="style1"><div align="center">ID��</div></td>
          <td width="13%" class="style1"><div align="center">������ʾ���</div></td>
          <td width="13%" class="style1"><div align="center">������ʾ</div></td>
          <!--td width="17%" class="style1"><div align="center">��ҳ��Ŀ��Ŀ��ʾ</div></td-->
          <td width="20%" class="style1"><div align="center">����ѡ��</div></td>
        </tr>
		
		<%do while not rsClass.eof%>
		 
        <tr>
          <td ><%=rsClass("ClassName")%></td>
          <td ><div align="center"><%=rsClass("ID")%></div></td>
          <td ><div align="center"><%=rsClass("showid")%></div></td>
          <td ><div align="center">
		  <%if rsClass("dhshow")=true then %>
		  		<a href='usernews_classshow.asp?act=dh&sh=0&id=<%=rsClass("id")%>'><font color="#339933">��</font></a> 
	 <% else%>
	  <a href='usernews_classshow.asp?act=dh&sh=1&id=<%=rsClass("id")%>'><font color="#CC0000">��</font></a> 
			<%end if
				
		  %></div></td>
          <!--td >
		  <div align="center">
		  <%'if rsClass("lmshow")=true then %>
		  		<a href='usernews_classshow.asp?act=lm&sh=0&id=<%=rsClass("id")%>'><font color="#339933">��</font></a> 
	 <%' else%>
	  <a href='usernews_classshow.asp?act=lm&sh=1&id=<%=rsClass("id")%>'><font color="#CC0000">��</font></a> 
			<%'end if
				
		  %></div></td-->
          <td ><div align="center"><a href="usernews_ClassModify.asp?ClassID=<%=rsClass("ID")%>">�޸�</a> | <a href="usernews_ClassDel.asp?ClassID=<%=rsClass("ID")%>"onClick="return ConfirmDelBig();">ɾ��</a>&nbsp;</div></td>
        </tr>
		<%
	  
	  rsClass.movenext
	loop
%>
  </table>
      <table class="list1" width="666" border="0" cellspacing="1" cellpadding="0">
        <tr>
          <td width="666">
            <div align="left">&nbsp;&nbsp;������ǡ��򡰷񡱿ɸı���ʾ״̬<br>
            ��������ʾ��š�����Ŀ�����ڵ����������д������Ϊ���֣����ԽСԽ��ǰ����
              <br>
              ��������ʾ��״̬���ǡ��ڵ�������ʾ����Ŀ���ƣ�״̬�����ڵ���������ʾ����Ŀ����<br>
          <!--����ҳ��Ŀ��Ŀ��ʾ��״̬���ǡ�����ҳ������Ŀ����ʾ����Ŀ��Ϣ��״̬��������ҳ������Ŀ�в���ʾ����Ŀ��Ϣ--></div></td>
        </tr>
  </table>
      <p>&nbsp;</p>
</div>
<%
rsClass.close
set rsClass=nothing
%>
</table>
</body>
</html>
