<!--#include file="../conn.asp"-->
<!--#include file="../config.asp"-->
<!--#include file="../SqlIn.Asp"-->
<!--#include file="../Include/Version.asp"-->
<%
'=====================================
'���칩����Ϣ��վ����ϵͳ
'��������Ƽ������ҳ�Ʒ
'�ͷ�����http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<style>
td{font-size:9pt;line-height:120%;color:#353535} 
body{font-size:9pt;line-height:120%} 

a:link          { color: #000000; text-decoration: none }
a:visited       { color: #000000; text-decoration: none }
a:active        { color: #000000; text-decoration: none }
a:hover         { color: #336699; text-decoration: none; position: relative; right: 0px; top: 1px }
</style>
<STYLE type=text/css>
.pad {
	PADDING-LEFT: 150px
}
</style>

<BODY>
<body bgcolor="#FFFFFF" leftmargin="4" topmargin="4" marginwidth="0" marginheight="0">
<TABLE WIDTH="98%" BORDER="0" ALIGN="center" CELLPADDING="0" CELLSPACING="0" height="18"> 
<TR BGCOLOR="#6699cc"> 
<TD WIDTH="40%"><FONT COLOR="#FFFFFF">���� ��ӭ</FONT>&nbsp;<b><%=request.cookies("administrator")%></b>&nbsp;<FONT COLOR="#FFFFFF">����������</FONT></TD>
<TD WIDTH="60%" ALIGN="RIGHT"><%Response.Write "<A HREF='http://www.ip126.com' TARGET='_blank'>Computer home&reg; INFO;  " & SystemVersion & " "&DatabaseType &" Build " & SystemBuildDate & "</a>"%>&nbsp;</TD>
</TR></TABLE>