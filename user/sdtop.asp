<!--#include file="../conn.asp"-->
<!--#include file="../config.asp"-->
<!--#include file="../SqlIn.Asp"-->
<!--#include file="../Include/filename.asp"-->
<!--#include file="../Include/head.asp"-->
<!--#include file="../Include/Function.asp"-->
<!--#include file="../Include/label.asp"-->
<!--#include file="../Include/ContentAutoPage.asp"-->
<%
'=====================================
'���칩����Ϣ��վ����ϵͳ
'��������Ƽ������ҳ�Ʒ
'�ͷ�����http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>
<%Call OpenConn
dim com,k,sdfg,username,vip,jf
com=CheckStr(trim(request("com")))
set rs = Server.CreateObject("ADODB.RecordSet")
sql="select * from [DNJY_user] where username='"&com&"'"
rs.open sql,conn,1,3
if rs.bof or rs.eof then
response.redirect "/"
end If
If rs("sdkg")<>1 Then
response.write "û�д������δ����ˣ�<a href=""/"&strInstallDir&"?"&C_ID&""">����ҳ</a>"
response.end
End if
sdfg=rs("sdfg")
vip=rs("vip")
jf=rs("jf")
username=rs("username")
if rs("vip")<>"1" then
response.redirect "../General_stores.asp?com="&com&"&"&C_ID&""
response.end
end if
rs("tong")=rs("tong")+1
rs.update
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<title><%=rs("sdname")%></title>
<meta name="keywords" content="<%=rs("sdname")%>" />
<meta name="description" content="<%=CutStr(rs("sdjj"),200,"...")%>" />
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel='stylesheet' type='text/css' href='../images/sdimg/<%=sdfg%>/style.css' />
</head>
<body style="margin:70px;padding:0px;text-align:center;background-color:#ffffff;">
<table width="960" border="0" cellspacing="0" cellpadding="0" align="center" >
  <tr>
    <td>
<div id='f_top'>
<div id='f_global_menu' style="width:100%px;">
<div id='f_tmenu'>
<table cellpadding='0' cellspacing='0' border='0' width="100%">
<tr><td align="right" valign="middle">������ 
 <script language="JavaScript">
today=new Date();
function initArray(){
this.length=initArray.arguments.length
for(var i=0;i<this.length;i++)
this[i+1]=initArray.arguments[i]  }
var d=new initArray(
" ������",
" ����һ",
" ���ڶ�",
" ������",
" ������",
" ������",
" ������");
document.write(
"<font color=#000000 style='font-size:11px;font-family: verdana'> ",
today.getFullYear(),"��",
today.getMonth()+1,"��",
today.getDate(),"��",
d[today.getDay()+1],
"</font>" ); 
</script>
<script type="text/javascript">
function doZoom(size)//����ѡ��
{document.getElementById('zoom').style.fontSize=size+'px';}
</script>
</td>
<td width='50'> </td>
<td>
<table cellpadding='0' cellspacing='0' border='0'>
<col width='400'><col width='1' align='absmiddle'><col><col width='1' align='absmiddle'>
<col><col width='1' align='absmiddle'><col width='60' align='center'>
<tr>
<td valign="middle">
<form method="Get" name="form" action="sdsearch.asp?<%=C_ID%>">
��Ʒ������<input type="text" name="key"  size=30 value="�ؼ���" maxlength="50" onFocus="this.select();">
<select name="otype" class="input" id="otype" style="line-height:30px;">
<option value="title" selected="selected" class="input">��Ʒ����</option>
<option value="msg" class="input">��Ʒ���</option>
</select>
<input type="hidden" name="com" value="<%=com%>">
<input type="submit" name="Submit"  value="����" height="10"></td></form>
<td><img src='../images/sdimg/<%=sdfg%>/icon_line.gif'></td>

<td width='65' align='center'><a href="collection_shops.asp?scid=<%=rs("username")%>&name=<%=rs("name")%>&sdname=<%=rs("sdname")%>&mpname=<%=rs("mpname")%>" target="_blank">��Ա�ղ�</a></td>
<td><img src='../images/sdimg/<%=sdfg%>/icon_line.gif'></td><td width='60' align='center'><a href="../">������վ</a></td><td><img src='../images/sdimg/<%=sdfg%>/icon_line.gif'></td><td width='60' align='center'><a href="../<%=reg%>">�������</a></td><td><img src='../images/sdimg/<%=sdfg%>/icon_line.gif'></td><td width='60' align='center'><a href="../user.asp">�������</a></td><td><img src='../images/sdimg/<%=sdfg%>/icon_line.gif'></td>
</td>
<td>&nbsp;&nbsp;</td>
</tr>
</table>
</td>
</tr>
</table>
</div>
</div>
</div>
</div>
</td>
</tr>
</table>
<table width="960" height="160" border="0" align="center" cellpadding="0" cellspacing="0" background="<%If rs("sdBanner")<>"" then%>../<%=rs("sdBanner")%><%else%>../images/sdimg/Banner.jpg<%End if%>">
  <tr>
    <td width="800"><div style='line-height: 300%;width: 100%; font-size:28pt; font-family: ����;margin-left:50px;'><b><span id="colora"><%=rs("sdname")%></span></b></div><br><span id="colorb" style='margin-left:50px;'>��Ҫ��Ӫ��<%=CutStr(rs("zhuying"),100,"...")%></span></td>    
  </tr>
</table>
<table width="960" height="40" border="0" align="center" cellpadding="0" cellspacing="0" background="../images/sdimg/<%=sdfg%>/bg-nav.jpg">
  <tr>
    <td><div id="nav"><ul>
	
		 <td align="center"><a href="com.asp?com=<%=com%>" class="menu_txt"><span id="colormenu">��˾��ҳ</span></a></td>
		 <td align="center" class="menu_txts"><span id="colormenu">|</span></td>
		 <td align="center"><a href="sdjj.asp?com=<%=com%>" class="menu_txt"><span id="colormenu">��������</span></a></td>
		 <td align="center" class="menu_txts"><span id="colormenu">|</span></td>	
         <td align="center"><a href="sdxx.asp?com=<%=com%>" class="menu_txt" id="menu4" style="color:#FFFFFF;" ><span id="colormenu">��Ʒ����</span></a></td>
         <td align="center" class="menu_txts"><span id="colormenu">|</span></td>
<%Dim rsfl,sqlfl
set rsfl=server.createobject("adodb.recordset")
sqlfl="select * From DNJY_userNewsClass where dhshow="&DN_True&" order by showid"
rsfl.open sqlfl,conn,1,1
if rsfl.eof then
response.write"<td align=""center""><span id=""colormenu"">���޷���|</span></td>"
Else
do while not rsfl.eof 
response.write"<td align=""center""><a href=""UserNews_list.asp?com="&com&"&classid="&rsfl("id")&""" class=""menu_txt""><span id=""colormenu"">"&rsfl("classname")&"</span></a></td>"
response.write"<td align=""center"" class=""menu_txts"" ><span id=""colormenu"">|</span></td>"		 
rsfl.movenext
Loop
rsfl.close
set rsfl=nothing
end if%>
<td align="center" style="color:#FFFFFF;"><a href="Userlink_list.asp?com=<%=com%>" class="menu_txt"><span id="colormenu">��������</span></a></td>		 
<td align="center" style="color:#FFFFFF" ></td>
	</ul></div></td>
  </tr>
</table>
