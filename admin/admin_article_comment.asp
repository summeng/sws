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
<%call checkmanage("06")%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="../css/manage.css" type="text/css">
<BODY background="img/back.gif">
<script language=javascript src=../Include/mouse_on_title.js></script>
<script language=javascript>
<!--
function left_menu(meval)
{
  var left_n=eval(meval);
  if (left_n.style.display=="none")
  { eval(meval+".style.display='';"); }
  else
  { eval(meval+".style.display='none';"); }
}
function DoEmpty(params)
{
if (confirm("���Ҫɾ��������¼��ɾ����˼�¼����������ݶ�����ɾ�������޷��ָ���"))
window.location = params ;
}
function mm_jumpmenu(targ,selobj,restore){ //v3.0
  eval(targ+".location='"+selobj.options[selobj.selectedindex].value+"'");
  if (restore) selobj.selectedindex=0;
}
//document.all.left_sys.style.display='';
//document.all.left_bm.style.display='none'
-->
</script>
<script language="JavaScript">
<!--
function CheckAll(form)
{
for (var i=0;i<form.elements.length;i++)
{
var e = form.elements[i];
if (e.name != 'chkall')
e.checked = form.chkall.checked;
}
}

//-->
</script>
<style type="text/css">
<!--
body {
	margin-top: 5px;
}
-->
</style></head>
<body>
<table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="#FFFFFF">
<tr class=backs><td align="center" class=td height=18 colspan="11"><font color="#000000">�������۹���</font></td>
  </tr>
  <%dim page,rs2,pages,j,sql2,newsid
  page=clng(request("page"))
  Call OpenConn
		set rs=server.CreateObject("adodb.recordset") 
		rs.open "select * from DNJY_news_pinglun  order by pinglundate "&DN_OrderType&"",conn,1,1 
		if rs.eof and rs.bof then 
		response.write"<tr><td align=center  colspan=9>��ʱû���κ�����</td></tr>"
		response.end
		else
rs.pagesize=30
if page=0 then page=1 
pages=rs.pagecount
if page > pages then page=pages
rs.absolutepage=page
for j=1 to rs.pagesize 
%>
<form name="form1" method="post" action="admin_article_pinglun.asp?action=del&pl=yes">
<%
set rs2=Server.CreateObject("ADODB.Recordset")
sql2="select id,title from DNJY_News where id="&rs("id")
rs2.open sql2,conn,1,1%>

            <tr>
              <td width="5%" align="center" bgcolor="#CCCCCC"><img src="img/review.gif" width="16" height="16" align="absmiddle" /></td>
              <td width="5%" align="center" bgcolor="#CCCCCC">ѡ��</td>
              <td width="5%" align="center"><input name="ID" type="checkbox" id="ID" value="<%=rs("pinglunid")%>"></td>              
              <td width="5%" align="left" bgcolor="#CCCCCC">������</td>
              <td width="23%" align="left" bgcolor="#CCCCCC"><b><%=rs("pinglunname")%></b></td>
              <td width="10%" align="left" bgcolor="#CCCCCC">����ʱ��</td>
              <td width="18%" align="left" bgcolor="#CCCCCC"><%=rs("time")%></td>
              <td width="10%" align="left" bgcolor="#CCCCCC">������� <a href=../news_show.asp?catid=0&Id=<%=Cstr(rs2("Id"))%> target=_blank><img src="img/prin.gif" width="18" height="18" border=0 align="absmiddle"  alt="<%=rs2("title")%>"></a></td>
              <td width="5%" align="left" bgcolor="#CCCCCC"><a href=admin_article_comment.asp?Id=<%=Cstr(rs("pinglunid"))%> >�༭</a></td>
		      <td width="10%" align="left" bgcolor="#CCCCCC">���: <%if rs("sh")=1 then%><a href=admin_article_pinglun.asp?action=shno&id=<%=rs("pinglunid")%>><font color=red>��</font></a><%else%><a href=admin_article_pinglun.asp?action=sh&id=<%=rs("pinglunid")%>>��</a><%end if%></td>
              <td width="18%" align="left" bgcolor="#CCCCCC"><%if session("aleave")="check" then%>ɾ��<%else%><a href="javascript:DoEmpty('admin_article_pinglun.asp?action=del&id=<%=rs("pinglunid")%>');">ɾ��</a><%end if%></td>
            </tr>
            <tr>
              <td height="20" align="center" bgcolor="#F3F3F3"><img src="img/b.gif" width="10" height="10" /></td>
              <td colspan="14" align="left" bgcolor="#F3F3F3"><%=rs("pingluncontent")%></td>
            </tr>
<%rs.movenext
if rs.eof then exit For
Next%>
<tr>
<td height="20" colspan="14" align="center" bgcolor="#F3F3F3">	
  <p align="center">
    <input type="checkbox" name="chkall" onClick="CheckAll(this.form)">
    ȫ��ѡ�� 
    <input type="submit" name="Submit" value="ɾ����ѡ����" onClick="{if(confirm('�ò������ɻָ���\n\nȷʵɾ��ѡ�������ۣ�')){return true;}return false;}">
  </p>
 </td> 
</tr>
</form>	
       <tr valign="bottom"> 
            <td height="100%" colspan="14" align="center" > 
          <form method=post action="?newsid=<%= newsid %>">
            <%if page<2 then      
    response.write "��ҳ ��һҳ&nbsp;"
  else
    response.write "<a href=?newsid="&newsid&"&page=1>��ҳ</a>&nbsp;"
    response.write "<a href=?newsid="&newsid&"&page=" & page-1 & ">��һҳ</a>&nbsp;"
  end if
  if rs.pagecount-page<1 then
    response.write "��һҳ βҳ"
  else
    response.write "<a href=?newsid="&newsid&"&page=" & (page+1) & ">"
    response.write "��һҳ</a> <a href=?newsid="&newsid&"&page="&rs.pagecount&">βҳ</a>"
  end if
   response.write "&nbsp;ҳ�Σ�<strong><font color=red>"&page&"</font>/"&rs.pagecount&"</strong>ҳ "
    response.write "&nbsp;��<b><font color='#ff0000'>"&rs.recordcount&"</font></b>����¼ <b>"&rs.pagesize&"</b>����¼/ҳ"
   response.write " ת����<input type='text' name='page' size=4 maxlength=10 value="&page&">"
   response.write " <input class=input type='submit'  value=' goto '  name='cndok'></span></p>"     
%>
          </form>  </td> 
        </tr>
<%end if
Call CloseRs(rs)
Call CloseRs(rs2)
Call CloseDB(conn)%>
</table>
</body>
</html>
