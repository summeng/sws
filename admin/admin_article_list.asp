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
call checkmanage("06")%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<%Call OpenConn
If request("yz")="yes" Or request("yz")="no" Then'��ʾ������
Dim str2,i
Server.ScriptTimeout=50000
id=trim(request("DelID"))
if trim(id)="" then
response.write "<script language=JavaScript>" & chr(13) & "alert('û��ѡ���¼��');" &"window.location='admin_article_list.asp'" & "</script>"
response.end
end if
str2=split(id,",")
for i=0 to ubound(str2)
set rs=server.createobject("adodb.recordset")
sql="select ifshow from [DNJY_news] where id="&trim(str2(i))
rs.open sql,conn,1,3
If request("yz")="yes" Then
rs("ifshow")=1
Else
rs("ifshow")=0
End if
rs.update
Next
Call CloseRs(rs)
End If
If request("tj")="yes" Or request("tj")="no" Then'�Ƽ����
id=trim(request("ID"))
if trim(id)="" then
response.write "<script language=JavaScript>" & chr(13) & "alert('û��ѡ���¼��');" &"window.location='admin_article_list.asp'" & "</script>"
response.end
end if
set rs=server.createobject("adodb.recordset")
sql="select tuijian from [DNJY_news] where id="&trim(id)
rs.open sql,conn,1,3
If request("tj")="yes" Then
rs("tuijian")=1
Else
rs("tuijian")=0
End if
rs.update
Call CloseRs(rs)
End If
If request("gd")="yes" Or request("gd")="no" Then'�̶����
id=trim(request("ID"))
if trim(id)="" then
response.write "<script language=JavaScript>" & chr(13) & "alert('û��ѡ���¼��');" &"window.location='admin_article_list.asp'" & "</script>"
response.end
end if
set rs=server.createobject("adodb.recordset")
sql="select newstop from [DNJY_news] where id="&trim(id)
rs.open sql,conn,1,3
If request("gd")="yes" Then
rs("newstop")=1
Else
rs("newstop")=0
End if
rs.update
Call CloseRs(rs)
End If%>
<title>��������</title>
<link rel="stylesheet" href="../css/manage.css" type="text/css">
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
if (confirm("���Ҫɾ��������¼��ɾ�����޷��ָ���"))
window.location = params ;
}
function test()
{
  if(!confirm('�Ƿ�ȷ�����������������������ָܻ���')) return false;
}
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
<SCRIPT language=JavaScript>
function showoperatealert(id)
{
		if (id==1)
        {

		{
		  	thisForm.target='_self';
			thisForm.action="admin_article_list.asp?yz=yes";
			thisForm.submit();
		}
		}
		if (id==2)
        {

		{
		  	thisForm.target='_self';
			thisForm.action="admin_article_list.asp?yz=no";
			thisForm.submit();
		}
		}
		if (id==3)
        {

		{
		  	thisForm.target='_self';
			thisForm.action="admin_article_modi.asp?del=yes";
			thisForm.submit();
		}
		}
}

//-->
    </SCRIPT>
</head>
<BODY background="img/back.gif"><br>

<table width="98%" border="0" align="center" cellpadding="3" cellspacing="1" bgcolor="#000000">
<tr class=backs><td align="center" class=td height=18 colspan="22"><font color="#FFFFFF">������Ѷ����</font></td>
  </tr>
<tr bgcolor="#FFFFFF"> 
 <td colspan="6" width="100%"><form name="addNEWS" method="post" action="admin_article_list.asp?no=modi" >
����Ŀ��ѯ��<%
set rs=conn.execute("select * from DNJY_news_class where oneid>0 and twoid>0 and threeid=0 order by sorting")
%>
<SCRIPT language = "JavaScript">
var onecount2;
onecount2=0;
subcat2 = new Array();
        <%dim count2:count2 = 0
        do while not rs.eof%>
subcat2[<%=count2%>] = new Array("<%=rs("name")%>","<%=rs("oneid")%>","<%=rs("twoid")%>");
        <%count2 = count2 + 1
        rs.movenext
        loop
Call CloseRs(rs)%>
onecount2=<%=count2%>;
          </SCRIPT>
          <%set rs=conn.execute("select * from DNJY_news_class where oneid>0 and twoid>0 and threeid>0 order by sorting")%>
          <SCRIPT language = "JavaScript">
var onecount3;
onecount3=0;
subcat3 = new Array();
        <%dim count3:count3 = 0
        do while not rs.eof%>
subcat3[<%=count3%>] = new Array("<%=rs("name")%>","<%=rs("oneid")%>","<%=rs("twoid")%>","<%=rs("threeid")%>");
        <%count3 = count3 + 1
        rs.movenext
        loop
Call CloseRs(rs)%>
onecount3=<%=count3%>;

function changelocation2(locationid)
    {
    document.addNEWS.c_twoid.length = 0; 
    document.addNEWS.c_twoid.options[0] = new Option('������Ŀ','');
	document.addNEWS.c_threeid.length = 0; 
    document.addNEWS.c_threeid.options[0] = new Option('������Ŀ','');
    var locationid=locationid;
	var i;
    for (i=0;i < onecount2; i++)
        {
            if (subcat2[i][1] == locationid)
            { 
                document.addNEWS.c_twoid.options[document.addNEWS.c_twoid.length] = new Option(subcat2[i][0], subcat2[i][2]);
            }        
        }
       
    }    
     function changelocation3(locationid,locationid1)
    {
    document.addNEWS.c_threeid.length = 0; 
    document.addNEWS.c_threeid.options[0] = new Option('������Ŀ','');
    var locationid=locationid;
	 var locationid1=locationid1;
    var i;
    for (i=0;i < onecount3; i++)
        {
            if (subcat3[i][2] == locationid)
            { 
			if (subcat3[i][1] == locationid1)
			{
                document.addNEWS.c_threeid.options[document.addNEWS.c_threeid.length] = new Option(subcat3[i][0], subcat3[i][3]);
            }        
        }
       } 
    }            
	      </SCRIPT>
          <SELECT name="c_oneid" size="1" id="select8" onChange="changelocation2(document.addNEWS.c_oneid.options[document.addNEWS.c_oneid.selectedIndex].value)">
            <OPTION value="" selected>һ����Ŀ</OPTION>
            <%set rs=conn.execute("select * from DNJY_news_class where oneid>0 and twoid=0 order by sorting asc")
if rs.eof or rs.bof then
response.write "<option value=''>û����Ŀ</option>"
else
do until rs.eof
response.write "<option value='"&rs("oneid")&"'>"&rs("name")&"</option>"
    rs.movenext
    loop
end if
Call CloseRs(rs)%>
          </SELECT>
          <SELECT name="c_twoid" id="select9" onChange="changelocation3(document.addNEWS.c_twoid.options[document.addNEWS.c_twoid.selectedIndex].value,document.addNEWS.c_oneid.options[document.addNEWS.c_oneid.selectedIndex].value)">
            <OPTION value="" selected>������Ŀ</OPTION>
          </SELECT>
          <SELECT name="c_threeid" id="select10">
            <OPTION value="" selected>������Ŀ</OPTION>
          </SELECT>
        <input type="submit" name="Submit" value="�ύ" class="input">		
		
		</form>
 </td>
 

 <td width="30%" colspan="10">�����Ͳ�ѯ:
     <select name="selectm" onChange="var jmpURL=this.options[this.selectedIndex].value ; if(jmpURL!='') {window.location=jmpURL;} else {this.selectedIndex=0 ;}" target=_blank>
			    <OPTION value="admin_article_list.asp?recommend=""">����</OPTION>
				<OPTION value="admin_article_list.asp?recommend=0">ȫ��</OPTION> 
		        <OPTION value="admin_article_list.asp?recommend=1">�Ƽ�</OPTION>
                <OPTION value="admin_article_list.asp?recommend=2">�ö�</OPTION>
                <OPTION value="admin_article_list.asp?recommend=3">����</OPTION>                
                           
     </select> 
            </td>       
          </tr>


 <FORM name="thisForm" action="admin_article_modi.asp?del=yes" method="POST">
  <tr> 
    <td width="2%" height="25" align="center" bgcolor="#cccccc">ѡ��</td>
    <td width="25%" align="center" bgcolor="#cccccc">���±���</td>
    <td width="10%" align="center" bgcolor="#cccccc">��������</td>
	<td width="10%" align="center" bgcolor="#cccccc">������Ŀ</td>
    <td width="5%" align="center" bgcolor="#cccccc">����</td>
	<td width="5%" align="center" bgcolor="#cccccc">������</td>
    <td width="10%" align="center" bgcolor="#cccccc">����ʱ��</td>
    <td width="3%" align="center" bgcolor="#cccccc">�ö�</td>
    <td width="3%" align="center" bgcolor="#cccccc">�Ƽ�</td>
	<td width="5%" align="center" bgcolor="#cccccc">���ͼƬ����</td>
	<td width="5%" align="center" bgcolor="#cccccc">����</td>
    <td width="17%" align="center" bgcolor="#cccccc">����</td>
  </tr>
<%Dim c_oneid,c_twoid,c_threeid,selectm,selectkey,selectid,ProdId,comment,id,page,pages,j,photosl
page=clng(request("page"))
c_oneid=trim(request("c_oneid"))
c_twoid=trim(request("c_twoid"))
c_threeid=trim(request("c_threeid"))
If c_oneid="" then c_oneid=0
If c_twoid="" then c_twoid=0
If c_threeid="" then c_threeid=0
selectkey=trim(request.form("selectkey"))
selectm=trim(request.form("selectm"))
if selectkey="" then
selectkey=request.QueryString("selectkey")
end if
if selectm="" then
selectm=request.QueryString("selectm")
end if

set rs=server.createobject("adodb.recordset")
if Request.QueryString("recommend")=""   Then
sql="select * from DNJY_news  "
if (request("no")="modi" Or request("fy")="ok") And (request("c_oneid")<>"") Then
sql=sql&"where c_oneid="&c_oneid&""
If c_twoid>0 Then sql=sql&" and c_twoid="&c_twoid&""
If c_threeid>0 Then sql=sql&" and c_threeid="&c_threeid&""
End if
else
sql="select * from DNJY_news  "	
if Request.QueryString("recommend")=1 then
sql=sql&"where ifshow=1 and tuijian=1 "
Else
if Request.QueryString("recommend")=2 then
sql=sql&"where ifshow=1 and newstop=1 "
Else
if Request.QueryString("recommend")=3 then
sql=sql&"where  ifshow=0 "
Else
sql=sql
End If
End If
End If
End If


			select case selectm
			case ""
            sql=sql
		    case "name"
			sql=sql&" where title like '%"&selectkey&"%'"
			case "ProdId"
            If  Not isNumeric(selectkey) Then  
            response.write "���ӦΪ���֣�"     
            response.End
            else
			sql=sql&" where   Id =cint("&selectkey&")"
            End if
			case "content"
			sql=sql&"where  content like '%"&selectkey&"%'"			
		    end select


sql=sql&" order by id "&DN_OrderType&""
rs.open sql,conn,1,1
if rs.eof and rs.bof then
%>
<tr bgcolor="#FFFFFF"><td colspan="22" align="center">��ʱû����ؼ�¼</td>
</tr>
<%response.end
Else
rs.pagesize=15
if page=0 then page=1 
pages=rs.pagecount
if page > pages then page=pages
rs.absolutepage=page  
    
for j=1 to rs.pagesize 
If Not IsNull(rs("photo")) Then
photosl=UBound(Split(rs("photo"),"|"))
If photosl=-1 Then photosl=0
Else
photosl=0
End if

	  dim rs1,sql1
		set rs1=server.createobject("adodb.recordset") 
		sql1="select * from DNJY_news_pinglun where id="&rs("id")&" order by pinglunid "&DN_OrderType&""
		rs1.open sql1,conn,1,1%>
  <tr bgcolor="#ffffff"> 
    <td height="22" align="center"><input name="DelID" type="checkbox" id="DelID" value="<%=rs("id")%>"></td>	
    <td><a href="../news_show.asp?id=<%=rs("id")%>&n=<%=rs("c_oneid")%>,<%=rs("c_twoid")%>,<%=rs("c_threeid")%>" target="_blank" title="<%=rs("title")%>"><%=left(rs("title"),20)%></a><% if rs("img")=true then response.write "<img src='img/pic.gif' border=0 alt='ͼƬ����'>" end if %></td>
    <TD align="center"><%=rs("city_one")%><%If rs("city_two")<>"" Then response.write "/"&rs("city_two")%><%If rs("city_three")<>"" Then response.write "/"&rs("city_three")%></TD>
    <TD align="center"><%=rs("c_one")%><%If rs("c_two")<>"" Then response.write "/"&rs("c_two")%><%If rs("c_three")<>"" Then response.write "/"&rs("c_three")%></TD>
    <td align="center"><%=rs("newsuser")%></td>
	<td align="center"><%=rs("username")%></td>
    <td align="center"><%=rs("infotime")%></td>
    <td align="center">
	<%if rs("newstop")=1 then%><a href=?gd=no&id=<%=rs("id")%>><font color=red>��</font></a><%else%><a href=?gd=yes&id=<%=rs("id")%>>��</a><%end if%></td>
    <td align="center"><%if rs("tuijian")=1 then%><a href=?tj=no&id=<%=rs("id")%>><font color=red>��</font></a><%else%><a href=?tj=yes&id=<%=rs("id")%>>��</a><%end if%></td>
	<td align="center"><%=photosl%></td>
	<td align="center"><img onClick="javascript:left_menu('left_<%=rs("id")%>');" src="img/dedeexplode.gif" alt="չ������" width="11" height="11"> <%=rs1.recordcount%></td>
    <td align="center"><%if rs("ifshow")=1 then%><a href="?yz=no&DelID=<%=rs("id")%>">��ʾ</a><%else%><a href="?yz=yes&DelID=<%=rs("id")%>"><font color=red>����</font></a><%end if%> <a href="admin_article_modi.asp?id=<%=rs("id")%>">�޸�</a> <a href="admin_article_upload.asp?photoid=<%=rs("id")%>">�ϴ����</a> <%If rs("img")=true Then response.write"<a href='admin_article_edit.asp?id="&rs("id")&"'>"%>�������</a>
	<a href="javascript:DoEmpty('admin_article_modi.asp?DelID=<%=rs("id")%>&del=yes');">ɾ��</a></td>
  </tr> 
  <tr bgcolor="#ffffff" id="left_<%=rs("id")%>" style="display:none;">
    <td height="22" colspan="12" align="right">
	<table width="95%" border="0" cellpadding="2" cellspacing="1" bgcolor="#999999">
      <tr>
        <td width="4%" height="25" align="center" bgcolor="#cccccc">id</td>
        <td width="8%" align="center" bgcolor="#cccccc">�û�����</td>
        <td width="55%" align="center" bgcolor="#cccccc">��������</td>
        <td width="15%" align="center" bgcolor="#cccccc">����ʱ��</td>
        <td width="8%" align="center" bgcolor="#cccccc">ǰ̨��ʾ</td>
        <td width="8%" align="center" bgcolor="#cccccc">����</td>
      </tr>
	  <%if rs1.eof and rs1.bof then%>
		 <tr>
        <td height="25" colspan="12" align="center" bgcolor="#f8f8f8">��ʱ��û������</td>
        </tr>
		<%else
		do while not rs1.eof%>
      <tr>
        <td height="25" align="center" bgcolor="#f8f8f8"><%=rs1("pinglunid")%></td>
        <td align="center" bgcolor="#f8f8f8"><%=rs1("pinglunname")%></td>
        <td bgcolor="#f8f8f8"><%=rs1("pingluncontent")%></td>
        <td align="center" bgcolor="#f8f8f8"><%=rs1("pinglundate")%></td>
       <td align="center" bgcolor="#f8f8f8">
	<%if rs1("sh")=1 then%><a href=admin_article_pinglun.asp?action=shno&id=<%=rs1("pinglunid")%>>��ʾ</a><%else%><a href=admin_article_pinglun.asp?action=sh&id=<%=rs1("pinglunid")%>><font color=red>����</font></a><%end if%></td>
        <td align="center" bgcolor="#f8f8f8"><%if session("aleave")="check" then%>ɾ��<%else%><a href="javascript:DoEmpty('admin_article_pinglun.asp?action=del&id=<%=rs1("pinglunid")%>');">ɾ��</a><%end if%></td>
      </tr>     
	  <%rs1.movenext
		loop
		rs1.close
		set rs1=nothing
		end if%>
    </table>	</td>
  </tr>  
<%photosl=""
rs.movenext
if rs.eof then exit For
next%>
</table>
  <p align="center">
    <input type="checkbox" name="chkall" onClick="CheckAll(this.form)">
    ȫ��ѡ��
<input onclick=javascript:showoperatealert(1) type="submit" value="������ʾ" name="B1">
<input onclick=javascript:showoperatealert(2) type="submit" value="��������" name="B2">	  
<input name='submit1' type='submit' value=' ����ɾ�� ' onClick="return test();"  style="color: #FF0000">	
  </p>
</form>
<table width="98%" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr bgcolor="#ffffff"> 
<form method=post action="admin_article_list.asp?recommend=<%=Request.QueryString("recommend")%>">  
      <td height="30" align="center"> 
    <%if page<2 then      
    response.write "��ҳ ��һҳ&nbsp;"
  else
    response.write "<a href=admin_article_list.asp?page=1&recommend="&Request.QueryString("recommend")&"&id="&trim(request("id"))&"&c_oneid="&request("c_oneid")&"&c_twoid="&request("c_twoid")&"&c_threeid="&request("c_threeid")&"&fy=ok&selectm="&selectm&"&selectkey="&selectkey&">��ҳ</a>&nbsp;"
    response.write "<a href=admin_article_list.asp?page=" & page-1 & "&recommend="&Request.QueryString("recommend")&"&id="&trim(request("id"))&"&c_oneid="&request("c_oneid")&"&c_twoid="&request("c_twoid")&"&c_threeid="&request("c_threeid")&"&fy=ok&selectm="&selectm&"&selectkey="&selectkey&">��һҳ</a>&nbsp;"
  end if
  if rs.pagecount-page<1 then
    response.write "��һҳ βҳ"
  else
    response.write "<a href=admin_article_list.asp?page=" & (page+1) & "&recommend="&Request.QueryString("recommend")&"&id="&trim(request("id"))&"&c_oneid="&request("c_oneid")&"&c_twoid="&request("c_twoid")&"&c_threeid="&request("c_threeid")&"&fy=ok&selectm="&selectm&"&selectkey="&selectkey&">"
    response.write "��һҳ</a> <a href=admin_article_list.asp?page="&rs.pagecount&"&recommend="&Request.QueryString("recommend")&"&id="&trim(request("id"))&"&c_oneid="&request("c_oneid")&"&c_twoid="&request("c_twoid")&"&c_threeid="&request("c_threeid")&"&fy=ok&selectm="&selectm&"&selectkey="&selectkey&">βҳ</a>"
  end if
   response.write "&nbsp;ҳ�Σ�<strong><font color=red>"&page&"</font>/"&rs.pagecount&"</strong>ҳ "
    response.write "&nbsp;��<b><font color='#ff0000'>"&rs.recordcount&"</font></b>����¼ <b>"&rs.pagesize&"</b>����¼/ҳ"
   response.write " ת����<input type='text' name='page' size=4 maxlength=10 class=input value="&page&">"
   response.write " <input class=input type='submit'  value=' goto '  name='cndok'></span></p>"     
%>
      </td></form>
  </tr>
</table>
<%end if
Call CloseRs(rs)
Call CloseDB(conn)%>
<table width="98%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="#6699cc">
  <tr> 
    <td align="center">������ѯ</td>
  </tr>
  <tr bgcolor="#FFFFFF"> 
    <form name="form2" method="post" action="admin_article_list.asp">
      <td>
	  <table width="100%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="#CCCCCC">
          <tr bgcolor="#FFFFFF" align="center"> 
            <td > <input name="selectkey" type="text" id="selectkey" onFocus="this.value=''" value="������ؽ���" size="30"> 
             <select name="selectm" id="select">			   
                <OPTION VALUE="name">������</OPTION>                
                <OPTION VALUE="content">������</OPTION>
				<OPTION VALUE="ProdId">�����</OPTION> 
              </select><input type="submit" name="Submit" value="�� ѯ" ></td>
          </tr>
        </table></td>
    </form>
  </tr>
</table>

</body>
</html>
