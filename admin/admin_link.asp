<!--#include file="../conn.asp"-->
<!--#include file="../SqlIn.Asp"-->
<!--#include file=cookies.asp-->
<!--#include file="../config.asp"-->
<!--#include file="../Include/head.asp"-->
<!--#include file="../Include/Function.asp"-->
<!--#include file="../Include/label.asp"-->

<%
'=====================================
'���칩����Ϣ��վ����ϵͳ
'��������Ƽ������ҳ�Ʒ
'�ͷ�����http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>
<%call checkmanage("12")%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="../css/manage.css" type="text/css">
<BODY background="img/back.gif">

<script language=javascript>
<!--

function DoEmpty(params)
{

if (confirm("���Ҫɾ��������¼��ɾ����˼�¼����������ݶ�����ɾ�������޷��ָ���"))
window.location = params ;
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

<title>������������</title>
</head>
<body ><br>

<table width="98%" border="0" align="center" cellpadding="3" cellspacing="1" bgcolor="#000000">
<tr class=backs><td align="center" class=td height=18 colspan="15"><font color="#FFFFFF">�������ӹ���</font></td>
  </tr>


          <tr bgcolor="#FFFFFF"> 
 <td  colspan="4">
 <%dim key,img,mail,linktop,follows,objFSO,url,addtime,known,ss,dd,selectkey,selectm,alldel,Num,u,page,cc,oneid
	city_oneid=Request("city_one")
	city_twoid=Request("city_two")
	city_threeid=Request("city_three")
   If city_oneid="" Then city_oneid=0
   If city_twoid="" Then city_twoid=0
   If city_threeid="" Then city_threeid=0
   ss=trim(request("ss"))
   If ss="" Then ss=0
 key=Request("key")
 selectkey=trim(request("selectkey"))
selectm=trim(request("selectm"))
if selectkey="" then
selectkey=request.QueryString("selectkey")
end if
if selectm="" then
selectm=request.QueryString("selectm")
end if
id=trim(request("id"))

If key="del" then
if not isnumeric(id) or id="" then
response.write "<li>��������"
response.end
end If
logo=trim(request("logo"))
if left(logo,6)="images" then
Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
if objFSO.fileExists(Server.MapPath("../"&logo)) then
objFSO.DeleteFile(Server.MapPath("../"&logo))
end if
set objfso=nothing
end if
set rs=server.createobject("adodb.recordset")
sql="delete from [DNJY_link] where id="&cstr(id)
rs.open sql,conn,1,3
response.Write "<script language=javascript>alert('ɾ���ɹ�!');location.replace('?page="&page&"&ss="&ss&"&city_one="&city_oneid&"&city_two="&city_twoid&"&city_three="&city_threeid&"');</script>"
end if

If key="kg" then
known=trim(request("known"))
set rs=server.createobject("adodb.recordset")
sql="select * from [DNJY_link] where id="&trim(id)
rs.open sql,conn,1,3
If known="no" Then
rs("known")=0
Else
rs("known")=1
End if
rs.update
Call CloseRs(rs)
end If

If key="img" then
img=trim(request("img"))
set rs=server.createobject("adodb.recordset")
sql="select * from [DNJY_link] where id="&trim(id)
rs.open sql,conn,1,3
If img="no" Then
rs("img")=1
Else
rs("img")=0
End if
rs.update
Call CloseRs(rs)
end If

If key="top" then
linktop=trim(request("linktop"))
set rs=server.createobject("adodb.recordset")
sql="select * from [DNJY_link] where id="&trim(id)
rs.open sql,conn,1,3
If linktop="no" Then
rs("linktop")=0
Else
rs("linktop")=1
End if
rs.update
Call CloseRs(rs)
end If

If key="c" then
cc=trim(request("cc"))
set rs=server.createobject("adodb.recordset")
sql="select * from [DNJY_link] where id="&trim(id)
rs.open sql,conn,1,3
If cc="no" Then
rs("c")=0
Else
rs("c")=1
End if
rs.update
Call CloseRs(rs)
end If
%>
 <form method="POST" name="form" id="form" language="javascript"  action="?">
<% set rs=server.createobject("adodb.recordset")
set rs=conn.execute("select * from DNJY_city where id>0 and twoid>0 and threeid=0")%>
<SCRIPT language = "JavaScript">
var onecount;
onecount=0;
subcat = new Array();
        <%
	dim count5:count5 = 0
        do while not rs.eof 
        %>
subcat[<%=count5%>] = new Array("<%=rs("city")%>","<%=rs("id")%>","<%=rs("twoid")%>");
        <%
        count5 = count5 + 1
        rs.movenext
        loop
Call CloseRs(rs)
        %>
onecount=<%=count5%>;
          </SCRIPT>
          <%
set rs=conn.execute("select * from DNJY_city where id>0 and twoid>0 and threeid>0")
%>
          <SCRIPT language = "JavaScript">
var onecount4;
onecount4=0;
subcat4 = new Array();
        <%
		dim count4:count4 = 0
        do while not rs.eof 
        %>
subcat4[<%=count4%>] = new Array("<%=rs("city")%>","<%=rs("id")%>","<%=rs("twoid")%>","<%=rs("threeid")%>");
        <%
        count4 = count4 + 1
        rs.movenext
        loop
Call CloseRs(rs)
        %>
onecount4=<%=count4%>;

function changelocation(locationid)
    {
    document.form.city_two.length = 0; 
	document.form.city_two.options[0] = new Option('ѡ�����','');
	document.form.city_three.length = 0; 
	document.form.city_three.options[0] = new Option('ѡ�����','');
    var locationid=locationid;
    var i;
    for (i=0;i < onecount; i++)
        {
            if (subcat[i][1] == locationid)
            { 
                document.form.city_two.options[document.form.city_two.length] = new Option(subcat[i][0], subcat[i][2]);
            }        
        }
        
    }    
	
	function changelocation4(locationid,locationid1)
    {
    document.form.city_three.length = 0; 
    document.form.city_three.options[0] = new Option('ѡ�����','');
    var locationid=locationid;
	 var locationid1=locationid1;
    var i;
    for (i=0;i < onecount4; i++)
        {
            if (subcat4[i][2] == locationid)
            { 
			if (subcat4[i][1] == locationid1)
			{
                document.form.city_three.options[document.form.city_three.length] = new Option(subcat4[i][0], subcat4[i][3]);
            }        
        }
       } 
    }            
              </SCRIPT>
          <SELECT name="city_one" size="1" id="select" onChange="changelocation(document.form.city_one.options[document.form.city_one.selectedIndex].value)">
            <OPTION value="" selected>ѡ�����</OPTION>
            <%set rs=conn.execute("select * from DNJY_city where id>0 and twoid=0 order by indexid asc")
if rs.eof or rs.bof then
response.write "<option value=''>û�з���</option>"
else
do until rs.eof
response.write "<option value='"&rs("id")&"'>"&rs("city")&"</option>"
    rs.movenext
    loop
end if
Call CloseRs(rs)
%>
          </SELECT>
          <SELECT name="city_two" id="select6" onChange="changelocation4(document.form.city_two.options[document.form.city_two.selectedIndex].value,document.form.city_one.options[document.form.city_one.selectedIndex].value)">
            <OPTION value="" selected>ѡ�����</OPTION>
          </SELECT>
          <SELECT name="city_three" id="select7">
            <OPTION value="" selected>ѡ�����</OPTION>
          </SELECT>
 <input class="inputb" name="I2" type="submit" style="BACKGROUND-COLOR: #eeeeee; BORDER-BOTTOM: #a2a2a2 1px solid; BORDER-LEFT: #ffffff 1px solid; BORDER-RIGHT: #a2a2a2 1px solid; BORDER-TOP: #ffffff 1px solid; COLOR: #333333; FONT-SIZE: 9pt; HEIGHT: 19px;"  value="��ѯ" border="0" /></font>
</form></td> <td colspan="11"><a href="?ss=0&page=<%=Page%>&city_one=0&city_two=0&city_three=0">ȫ��������������</a>&nbsp;&nbsp;&nbsp;<a href="?ss=0&page=<%=Page%>&city_one=<%=city_oneid%>&city_two=<%=city_twoid%>&city_three=<%=city_threeid%>">��������������</a>&nbsp;&nbsp;&nbsp;&nbsp;<a href="?ss=1&page=<%=Page%>&city_one=<%=city_oneid%>&city_two=<%=city_twoid%>&city_three=<%=city_threeid%>&selectkey=<%=selectkey%>&selectm=<%=selectm%>">�ö�����վ</a>&nbsp;&nbsp;&nbsp;&nbsp;<a href="?ss=2&page=<%=Page%>&city_one=<%=city_oneid%>&city_two=<%=city_twoid%>&city_three=<%=city_threeid%>&selectkey=<%=selectkey%>&selectm=<%=selectm%>">��ɫ����վ</a>&nbsp;&nbsp;&nbsp;&nbsp;<a href="?ss=3&page=<%=Page%>&city_one=<%=city_oneid%>&city_two=<%=city_twoid%>&city_three=<%=city_threeid%>&selectkey=<%=selectkey%>&selectm=<%=selectm%>">���ص���վ</a>
    &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href="admin_linkedit.asp?adminlink=2">�������</a> </td>      
          </tr>

<form name="form1" method="post" action="admin_linkedit.asp?action=del">
  <tr> 
    <td width="3%" height="25" align="center" bgcolor="#cccccc">ѡ��</td>
	    <td width="5%" align="center" bgcolor="#cccccc">���</td>
    <td width="15%" align="center" bgcolor="#cccccc">��վ����</td>
    <td width="10%" align="center" bgcolor="#cccccc">��ַ</td>
    <td width="5%" align="center" bgcolor="#cccccc">LOGO</td>
    <td width="5%" align="center" bgcolor="#cccccc">��ʾ��ʽ</td>
    <td width="5%" align="center" bgcolor="#cccccc">�û���</td>
	<td width="12%" align="center" bgcolor="#cccccc">��������</td>
    <td width="10%" align="center" bgcolor="#cccccc">����</td>
    <td width="12%" align="center" bgcolor="#cccccc">���ʱ��</td>
	<td width="5%" align="center" bgcolor="#cccccc">����</td>
	<td width="2%" align="center" bgcolor="#cccccc">�ö�</td>
	<td width="2%" align="center" bgcolor="#cccccc">��ɫ</td>
    <td width="2%" align="center" bgcolor="#cccccc">�</td>
    <td width="10%" align="center" bgcolor="#cccccc">����</td>
  </tr>
<%Dim wzname,paixu

	IF city_oneid=0 then
	ttcclink=""
	elseIF city_threeid>0 then
	ttcclink=" city_oneid="&city_oneid&" and city_twoid="&city_twoid&" and city_threeid="&city_threeid
	elseIF city_twoid>0 then
	ttcclink=" city_oneid="&city_oneid&" and city_twoid="&city_twoid
	Else
	ttcclink=" city_oneid="&city_oneid
	End If
set rs=server.CreateObject("adodb.recordset")
sql = "select wzname,url,logo,img,addtime,paixu,id,known,name,city_one,city_two,city_three,linktop,mail,c from DNJY_link where wzname<>''"
If city_oneid>0 Then sql=sql&" and "&ttcclink&""
If ss=1 Then sql=sql&" and linktop=1"
If ss=2 Then sql=sql&" and c=1"
If ss=3 Then sql=sql&" and known=0"
			select case selectm
			case ""
            sql=sql
		    case "wzname"
			sql=sql&" and wzname like '%"&selectkey&"%'"
			case "url"
            sql=sql&" and url like '%"&selectkey&"%'"
            end Select
sql=sql&" order by id "&DN_OrderType&""           
rs.open sql,conn,1,1
  
		  if rs.EOF and rs.BOF Then
		  IF oneid>0 Then dd=conn.Execute("Select city From DNJY_city Where id="&city_oneid&" and twoid="&city_twoid&" and threeid="&city_threeid&"")(0)
          response.write"<tr bgcolor=#FFFFFF><td colspan='15'><p align='center'><font color='red'>����"&dd&"�������ӣ�</font></td></tr></table><br>"
		  follows=0
		  Else
Dim Pagesize,Allrecord,Allpage,ThisPage,k
if request("page")="" then
  ThisPage=1		
else
  ThisPage=request("page")
end if
rs.Pagesize=10 'ÿҳ��ʾ����
Pagesize=rs.Pagesize
Allrecord=rs.Recordcount
Allpage=rs.Pagecount
if ThisPage<1 then                           
ThisPage=1
end if
'On Error Resume Next
rs.move (ThisPage-1)*Pagesize
k=1
		  do while not rs.EOF
		  i=i+1
%>
  <tr bgcolor="#ffffff">
   <td height="22" align="center"><input name="DelID" type="checkbox" id="DelID" value="<%=rs("id")%>"></td>
    <td height="22" align="center"><%=i%></td>
    <td align="center"><a href="<%=trim(rs(1))%>" target=_blank><%if rs("c")=1 Then%><font color=#FF0000><%=trim(rs(0))%></font ><%else%><%=trim(rs(0))%><%
end If%></a><%if rs("linktop")=1 Then%><img src="../images/reg_ok.gif"  width="16"  height="16" border=0 alt="�ö�"><%End if%><%if rs("known")=0 Then%><img src="../images/note_error.gif" width="16"  height="16" border=0 alt="����"><%End if%></td>
    <td><%=trim(rs(1))%></td>
    <td align="center"><a href="<%=trim(rs(1))%>"  border=0 target=_blank><%If rs(2)<>"" And rs(2)<>"0" And Left(rs(2),7)<>"http://" Then%> <img src="../<%=rs(2)%>" border=1 width="88"  height="31"><%else%><%If Left(rs(2),7)="http://" then%><img src="<%=rs(2)%>" border=1 width="88"  height="31"><%else%><img src="../images/wu.gif" border=1 width="88"  height="31"><%End if%><%End if%></a></td>
    <td align="center"><%if rs(3)=0 then%><a href="?img=no&id=<%=rs("id")%>&key=img&page=<%=ThisPage%>&ss=<%=ss%>&city_one=<%=city_oneid%>&city_two=<%=city_twoid%>&city_three=<%=city_threeid%>">����</a><%else%><b><a href="?img=yes&id=<%=rs("id")%>&key=img&page=<%=ThisPage%>&ss=<%=ss%>&city_one=<%=city_oneid%>&city_two=<%=city_twoid%>&city_three=<%=city_threeid%>"><font color="#FF0000">ͼ��</font></a></b><%end if%></td>
	<td align="center"><%=trim(rs("name"))%></td>
	<td align="center"><a href="#" onClick="window.open('admin_link_mail.asp?mail=<%=trim(rs("mail"))%>&name=<%=trim(rs(0))%>&key=ok','Sample','toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=no,resizable=no,copyhistory=no,width=355,height=328,left=300,top=100')"><%=trim(rs("mail"))%></a></td>
     <TD align="center"><A title=<%=rs("city_one")%>/<%=rs("city_two")%>/<%=rs("city_three")%> href="#"><%=rs("city_one")%></A></TD>

    <td align="center"><%=rs(4)%></td>
    <td align="center"><%=(rs(5))%></td>

    <td align="center"><%if rs("linktop")=1 then%><a href="?linktop=no&id=<%=rs("id")%>&key=top&page=<%=ThisPage%>&ss=<%=ss%>&city_one=<%=city_oneid%>&city_two=<%=city_twoid%>&city_three=<%=city_threeid%>">��</a><%else%><b><a href="?linktop=yes&id=<%=rs("id")%>&key=top&page=<%=ThisPage%>&ss=<%=ss%>&city_one=<%=city_oneid%>&city_two=<%=city_twoid%>&city_three=<%=city_threeid%>"><font color="#FF0000">��</font></a></b><%end if%></td>

    <td align="center"><%if rs("c")=1 then%><a href="?cc=no&id=<%=rs("id")%>&key=c&page=<%=ThisPage%>&ss=<%=ss%>&city_one=<%=city_oneid%>&city_two=<%=city_twoid%>&city_three=<%=city_threeid%>">��</a><%else%><b><a href="?cc=yes&id=<%=rs("id")%>&key=c&page=<%=ThisPage%>&ss=<%=ss%>&city_one=<%=city_oneid%>&city_two=<%=city_twoid%>&city_three=<%=city_threeid%>"><font color="#FF0000">��</font></a></b><%end if%></td>

    <td align="center"><%if rs("known")=1 then%><a href="?known=no&id=<%=rs("id")%>&key=kg&page=<%=ThisPage%>&ss=<%=ss%>&city_one=<%=city_oneid%>&city_two=<%=city_twoid%>&city_three=<%=city_threeid%>">��</a><%else%><b><a href="?known=yes&id=<%=rs("id")%>&key=kg&page=<%=ThisPage%>&ss=<%=ss%>&city_one=<%=city_oneid%>&city_two=<%=city_twoid%>&city_three=<%=city_threeid%>"><font color="#FF0000">��</font></a></b><%end if%></td>
	<td align="center"><a href="#" onClick="window.open('admin_linkedit.asp?adminlink=1&id=<%=rs("id")%>','Sample','toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=no,resizable=no,copyhistory=no,width=600,height=650,left=300,top=100')">�޸�</a> 
	<%if session("aleave")="check" then%>ɾ��<%else%><a href="javascript:DoEmpty('?key=del&id=<%=rs("id")%>&logo=<%=rs("logo")%>&page=<%=ThisPage%>&ss=<%=ss%>&city_one=<%=city_oneid%>&city_two=<%=city_twoid%>&city_three=<%=city_threeid%>');">ɾ��</a><%end if%></td>
  </tr>
 
       <%rs.MoveNext
	   if k>=Pagesize then exit Do
	   k=k+1
          loop
          follows=rs.RecordCount
          end If
          Call CloseRs(rs)
		  conn.close:Set conn=nothing%>
</table>
  <p align="center">
  <input type="checkbox" name="chkall" onClick="CheckAll(this.form)">
    ȫ��ѡ�� 
    <input type="submit" name="Submit" value="ɾ����ѡ����" onClick="{if(confirm('�ò������ɻָ���\n\nȷʵɾ��ѡ�������£�')){return true;}return false;}">
  </p>
</form>
 <table width="98%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="#6699cc">
<tr> 
<td width="595" height="25" align="center">
  ����&nbsp;<font color="#CC5200"><%=Allrecord%></font>&nbsp;����¼ �� <font color="#CC5200"><%=Allpage%></font> ҳ �����ǵ� <font color="#CC5200"><%=ThisPage%></font> ҳ
<%
if ThisPage<2 then     
response.write "<font color=""#808080"">��ҳ</font>&nbsp;"
response.write "<font color=""#808080"">��һҳ</font>&nbsp;"     
else     
response.write "<a href=?page=1&ss="&ss&"&city_one="&city_oneid&"&city_two="&city_twoid&"&city_three="&city_threeid&"&selectkey="&selectkey&"&selectm="&selectm&">��ҳ</a>&nbsp;"
response.write "<a href=?page="&ThisPage-1&"&ss="&ss&"&city_one="&city_oneid&"&city_two="&city_twoid&"&city_three="&city_threeid&"&selectkey="&selectkey&"&selectm="&selectm&">��һҳ</a>&nbsp;"     
end if
if Allpage-ThisPage<1 then     
response.write "<font color=""#808080"">��һҳ</font>&nbsp;"
response.write "<font color=""#808080"">βҳ</font>&nbsp;"  
else     
response.write "<a href=?page="&(ThisPage+1)&"&ss="&ss&"&city_one="&city_oneid&"&city_two="&city_twoid&"&city_three="&city_threeid&"&selectkey="&selectkey&"&selectm="&selectm&">��һҳ</a>&nbsp;"   
response.write "<a href=?page="&Allpage&"&ss="&ss&"&city_one="&city_oneid&"&city_two="&city_twoid&"&city_three="&city_threeid&"&selectkey="&selectkey&"&selectm="&selectm&">βҳ</a>&nbsp;"     
end if
response.write "ÿҳ"&Pagesize&"����¼"%></td>
</tr>
      </table>   

<table width="98%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="#6699cc">
  <tr> 
    <td align="center" colspan="3">������ѯ</td>
  </tr>
  <tr bgcolor="#FFFFFF"> 
    <form name="form2" method="post" action="?page=<%=Page%>&city_one=<%=city_oneid%>&city_two=<%=city_twoid%>&city_three=<%=city_threeid%>">
      <td>
	  <table width="100%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="#CCCCCC">
          <tr bgcolor="#FFFFFF" align="center"> 
            <td width="35%"> <input name="selectkey" type="text" id="selectkey" onFocus="this.value=''" value="������ؽ���" size="30"> 
             <select name="selectm" id="select">                
                <OPTION VALUE="wzname">����վ����</OPTION> 
                <OPTION VALUE="url">����ַ</OPTION>
              </select> <input type="submit" name="Submit" value="�� ѯ" ></td>
          </tr>
        </table></td>
    </form>
  </tr>
</table>

</body>
</html>
