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
<%call checkmanage("05")%>
<link rel="stylesheet" type="text/css" href="../css/css.css">
<BODY>
<style type="text/css">
<!--
body {
	background-color: #E3EBF9;
	margin-top: 16px;
}
-->
</style>
<body leftmargin="0">
<script language=javascript>
//�����͹���ظ���ʼ
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

<!---������忪ʼ--->
<script>
function loadThreadFollow(t_id,b_id){
	var targetImg =eval("document.all.followImg" + t_id);
	var targetDiv =eval("document.all.follow" + t_id);
	
	if ("object"==typeof(targetImg)){
		if (targetDiv.style.display!=''){
			targetDiv.style.display="";
			
		}else{
			targetDiv.style.display="none";
		}
	}
}
</script>
<!---����������--->
<!---ȫѡ��֤ɾ����ʼ--->
<SCRIPT language=JavaScript>
function showoperatealert(id)
{
		if (id==1)
        {

		{
		  	thisForm.target='_self';
			thisForm.action="info_yz.asp?qxyz=yes";
			thisForm.submit();
		}
		}
		if (id==2)
        {

		{
		  	thisForm.target='_self';
			thisForm.action="info_yz.asp?qxyz=no";
			thisForm.submit();
		}
		}
		if (id==3)
        {

		{
		  	thisForm.target='_self';
			thisForm.action="info_xxtj.asp?tj=1";
			thisForm.submit();
		}
		}
		if (id==4)
        {

		{
		  	thisForm.target='_self';
			thisForm.action="info_xxtj.asp?tj=2";
			thisForm.submit();
		}
		}
		if (id==5)
        {

		{
		  	thisForm.target='_self';
			thisForm.action="info_tjdel.asp";
			thisForm.submit();
		}
		}		
		if (id==6)
        {
		{
			
		  	thisForm.target='_self';
			thisForm.action="info_del.asp?urlpage=infolist.asp";
			thisForm.submit();
		}
		}						
}

function CheckAll(form)
  {
  for (var i=0;i<form.elements.length;i++)
    {
    var e = form.elements[i];
    if (e.name != 'chkall' )
       e.checked = form.chkall.checked;
    }
  }
//-->
    </SCRIPT>

<!---ȫѡ��֤ɾ������--->
<div align="center">
  <center>
<table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="99%" height="1" bordercolor="#ADAED6">
  <tr>
    <td width="100%" height="24" colspan="14">
    <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%" height="25">
<TR>
    <TD height="20" bgcolor="#799AE1" align="center" colspan="14"><FONT color="#FFFFFF" style="font-size:14px">�� �� �� Ϣ �� ��</FONT></TD>
 </TR>
      <tr>
        <td width="200" height="24">
    <FORM name=sou1 action="?dick=1" method=POST>
    <p align="center">���ʺţ�<input type="text" name="key1" size="15">
    <input type="submit" value="����" name="B1">
    </form>
        </td>
        <td width="200" height="24">
    <FORM name=sou2 action="?dick=2" method=POST>
    <p align="center">��������<input type="text" name="key2" size="15">
    <input type="submit" value="����" name="B1">
    </form>
        </td>
        <td width="360">
    <FORM name=sou3 action="?dick=3" method=POST>
    <p align="center">���ؼ��ʣ�<input type="text" name="key3" size="28">
	<select name="skey" style="BACKGROUND-COLOR: #FFCCFF; BORDER-BOTTOM: #FFFFFF 1px solid; BORDER-LEFT: #a2a2a2 1px solid; BORDER-RIGHT: #ffffff 1px solid; BORDER-TOP: #a2a2a2 1px solid; COLOR: #333333; FONT-SIZE: 9pt; HEIGHT: 19px;">
                            <option value="0">������</option>
                            <option value="1">�����</option>                            
                          </select>
    <input type="submit" value="����" name="B1">
    </form>
        </td>
      </tr>
	  <tr>
</table>
    <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%" height="25">
<%
dim k,dick,tj
dim ThisPage,Pagesize,Allrecord,Allpage

   dick=trim(request("dick"))
	city_oneid=Request("city_one")
	city_twoid=Request("city_two")
	city_threeid=Request("city_three")
   If city_oneid="" Then city_oneid=0
   If city_twoid="" Then city_twoid=0
   If city_threeid="" Then city_threeid=0

	IF city_oneid=0 then
	ttcc=""
	elseIF city_threeid>0 then
	ttcc=" and city_oneid="&city_oneid&" and city_twoid="&city_twoid&" and city_threeid="&city_threeid
	elseIF city_twoid>0 then
	ttcc=" and city_oneid="&city_oneid&" and city_twoid="&city_twoid
	Else
	ttcc=" and city_oneid="&city_oneid
	End If
	
	type_oneid=Request("type_one")
	type_twoid=Request("type_two")
	type_threeid=Request("type_three")
   If type_oneid="" Then type_oneid=0
   If type_twoid="" Then type_twoid=0
   If type_threeid="" Then type_threeid=0

	IF type_oneid=0 then
	tttt=""
	elseIF type_threeid>0 then
	tttt=" and type_oneid="&type_oneid&" and type_twoid="&type_twoid&" and type_threeid="&type_threeid
	elseIF type_twoid>0 then
	tttt=" and type_oneid="&type_oneid&" and type_twoid="&type_twoid
	Else
	tttt=" and type_oneid="&type_oneid
	End If%>
	<td height="25" width="20%">����������ѯ��</td>
        <td width="50%" colspan="4" height="24">
 <form method="POST" name="form" id="form" language="javascript"  action="?t1=<%=trim(request("t1"))%>&t2=<%=trim(request("t2"))%>&t3=<%=trim(request("t3"))%>&dick=<%=dick%>">
<% 
set rs=server.createobject("adodb.recordset")
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
</form> </td><%If city_oneid>0 then%><td width="30%" colspan="2" align="left"><%city_one=Conn.Execute("select city from DNJY_city where twoid=0 and id="&city_oneid)(0)
	If city_twoid>0 Then city_two=Conn.Execute("select city from DNJY_city where id="&city_oneid&" and threeid=0 and twoid="&city_twoid)(0)
	If city_threeid<>0 Then city_three=Conn.Execute("select city from DNJY_city where id="&city_oneid&" and twoid="&city_twoid&" and threeid="&city_threeid)(0)
	response.write ""&city_one&"/"&city_two&"/"&city_three&""%></td><%End if%>
      </tr>

<tr>
<form method="POST" name="myform" id="myform" language="javascript"  action="?t1=<%=trim(request("t1"))%>&t2=<%=trim(request("t2"))%>&t3=<%=trim(request("t3"))%>&dick=<%=dick%>">
<td height="25" width="20%">����Ϣ����ѯ��</td>
<td height="25" width="50%"><%
set rs=conn.execute("select * from DNJY_type where id>0 and twoid>0 and threeid=0 order by indexid")
%>
          <SCRIPT language = "JavaScript">
var onecount2;
onecount2=0;
subcat2 = new Array();
        <%
		dim count2:count2 = 0
        do while not rs.eof 
        %>
subcat2[<%=count2%>] = new Array("<%=rs("name")%>","<%=rs("id")%>","<%=rs("twoid")%>");
        <%
        count2 = count2 + 1
        rs.movenext
        loop
Call CloseRs(rs)
        %>
onecount2=<%=count2%>;
          </SCRIPT>
          <%
set rs=conn.execute("select * from DNJY_type where id>0 and twoid>0 and threeid>0 order by indexid")
%>
          <SCRIPT language = "JavaScript">
var onecount3;
onecount3=0;
subcat3 = new Array();
        <%dim count3:count3 = 0
        do while not rs.eof%>
subcat3[<%=count3%>] = new Array("<%=rs("name")%>","<%=rs("id")%>","<%=rs("twoid")%>","<%=rs("threeid")%>");
        <%
        count3 = count3 + 1
        rs.movenext
        loop
Call CloseRs(rs)
        %>
onecount3=<%=count3%>;

function changelocation2(locationid)
    {
    document.myform.type_two.length = 0; 
    document.myform.type_two.options[0] = new Option('��������','');
	document.myform.type_three.length = 0; 
    document.myform.type_three.options[0] = new Option('��������','');
    var locationid=locationid;
	var i;
    for (i=0;i < onecount2; i++)
        {
            if (subcat2[i][1] == locationid)
            { 
                document.myform.type_two.options[document.myform.type_two.length] = new Option(subcat2[i][0], subcat2[i][2]);
            }        
        }
       
    }    
     function changelocation3(locationid,locationid1)
    {
    document.myform.type_three.length = 0; 
    document.myform.type_three.options[0] = new Option('��������','');
    var locationid=locationid;
	 var locationid1=locationid1;
    var i;
    for (i=0;i < onecount3; i++)
        {
            if (subcat3[i][2] == locationid)
            { 
			if (subcat3[i][1] == locationid1)
			{
                document.myform.type_three.options[document.myform.type_three.length] = new Option(subcat3[i][0], subcat3[i][3]);
            }        
        }
       } 
    }            
	      </SCRIPT>
          <SELECT name="type_one" size="1" id="select8" onChange="changelocation2(document.myform.type_one.options[document.myform.type_one.selectedIndex].value)">
            <OPTION value="" selected>һ������</OPTION>
            <%set rs=conn.execute("select * from DNJY_type where id>0 and twoid=0 order by indexid asc")
if rs.eof or rs.bof then
response.write "<option value=''>û�з���</option>"
else
do until rs.eof
response.write "<option value='"&rs("id")&"'>"&rs("name")&"</option>"
    rs.movenext
    loop
end if
Call CloseRs(rs)%>
          </SELECT>
          <SELECT name="type_two" id="select9" onChange="changelocation3(document.myform.type_two.options[document.myform.type_two.selectedIndex].value,document.myform.type_one.options[document.myform.type_one.selectedIndex].value)">
            <OPTION value="" selected>��������</OPTION>
          </SELECT>
          <SELECT name="type_three" id="select10">
            <OPTION value="" selected>��������</OPTION>
          </SELECT>  <input class="inputb" name="t2" type="submit" style="BACKGROUND-COLOR: #eeeeee; BORDER-BOTTOM: #a2a2a2 1px solid; BORDER-LEFT: #ffffff 1px solid; BORDER-RIGHT: #a2a2a2 1px solid; BORDER-TOP: #ffffff 1px solid; COLOR: #333333; FONT-SIZE: 9pt; HEIGHT: 19px;"  value="��ѯ" border="0" /></font>
</form></td><%If type_oneid>0 then%><td width="30%" colspan="2" align="left"><%type_one=Conn.Execute("select name from DNJY_type where twoid=0 and id="&type_oneid)(0)
	If type_twoid>0 Then type_two=Conn.Execute("select name from DNJY_type where id="&type_oneid&" and threeid=0 and twoid="&type_twoid)(0)
	If type_threeid<>0 Then type_three=Conn.Execute("select name from DNJY_type where id="&type_oneid&" and twoid="&type_twoid&" and threeid="&type_threeid)(0)
	response.write ""&type_one&"/"&type_two&"/"&type_three&""%></td><%End if%>
                        </tr>
						

      <tr>
        <td width="100%" colspan="14" height="24">
        <p align="center"><b>��ʾ��</b><a href="?dick=4&city_one=<%=city_oneid%>&city_two=<%=city_twoid%>&city_three=<%=city_threeid%>&type_one=<%=type_oneid%>&type_two=<%=type_twoid%>&type_three=<%=type_threeid%>" title='ָ��δ��֤����Ϣ'>δ(<font color="#FF0000">��</font>)��֤��Ϣ</a>---<a href="?dick=5&city_one=<%=city_oneid%>&city_two=<%=city_twoid%>&city_three=<%=city_threeid%>&type_one=<%=type_oneid%>&type_two=<%=type_twoid%>&type_three=<%=type_threeid%>" title='ָ�Ѿ���֤����Ϣ'>��(��)��֤��Ϣ</a>---<a href="?dick=14&city_one=<%=city_oneid%>&city_two=<%=city_twoid%>&city_three=<%=city_threeid%>&type_one=<%=type_oneid%>&type_two=<%=type_twoid%>&type_three=<%=type_threeid%>" title='ָ�Ѿ��ö�����Ϣ'>���ö���Ϣ</a>---<a href="?dick=6&city_one=<%=city_oneid%>&city_two=<%=city_twoid%>&city_three=<%=city_threeid%>&type_one=<%=type_oneid%>&type_two=<%=type_twoid%>&type_three=<%=type_threeid%>" title='ָ�Ѿ��Ƽ�����Ϣ'>���Ƽ���Ϣ</a>---<a href="?dick=7&city_one=<%=city_oneid%>&city_two=<%=city_twoid%>&city_three=<%=city_threeid%>&type_one=<%=type_oneid%>&type_two=<%=type_twoid%>&type_three=<%=type_threeid%>" title='ָ��Ч���ѵ�����Ϣ'>�ѵ�����Ϣ</a>---<a href="?dick=15&city_one=<%=city_oneid%>&city_two=<%=city_twoid%>&city_three=<%=city_threeid%>&type_one=<%=type_oneid%>&type_two=<%=type_twoid%>&type_three=<%=type_threeid%>"  title='�����������'>������Ϣ</a>---<a href="?dick=16&city_one=<%=city_oneid%>&city_two=<%=city_twoid%>&city_three=<%=city_threeid%>&type_one=<%=type_oneid%>&type_two=<%=type_twoid%>&type_three=<%=type_threeid%>" title='ָ�����Ϊ0����Ϣ'>������Ϣ</a>---<a href="?dick=8&city_one=<%=city_oneid%>&city_two=<%=city_twoid%>&city_three=<%=city_threeid%>&type_one=<%=type_oneid%>&type_two=<%=type_twoid%>&type_three=<%=type_threeid%>" title='ָ����3���������ҵ����Ϊ0����Ϣ'>3������������Ϣ</a>---<a href="?dick=17&city_one=<%=city_oneid%>&city_two=<%=city_twoid%>&city_three=<%=city_threeid%>&type_one=<%=type_oneid%>&type_two=<%=type_twoid%>&type_three=<%=type_threeid%>" title='ָ��Ϣ��Чʱ���ѹ��ҵ����Ϊ0����Ϣ'>����������Ϣ</a><br><a href="?dick=10&city_one=<%=city_oneid%>&city_two=<%=city_twoid%>&city_three=<%=city_threeid%>&type_one=<%=type_oneid%>&type_two=<%=type_twoid%>&type_three=<%=type_threeid%>" title='ָ����˵ľ�����Ϣ'>������Ϣ</a>---<a href="?dick=9&city_one=<%=city_oneid%>&city_two=<%=city_twoid%>&city_three=<%=city_threeid%>&type_one=<%=type_oneid%>&type_two=<%=type_twoid%>&type_three=<%=type_threeid%>" title='ָ����˵ľ�����Ϣ'>���뾺����Ϣ</a>---<a href="?dick=11&city_one=<%=city_oneid%>&city_two=<%=city_twoid%>&city_three=<%=city_threeid%>&type_one=<%=type_oneid%>&type_two=<%=type_twoid%>&type_three=<%=type_threeid%>" title='ָ�ѵ��ڵľ�����Ϣ'>���ھ�����Ϣ</a>---<a href="?dick=12&city_one=<%=city_oneid%>&city_two=<%=city_twoid%>&city_three=<%=city_threeid%>&type_one=<%=type_oneid%>&type_two=<%=type_twoid%>&type_three=<%=type_threeid%>" title='ָ�Ѿ�����html�ļ�����Ϣ'>������html��Ϣ</a>---<a href="?dick=13&city_one=<%=city_oneid%>&city_two=<%=city_twoid%>&city_three=<%=city_threeid%>&type_one=<%=type_oneid%>&type_two=<%=type_twoid%>&type_three=<%=type_threeid%>" title='ָ��δ����html�ļ�����Ϣ'>δ����html��Ϣ</a>---<a href="?dick=19&city_one=<%=city_oneid%>&city_two=<%=city_twoid%>&city_three=<%=city_threeid%>&type_one=<%=type_oneid%>&type_two=<%=type_twoid%>&type_three=<%=type_threeid%>" title='ָ��δ����html�ļ�����Ϣ'>��ͨ��Ա�Ķ�</a>---<a href="?dick=20&city_one=<%=city_oneid%>&city_two=<%=city_twoid%>&city_three=<%=city_threeid%>&type_one=<%=type_oneid%>&type_two=<%=type_twoid%>&type_three=<%=type_threeid%>" title='ָ��δ����html�ļ�����Ϣ'>VIP��Ա�Ķ�</a>---<a href="?city_one=<%=city_oneid%>&city_two=<%=city_twoid%>&city_three=<%=city_threeid%>&type_one=<%=type_oneid%>&type_two=<%=type_twoid%>&type_three=<%=type_threeid%>" title='ָ�������鿴��ȫ����Ϣ'>������ȫ����Ϣ</a>---<a href="?dick=" title='ָȫ������Ϣ'>ȫ����Ϣ</a><br>&nbsp;&nbsp;&nbsp;&nbsp;<a href="info_html.asp" title='��ת��html���ɹ���ҳ��'><b>����html����</b></a>&nbsp;&nbsp;&nbsp;&nbsp;<a href="info_del_dqxx.asp" onClick="return confirm('ȷ��Ҫ��չ�����Ϣ��\n\n��ͬʱɾ�����ͼƬ���ļ�����\n\n�ò������ɻָ���')" ><b>һ��ɾ��������Ϣ</b></a>&nbsp;&nbsp;&nbsp;&nbsp;<a href="#" ONCLICK="window.open('info_replyxx.asp','Sample','toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=yes,resizable=no,copyhistory=no,width=570,height=200,left=300,top=20')" title='������һ������'><b>ȡ�����ھ�����Ϣ�ʸ�</b></a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
<%set rs=server.createobject("adodb.recordset")
sql = "select * from [DNJY_info] where yz<>1"
rs.open sql,conn,1,1
if rs.eof then
response.write "<font color=#0000FF>Ŀǰû�д������Ϣ��</font>"
Else
response.write "<b><img src='../img/notice.gif' width=15 height=15 border=0><font color=#FF0000>�������д������Ϣ��</font></b>"
response.write "<bgsound src='../img/msg.wav' loop='1'>"
end If
Call CloseRs(rs)%>
 </td>
      </tr>
    </table>
    </td>
  </tr>
  <FORM name=thisForm method=POST action="info_del.asp?urlpage=infolist.asp">
<%
function HTMLEncode(fString)
if not isnull(fString) then
    fString = replace(fString, ">", "&gt;")
    fString = replace(fString, "<", "&lt;")
    fString = Replace(fString, CHR(32), "&nbsp;")
    fString = Replace(fString, CHR(9), "&nbsp;")
    fString = Replace(fString, CHR(34), "&quot;")
    fString = Replace(fString, CHR(39), "&#39;")
    fString = Replace(fString, CHR(13), "")
    fString = Replace(fString, CHR(10) & CHR(10), "</P><P> ")
    fString = Replace(fString, CHR(10), "<BR> ")
    HTMLEncode = fString
end if
end Function


set rs = Server.CreateObject("ADODB.RecordSet")
Select Case dick
Case "1"
sql="select * from [DNJY_info] where username='"&trim(request("key1"))&"'"&ttcc&""&tttt&" order by fbsj "&DN_OrderType&"" 
Case "2"
sql="select * from [DNJY_info] where name like '%"&trim(request("key2"))&"%'"&ttcc&""&tttt&" order by fbsj "&DN_OrderType&"" 
Case "3"
If request("skey")=0 then
sql="select * from [DNJY_info] where (biaoti like '%"&trim(request("key3"))&"%')"&ttcc&""&tttt&" order by fbsj "&DN_OrderType&"" 
Else
sql="select * from [DNJY_info] where (memo like '%"&trim(request("key3"))&"%')"&ttcc&""&tttt&" order by fbsj "&DN_OrderType&""
End if
Case "4"
sql="select * from [DNJY_info] where yz<>1"&ttcc&""&tttt&" order by fbsj "&DN_OrderType&"" 
Case "5"
sql="select * from [DNJY_info] where yz=1"&ttcc&""&tttt&" order by fbsj "&DN_OrderType&""
Case "6"
sql="select * from [DNJY_info] where yz=1 and tj>0"&ttcc&""&tttt&" order by fbsj "&DN_OrderType&""
Case "7"
sql="select * from [DNJY_info] where "&DN_Now&">=dqsj"&ttcc&""&tttt&" order by fbsj "&DN_OrderType&""
Case "8"
sql="select * from [DNJY_info] where DateDiff("&DN_DatePart_D&",fbsj,"&DN_Now&")>=90"&ttcc&""&tttt&" and llcs=0 order by fbsj "&DN_OrderType&""
Case "9"
sql="select * from [DNJY_info] where hfxx=2"&ttcc&""&tttt&" order by fbsj "&DN_OrderType&""
Case "10"
sql="select * from [DNJY_info] where (hfxx=1 or hfxx=2)"&ttcc&""&tttt&" order by fbsj "&DN_OrderType&""
Case "11"
sql="select * from [DNJY_info] where hfxx=1"&ttcc&""&tttt&" and "&DN_Now&">=dqsj order by fbsj "&DN_OrderType&""
Case "12"
sql="select * from [DNJY_info] where fsohtm=1"&ttcc&""&tttt&" order by fbsj "&DN_OrderType&""
Case "13"
sql="select * from [DNJY_info] where fsohtm=0"&ttcc&""&tttt&" order by fbsj "&DN_OrderType&""
Case "14"
sql="select * from [DNJY_info] where b>0"&ttcc&""&tttt&" order by b "&DN_OrderType&""
Case "15"
sql="select * from [DNJY_info] where llcs>0"&ttcc&""&tttt&" order by llcs "&DN_OrderType&""
Case "16"
sql="select * from [DNJY_info] where llcs=0"&ttcc&""&tttt&" order by fbsj"
Case "17"
sql="select * from [DNJY_info] where "&DN_Now&">=dqsj"&ttcc&""&tttt&" and llcs=0 order by fbsj "&DN_OrderType&""
Case "18"
sql="select * from [DNJY_info] where username='"&trim(request("key1"))&"'"&ttcc&""&tttt&" order by fbsj "&DN_OrderType&"" 
Case "19"
sql="select * from [DNJY_info] where Readinfo=1"&ttcc&""&tttt&" order by fbsj "&DN_OrderType&""
Case "20"
sql="select * from [DNJY_info] where Readinfo=2"&ttcc&""&tttt&" order by fbsj "&DN_OrderType&""
Case Else
sql="select * from [DNJY_info] where fbts>0"&ttcc&""&tttt&" order by fbsj "&DN_OrderType&"" 
End Select



rs.open sql,conn,1,1
if rs.eof then
response.write "<table width=""96%""><tr><td width=""100%""  colspan=""12"" align=""center""><br><li>û�м�¼</table></td></tr>"
response.end
end if
if request("page")="" then
  ThisPage=1		
else
  ThisPage=request("page")
end if
rs.Pagesize=20'ÿҳ����
Pagesize=rs.Pagesize
Allrecord=rs.Recordcount
Allpage=rs.Pagecount
if ThisPage<1 then                           
ThisPage=1
end if
On Error Resume Next
rs.move (ThisPage-1)*Pagesize
%>
  <tr bgcolor="#BDBEDE">
    <td width="5%" height="24" align="center" style="border-bottom-style: solid; border-bottom-width: 1" bordercolor="#666666">ѡ��</td>
    <td width="20%" height="28" bordercolor="#666666" bgcolor="#BDBEDE" style="border-bottom-style: solid; border-bottom-width: 1">����</td>
    <td width="10%" height="24" style="border-bottom-style: solid; border-bottom-width: 1" bordercolor="#666666">
    <p align="center">ע��ID/����</td>
    <td width="16%" height="24" style="border-bottom-style: solid; border-bottom-width: 1" bordercolor="#666666">
    <p align="center">���ʱ��/����ʱ��</td>
    <td width="5%" height="24" align="center" style="border-bottom-style: solid; border-bottom-width: 1" bordercolor="#666666">�Ķ�Ȩ��</td>
	<td width="5%" height="24" align="center" style="border-bottom-style: solid; border-bottom-width: 1" bordercolor="#666666">���</td>
	<td width="5%" height="24" align="center" style="border-bottom-style: solid; border-bottom-width: 1" bordercolor="#666666">�ظ�</td>
	<td width="5%" height="24" align="center" style="border-bottom-style: solid; border-bottom-width: 1" bordercolor="#666666">�ٱ�</td>
    <td width="4%" height="24" style="border-bottom-style: solid; border-bottom-width: 1" bordercolor="#666666">
    <p align="center">��֤</td>
    <td width="4%" height="24" style="border-bottom-style: solid; border-bottom-width: 1" bordercolor="#666666">
    <p align="center">�ö�</td>
    <td width="4%" height="24" style="border-bottom-style: solid; border-bottom-width: 1" bordercolor="#666666">
    <p align="center">�Ƽ�</td>
    <td width="5%" height="24" style="border-bottom-style: solid; border-bottom-width: 1" bordercolor="#666666">
    <p align="center">����|����</td>
	<td width="4%" height="24" style="border-bottom-style: solid; border-bottom-width: 1" bordercolor="#666666">
    <p align="center">������</td>  
    <td width="4%" height="24" bordercolor="#666666" bgcolor="#BDBEDE" style="border-bottom-style: solid; border-bottom-width: 1"><div align="center">����</div></td>
  </tr>
<%dim uid
k=0
do while not rs.eof
uid=rs("id")

	    dim rs1,sql1
		set rs1=server.createobject("adodb.recordset") 
		sql1="select * from DNJY_reply where xxid="&rs("id")&" order by id "&DN_OrderType&""
		rs1.open sql1,conn,1,1
		
		dim rs2,sql2
		set rs2=server.createobject("adodb.recordset") 
		sql2="select * from DNJY_Bad_info where infoid="&rs("id")&" order by id "&DN_OrderType&""
		rs2.open sql2,conn,1,1%>
  <tr>
    <td width="5%" height="24" align="center" bgcolor="#FFFFFF" style="border-top-style: solid; border-top-width: 1">
    <INPUT type=checkbox value="<%=uid%>" name=selectedid></td>

    <td width="20%" height="24" bgcolor="#FAFAFA" style="border-top-style: solid; border-top-width: 1">
   <% if rs("tupian")<>"0" then
response.write "<img src=""../images/num/pic.gif"">" 
end if%>
<%	belongtype=""&rs("type_one")&""
	IF rs("type_two")<>"" and not isnull(rs("type_two")) Then belongtype=belongtype&" / "&rs("type_two")&""
	IF rs("type_three")<>"" and not isnull(rs("type_three")) Then belongtype=belongtype&" / "&rs("type_three")&""
	%>
<A href="admin_info.asp?id=<%=uid%>" target="_blank" alt="<%=rs("biaoti")%><br>���<%=belongtype%>|��Ч�ڣ�<%=DateDiff("d",Now(),""&rs("dqsj")&"")%>��"><%=CutStr(rs("biaoti"),36,"...")%></a></td>
    <td width="10%" height="24" align="center" bgcolor="#FFFFFF" style="border-top-style: solid; border-top-width: 1"><a href="?dick=18&key1=<%= rs("username")%>&city_one=<%=city_oneid%>&city_two=<%=city_twoid%>&city_three=<%=city_threeid%>&type_one=<%=type_oneid%>&type_two=<%=type_twoid%>&type_three=<%=type_threeid%>" title='��<%=rs("username")%>�����Ĺ�����Ϣ'><%=rs("username")%><font color="#FF0000">|</font><%=rs("name")%></a></td>
    <td width="15%" height="24" align="center" bgcolor="#FAFAFA" style="border-top-style: solid; border-top-width: 1"><%=rs("addsj")%><br><font color="#FF0000">/</font><%If Now()>=rs("dqsj") then%><font color="#FF0000"><%=rs("fbsj")%></font><%else%><%=rs("fbsj")%><%End if%></td>
<td width="4%" height="24" align="center" bgcolor="#FFFFFF" style="border-top-style: solid; border-top-width: 1"><%if rs("Readinfo")=0 then%>�ο�<%else%><%if rs("Readinfo")=1 then%><font color="#0000FF">��ͨ��Ա</font><%else%><font color="#FF0000">VIP��Ա</font><%end if%><%end if%></td>
	<td width="4%" height="24" align="center" bgcolor="#FFFFFF" style="border-top-style: solid; border-top-width: 1"><%=rs("llcs")%></td>
	<td width="5%" align="center" bgcolor="#FFFFFF"><img onClick="javascript:left_menu('left_<%=rs("id")%>');" src="img/dedeexplode.gif" alt="չ���ظ�" width="11" height="11"> <%if rs1.recordcount>0 then%><font color="#FF0000"><%=rs1.recordcount%></font><%else%><%=rs1.recordcount%><%End if%></td>
	<td width="5%" align="center" bgcolor="#FFFFFF"><img onClick="javascript:left_menu('left2_<%=rs("id")%>');" src="img/dedeexplode.gif" alt="չ���ٱ�" width="11" height="11"> <%if rs2.recordcount>0 then%><font color="#FF0000"><%=rs2.recordcount%></font><%else%><%=rs2.recordcount%><%End if%></td>
    <td width="4%" height="24" align="center" bgcolor="#FFFFFF" style="border-top-style: solid; border-top-width: 1"><%if rs("yz")=1 then%><a href="info_yz.asp?selectedid=<%=uid%>&qxyz=no&page=<%=(ThisPage)%>&t1=<%=trim(request("t1"))%>&t2=<%=trim(request("t2"))%>&t3=<%=trim(request("t3"))%>&dick=<%=dick%>&city_one=<%=city_oneid%>&city_two=<%=city_twoid%>&city_three=<%=city_threeid%>&type_one=<%=type_oneid%>&type_two=<%=type_twoid%>&type_three=<%=type_threeid%>">��</a><%else%><b><a href="info_yz.asp?selectedid=<%=uid%>&qxyz=yes&page=<%=(ThisPage)%>&t1=<%=trim(request("t1"))%>&t2=<%=trim(request("t2"))%>&t3=<%=trim(request("t3"))%>&dick=<%=dick%>&city_one=<%=city_oneid%>&city_two=<%=city_twoid%>&city_three=<%=city_threeid%>&type_one=<%=type_oneid%>&type_two=<%=type_twoid%>&type_three=<%=type_threeid%>"><font color="#FF0000">��</font></a></b><%end if%></td>
	<td width="4%" height="24" align="center" bgcolor="#FFFFFF" style="border-top-style: solid; border-top-width: 1"><%=rs("b")%></td>
    <td width="4%" height="24" align="center" bgcolor="#FFFFFF" style="border-top-style: solid; border-top-width: 1"><%if rs("tj")=0 then%>δ��<%else%><%if rs("tj")=1 then%><font color="#FF0000">��Ϣ</font><%else%><font color="#FF0000">ͼƬ</font><%end if%><%end if%></td>
    <td width="5%" height="24" align="center" bgcolor="#FFFFFF" style="border-top-style: solid; border-top-width: 1"><%if rs("hfxx")=1 then%><font color="#0000FF"><span alt="���۾���<%=FormatNumber(rs("hfje")/rs("fbts"),2,-1)%><br>����ʣ<b><%=Fix(rs("dqsj")-now())%></b>��" style="cursor:help">��</span>|<a href="info_gl.asp?selectedid=<%=uid%>&gl=jj&qxyz=no&page=<%=(ThisPage)%>&t1=<%=trim(request("t1"))%>&t2=<%=trim(request("t2"))%>&t3=<%=trim(request("t3"))%>&dick=<%=dick%>&city_one=<%=city_oneid%>&city_two=<%=city_twoid%>&city_three=<%=city_threeid%>&type_one=<%=type_oneid%>&type_two=<%=type_twoid%>&type_three=<%=type_threeid%>">��</a></font><%else%><%if rs("hfxx")=2 then%><font color="#FF0000">����<%=FormatNumber(rs("hfje")/rs("fbts"),2,-1)%></font><%else%>��|<a href="info_gl.asp?selectedid=<%=uid%>&gl=jj&qxyz=yes&page=<%=(ThisPage)%>&t1=<%=trim(request("t1"))%>&t2=<%=trim(request("t2"))%>&t3=<%=trim(request("t3"))%>&dick=<%=dick%>&city_one=<%=city_oneid%>&city_two=<%=city_twoid%>&city_three=<%=city_threeid%>&type_one=<%=type_oneid%>&type_two=<%=type_twoid%>&type_three=<%=type_threeid%>"><font color="#FF0000">��</font></a><%end if%><%end if%></td> 
    <td width="4%" height="24" align="center" bgcolor="#FFFFFF" style="border-top-style: solid; border-top-width: 1"><%if rs("fsohtm")=1 then%><span title=�ļ�λ�ã�/<%=strInstallDir%>html/<%=rs("id")%>.html style="cursor:help">��</span><%else%><b><font color="#FF0000">��</font></a></b><%end if%></td>    
    <td width="4%" height="24" bgcolor="#FAFAFA" style="border-top-style: solid; border-top-width: 1">
    <p align="center"><font color="#FF0000"><span id="followImg<%=k%>" style="CURSOR: hand" onClick="loadThreadFollow(<%=k%>,5)">���</span></font></td>
  </tr>
  <tr style="display:none" id="follow<%=k%>">
    <td width="100%" height="24" colspan="14" style="border-bottom-style: solid; border-bottom-width: 1" bgcolor="#FFF8EC">
    <div align="center">
      <center>
    <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%" height="100%">
      <tr>
        <td bgcolor="#CCCCFF"><div align="right">IP:<%= rs("ip") %> <a href='info_edit.asp?id=<%=uid%>'>�༭����Ϣ</a> <a href="#" ONCLICK="window.open('user_mail.asp?username=<%=rs("username")%>&email=<%=rs("email")%>','Sample','toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=yes,resizable=no,copyhistory=no,width=700,height=700,left=300,top=100')">�����ʼ�</a>  <a href='message2.asp?username=<%=trim(rs("username"))%>'>���Ͷ���</a></font>         
          <%if rs("yz")=0 then%>
              <a href="info_yz.asp?selectedid=<%=uid%>">ͨ����֤</a>
              <%else%>
 <a href="#" ONCLICK="window.open('info_tj.asp?id=<%=uid%>','Sample','toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=no,resizable=no,copyhistory=no,width=250,height=150,left=300,top=100')">��Ϣ�Ƽ�</a>                
              <a href="info_yz.asp?selectedid=<%=uid%>&qxyz=no">ȡ����֤</a>              
              <%end if%>
			  <a href="javascript:DoEmpty('info_del.asp?selectedid=<%=uid%>&urlpage=infolist.asp');">ɾ������Ϣ</a> <a href="javascript:DoEmpty('info_replydel.asp?selectedid=<%=uid%>');">ɾ�����лظ�</a>
            &nbsp;</div></td>
        </tr>
    </table>
      </center>
    </div>
    </td>
  </tr>
  
  <tr bgcolor="#ffffff" id="left_<%=rs("id")%>" style="display:none;">
    <td height="22" colspan="14" align="right">
	<table width="95%" border="0" cellpadding="2" cellspacing="1" bgcolor="#999999">
      <tr>
        <td width="4%" height="25" align="center" bgcolor="#cccccc">id</td>
        <td width="8%" align="center" bgcolor="#cccccc">�û�����</td>
        <td width="45%" align="center" bgcolor="#cccccc">�ظ�����</td>
        <td width="13%" align="center" bgcolor="#cccccc">�ظ�ʱ��</td>
		 <td width="15%" align="center" bgcolor="#cccccc">IP��ַ</td>
        <td width="5%" align="center" bgcolor="#cccccc">�༭</td>
        <td width="5%" align="center" bgcolor="#cccccc">���</td>
        <td width="5%" align="center" bgcolor="#cccccc">����</td>
      </tr>
	  <%if rs1.eof and rs1.bof then%>
		 <tr>
        <td height="25" colspan="14" align="center" bgcolor="#E3EBF9">��ʱ��û�лظ�</td>
        </tr>
		<%else
		do while not rs1.eof%>
      <tr>
        <td height="25" align="center" bgcolor="#E3EBF9"><%=rs1("id")%></td>
        <td align="center" bgcolor="#E3EBF9"><%if rs1("username")<>"" then%><a target="_blank" href=../preview.asp?username=<%=rs1("username")%>><%=rs1("username")%><%else%>����<%end if%></td>
        <td bgcolor="#E3EBF9"><%=rs1("neirong")%></td>
        <td align="center" bgcolor="#E3EBF9"><%=rs1("hfsj")%></td>
		<td align="center" bgcolor="#E3EBF9"><%=rs1("ip")%></td>
		<td align="center" bgcolor="#E3EBF9"><a href="chkcomment.asp?edit=replay&id=<%=rs1("id")%>">�༭</a></td>		
       <td align="center" bgcolor="#E3EBF9">
	<%if rs1("hfyz")=1 then%><a href="info_replyyz.asp?dick=delyz&selectedid=<%=rs1("id")%>">��</a><%else%><b><a href="info_replyyz.asp?selectedid=<%=rs1("id")%>"><font color="#FF0000">��</font></a></b><%end if%></td>
        <td align="center" bgcolor="#E3EBF9"><%if session("aleave")="check" then%>ɾ��<%else%><a href="javascript:DoEmpty('info_replydel.asp?selectedid=<%=rs1("id")%>&xid=<%=uid%>&back=xinxi');">ɾ��</a><%end if%></td>
      </tr>
     
	  <%rs1.movenext
		loop
		rs1.close
		set rs1=nothing
		end if%>
    </table>	</td>
  </tr>

  <tr bgcolor="#ffffff" id="left2_<%=rs("id")%>" style="display:none;">
    <td height="22" colspan="14" align="right">
	<table width="95%" border="0" cellpadding="2" cellspacing="1" bgcolor="#999999">
      <tr>
        <td width="5%" height="25" align="center" bgcolor="#cccccc">id</td>
        <td width="10%" align="center" bgcolor="#cccccc">�ٱ�������</td>
        <td width="12%" align="center" bgcolor="#cccccc">�ٱ�����</td>
        <td width="40%" align="center" bgcolor="#cccccc">�ٱ�����</td>
		<td width="13%" align="center" bgcolor="#cccccc">�ٱ�ʱ��</td>
		<td width="10%" align="center" bgcolor="#cccccc">IP��ַ</td>        
        <td width="5%" align="center" bgcolor="#cccccc">���</td>
        <td width="5%" align="center" bgcolor="#cccccc">����</td>
      </tr>
	  <%

		if rs2.eof and rs2.bof then
		
		%>
		 <tr>
        <td height="25" colspan="14" align="center" bgcolor="#E3EBF9">��ʱ��û�оٱ�</td>
        </tr>
		<%else
		do while not rs2.eof 
	  %>
      <tr>
        <td height="25" align="center" bgcolor="#E3EBF9"><%=rs2("id")%></td>
        <td align="center" bgcolor="#E3EBF9"><a href="#" onClick="window.open('Bad_info_mail.asp?email=<%=rs2("Bad_infomail")%>&biaoti=<%=rs("biaoti")%>','Sample','toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=no,resizable=no,copyhistory=no,width=455,height=328,left=300,top=100')"><%=rs2("Bad_infomail")%></a></td>
        <td bgcolor="#E3EBF9"><%=rs2("Bad_infolx")%></td>
		<td bgcolor="#E3EBF9"><%=HTMLEncode(rs2("Bad_info"))%></td>
        <td align="center" bgcolor="#E3EBF9"><%=rs2("addsj")%></td>
		<td align="center" bgcolor="#E3EBF9"><%=rs2("ip")%></td>			
       <td align="center" bgcolor="#E3EBF9">
	<%if rs2("sh")=1 then%><a href="Bad_info_yz.asp?dick=delyz&selectedid=<%=rs2("id")%>">��</a><%else%><b><a href="Bad_info_yz.asp?selectedid=<%=rs2("id")%>"><font color="#FF0000">��</font></a></b><%end if%></td>
        <td align="center" bgcolor="#E3EBF9"><a href="javascript:DoEmpty('Bad_info_del.asp?selectedid=<%=rs2("id")%>');">ɾ��</a></td>
      </tr>
     
	  <%rs2.movenext
		loop
		rs2.close
		set rs2=nothing
		end if%>
    </table>	</td>
  </tr>
<%
k=k+1
rs.movenext
if k>=Pagesize then exit do
loop
Call CloseRs(rs)
Call CloseDB(conn)
%>
  <tr>
    <td width="100%" height="1" colspan="14" style="border-top-style: solid; border-top-width: 1">
<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%" height="1">
<tr> 
<td width="100%" height="25" align="center" bgcolor="#eeeeee"><div align="center">
      <input onClick=CheckAll(this.form) type=checkbox value=on name=chkall>
      ѡ������
<input onclick=javascript:showoperatealert(1) type="submit" value="��֤ͨ��" name="B1">
<input onclick=javascript:showoperatealert(2) type="submit" value="ȡ����֤" name="B2">	  
<input onclick=javascript:showoperatealert(3) type="submit" value="��Ϣ�Ƽ�" name="B3">
<input onclick=javascript:showoperatealert(4) type="submit" value="ͼƬ�Ƽ�" name="B4">
<input onclick=javascript:showoperatealert(5) type="submit" value="ȡ���Ƽ�" name="B5">
<input name='submit1' type='submit' value=' ����ɾ�� ' onClick="return test();"  style="color: #FF0000">

<br>����&nbsp;<font color="#CC5200"><%=Allrecord%></font>&nbsp;����¼ �� <font color="#CC5200"><%=Allpage%></font> ҳ ÿҳ <font color="#CC5200"><%=Pagesize%></font> �� �����ǵ� <font color="#CC5200"><%=ThisPage%></font> ҳ
    <%
if ThisPage<2 then     
response.write "<font color=""#808080"">��ҳ</font>&nbsp;"
response.write "<font color=""#808080"">��һҳ</font>&nbsp;"     
else     
response.write "<a href=?page=1&key1="&trim(request("key1"))&"&key2="&trim(request("key2"))&"&key3="&trim(request("key3"))&"&skey="&trim(request("skey"))&"&t1="&trim(request("t1"))&"&t2="&trim(request("t2"))&"&t3="&trim(request("t3"))&"&dick="&dick&"&city_one="&city_oneid&"&city_two="&city_twoid&"&city_three="&city_threeid&"&type_one="&type_oneid&"&type_two="&type_twoid&"&type_three="&type_threeid&">��ҳ</a>&nbsp;"
response.write "<a href=?page="&ThisPage-1&"&key1="&trim(request("key1"))&"&key2="&trim(request("key2"))&"&key3="&trim(request("key3"))&"&skey="&trim(request("skey"))&"&t1="&trim(request("t1"))&"&t2="&trim(request("t2"))&"&t3="&trim(request("t3"))&"&dick="&dick&"&city_one="&city_oneid&"&city_two="&city_twoid&"&city_three="&city_threeid&"&type_one="&type_oneid&"&type_two="&type_twoid&"&type_three="&type_threeid&">��һҳ</a>&nbsp;"     
end if
if Allpage-ThisPage<1 then     
response.write "<font color=""#808080"">��һҳ</font>&nbsp;"
response.write "<font color=""#808080"">βҳ</font>&nbsp;"  
else     
response.write "<a href=?page="&(ThisPage+1)&"&key1="&trim(request("key1"))&"&key2="&trim(request("key2"))&"&key3="&trim(request("key3"))&"&skey="&trim(request("skey"))&"&t1="&trim(request("t1"))&"&t2="&trim(request("t2"))&"&t3="&trim(request("t3"))&"&dick="&dick&"&city_one="&city_oneid&"&city_two="&city_twoid&"&city_three="&city_threeid&"&type_one="&type_oneid&"&type_two="&type_twoid&"&type_three="&type_threeid&">��һҳ</a>&nbsp;"   
response.write "<a href=?page="&Allpage&"&key1="&trim(request("key1"))&"&key2="&trim(request("key2"))&"&key3="&trim(request("key3"))&"&skey="&trim(request("skey"))&"&t1="&trim(request("t1"))&"&t2="&trim(request("t2"))&"&t3="&trim(request("t3"))&"&dick="&dick&"&city_one="&city_oneid&"&city_two="&city_twoid&"&city_three="&city_threeid&"&type_one="&type_oneid&"&type_two="&type_twoid&"&type_three="&type_threeid&">βҳ</a>&nbsp;"     
end if
%>
</div></td>
</tr>
</table>
    </td>
  </tr>
  </form>
  <tr>
    <td width="100%" height="1" colspan="10" style="border-top-style: solid; border-top-width: 1"></td>
  </tr>
</table></center>
</div>
