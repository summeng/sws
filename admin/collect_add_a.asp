<!--#include file="../conn.asp"-->
<!--#include file=cookies.asp-->
<!--#include file=../config.asp-->
<!--#include file="../Include/head.asp"-->
<!--#include file="../Include/Function.asp"-->
<!--#include file="../Include/label.asp"-->
<!--#include file="inc/Function.asp"-->
<%
'=====================================
'���칩����Ϣ��վ����ϵͳ
'��������Ƽ������ҳ�Ʒ
'�ͷ�����http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>
<%response.buffer=true	

Call Main()
'�ر����ݿ�����
Conn.Close:Set Conn=Nothing
Sub Main
Call OpenConn%>
<html>
<head>
<title>�ɼ�ϵͳ</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="../css/manage.css" type="text/css">
</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" class="border">
  <tr> 
    <td height="22" colspan="2" align="center" class="topbg"><strong>��&nbsp;&nbsp;��&nbsp;&nbsp;ϵ&nbsp;&nbsp;ͳ&nbsp;&nbsp;��&nbsp;&nbsp;Ŀ&nbsp;&nbsp;��&nbsp;&nbsp;��</td>
  </tr>
</table>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" class="border">
  <tr class="tdbg"> 
    <td width="65" height="30"><strong>��������</strong></td>
    <td height="30"><a href="collect_add_a.asp">�����Ŀ</a> >> <font color=red>��������</font> >> �б����� >> �������� >> �������� >> �������� >> �������� >> ���</td>         
  </tr>         
</table>
<br>         
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" class="border" >
<form method="post" action="collect_add_b.asp" name="myform">
    <tr> 
      <td height="22" colspan="2" class="title"> <div align="center"><strong>�� �� �� 
		�� Ŀ--�� �� �� ��</strong></div></td>
    </tr>
    <tr class="tdbg"> 
      <td width="20%" class="tdbg"><strong>��Ŀ���ƣ�</strong></td>
      <td class="tdbg" width="75%">
		<input name="ItemName" type="text" size="27" maxlength="30">&nbsp;&nbsp;<font color=red>*</font>�磺���������������� �ȼ�������
      </td>
    </tr>
    <tr class="tdbg"> 
      <td width="20%" class="tdbg"><strong> ����������</strong></td>
      <td class="tdbg" width="75%"><%
set rs=conn.execute("select * from DNJY_city where id>0 and twoid>0 and threeid=0")
%>
        <SCRIPT language = "JavaScript">
var onecount;
onecount=0;
subcat = new Array();
        <%
		dim count:count = 0
        do while not rs.eof 
        %>
subcat[<%=count%>] = new Array("<%=rs("city")%>","<%=rs("id")%>","<%=rs("twoid")%>");
        <%
        count = count + 1
        rs.movenext
        loop
Call CloseRs(rs)
        %>
onecount=<%=count%>;
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
    document.all.city_two.length = 0; 
	document.all.city_two.options[0] = new Option('ѡ�����','');
	document.all.city_three.length = 0; 
	document.all.city_three.options[0] = new Option('ѡ�����','');
    var locationid=locationid;
    var i;
    for (i=0;i < onecount; i++)
        {
            if (subcat[i][1] == locationid)
            { 
                document.all.city_two.options[document.all.city_two.length] = new Option(subcat[i][0], subcat[i][2]);
            }        
        }
        
    }    
	
	function changelocation4(locationid,locationid1)
    {
    document.all.city_three.length = 0; 
    document.all.city_three.options[0] = new Option('ѡ�����','');
    var locationid=locationid;
	 var locationid1=locationid1;
    var i;
    for (i=0;i < onecount4; i++)
        {
            if (subcat4[i][2] == locationid)
            { 
			if (subcat4[i][1] == locationid1)
			{
                document.all.city_three.options[document.all.city_three.length] = new Option(subcat4[i][0], subcat4[i][3]);
            }        
        }
       } 
    }            
              </SCRIPT>
        <SELECT name="city_one" size="1" id="select2" onChange="changelocation(document.all.city_one.options[document.all.city_one.selectedIndex].value)">
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
rs.close
set rs = nothing
%>
        </SELECT>
        <SELECT name="city_two" id="city_two" onChange="changelocation4(document.all.city_two.options[document.all.city_two.selectedIndex].value,document.all.city_one.options[document.all.city_one.selectedIndex].value)">
          <OPTION value="" selected>ѡ�����</OPTION>
        </SELECT>
        <SELECT name="city_three" id="city_three">
          <OPTION value="" selected>ѡ�����</OPTION>
        </SELECT>
        <FONT color="#FF0000">*</FONT> </td>
    </tr>
    <tr class="tdbg"> 
      <td width="20%" class="tdbg"><strong> ������Ŀ��</strong></td>
      <td class="tdbg" width="75%"><%
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
        <%
		dim count3:count3 = 0
        do while not rs.eof 
        %>
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
Call CloseRs(rs)
%>
          </SELECT>
          <SELECT name="type_two" id="select9" onChange="changelocation3(document.myform.type_two.options[document.myform.type_two.selectedIndex].value,document.myform.type_one.options[document.myform.type_one.selectedIndex].value)">
            <OPTION value="" selected>��������</OPTION>
          </SELECT>
          <SELECT name="type_three" id="select10">
            <OPTION value="" selected>��������</OPTION>
          </SELECT></td>
    </tr>
    <tr>
      <td width="20%" class="tdbg"><strong> ��վ���ƣ�</strong></td>
      <td class="tdbg" width="75%">
		<input name="WebName" type="text" size="27" maxlength="30">
      </td>
    </tr>
    <tr>
      <td width="20%" class="tdbg"><strong> ��վ��ַ��</strong></td>
      <td class="tdbg" width="75%"><input name="WebUrl" type="text" size="49" maxlength="100">
      </td>
    </tr>
   <tr class="tdbg"> 
      <td width="150" class="tdbg"><strong> ��վ��¼��</strong></td>
      <td class="tdbg">
		<input type="radio" value="0" name="LoginType" checked onClick="Login.style.display='none'">����Ҫ��¼<span lang="en-us">&nbsp;
		</span>
		<input type="radio" value="1" name="LoginType" onClick="Login.style.display=''">���ò���</td>
    </tr>
   <tr class="tdbg" id="Login" style="display:none"> 
      <td width="150" class="tdbg"><strong> ��¼������</strong></td>
      <td class="tdbg">
        ��¼��ַ��<input name="LoginUrl" type="text" size="40" maxlength="100" value=""><br>
        �ύ��ַ��<input name="LoginPostUrl" type="text" size="40" maxlength="100" value=""><br>
        �û�������<input name="LoginUser" type="text" size="30" maxlength="30" value=""><br>
        ���������<input name="LoginPass" type="text" size="30" maxlength="30" value=""><br> 
		ʧ����Ϣ��<input name="LoginFalse" type="text" size="30" maxlength="50" value=""></td>
    </tr>
    <tr class="tdbg"> 
      <td width="20%" class="tdbg"><strong>��Ŀ��ע��</strong></td>
      <td class="tdbg" width="75%"><textarea name="ItemDemo" cols="49" rows="5"></textarea></td>
    </tr>
    <tr class="tdbg"> 
      <td colspan="2" align="center" class="tdbg"><input name="Action" type="hidden" id="Action" value="SaveAdd">
        <input name="Cancel" type="button" id="Cancel" value=" ��&nbsp;&nbsp;�� " onClick="window.location.href='collect_management.asp'" style="cursor: hand;background-color: #cccccc;">
        &nbsp; 
        <input  type="submit" name="Submit" value="��&nbsp;һ&nbsp;��" style="cursor: hand;background-color: #cccccc;"></td>
    </tr>
</form>
</table>    
</body>         
</html>
<%end Sub%>
