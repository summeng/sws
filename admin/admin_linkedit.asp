<!--#include file="../conn.asp"-->
<!--#include file=cookies.asp-->
<!--#include file="../config.asp"-->
<!--#include file="../Include/head.asp"-->
<!--#include file="../Include/Function.asp"-->
<!--#include file="../Include/label.asp"-->
<!--#include file="../Include/md5.asp"-->
<%
'=====================================
'���칩����Ϣ��վ����ϵͳ
'��������Ƽ������ҳ�Ʒ
'�ͷ�����http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>

<%call checkmanage("12")
dim action,url,img,info,name,wzname,dd,adminlink,ttt
id=strint(request.form("id"))
If request.form("id")="" Then id=strint(request("id"))
action=HtmlEncode(request.querystring("action"))
name=HtmlEncode(request("name"))
wzname=HtmlEncode(request("wzname"))
url=request("url")
ttt=request("ttt")
adminlink=trim(request("adminlink"))
city_oneid=strint(request.form("city_one"))
city_twoid=strint(request.form("city_two"))
city_threeid=strint(request.form("city_three"))
if city_oneid>0 then
set rs=server.createobject("adodb.recordset")
rs.open "select city from DNJY_city where twoid=0 and threeid=0 and id="&city_oneid,conn,1,1
city_one=rs("city")
rs.close
if city_twoid>0 then
rs.open "select city from DNJY_city where id="&city_oneid&" and threeid=0 and twoid="&city_twoid,conn,1,1
city_two=rs("city")
rs.close
end if
if city_twoid>0 and city_threeid>0 then
rs.open "select city from DNJY_city where id="&city_oneid&" and twoid="&city_twoid&" and threeid="&city_threeid,conn,1,1
city_three=rs("city")
rs.close
end if
set rs=nothing
end If
select case action
case "add" 
set rs=server.CreateObject("adodb.recordset")
	rs.open "select wzname from DNJY_link where name='"&name&"'",conn,1,1
	if not rs.eof Then
	Response.Write ("<script language=javascript>alert('�û���"&name&"�Ѿ����ڣ�������ѡ��һ���û���!');history.go(-1);</script>")
	set rs=Nothing
response.end	
	end If	
set rs=server.CreateObject("adodb.recordset")
rs.open "select wzname,url,logo,info,img,name,pass,paixu,addtime,known,city_oneid,city_twoid,city_threeid,city_one,city_two,city_three,mail,linktop,c from DNJY_link",conn,1,3
rs.AddNew
rs(0)=wzname
rs(1)=url
rs("logo")=trim(request("logo"))
rs("info")=trim(request("info"))
rs("img")=trim(request("img"))
rs("name")=trim(request("name"))
rs("pass")=md5(trim(request("pass")))
rs("paixu")=strint(trim(request("paixu")))
rs("mail")=trim(request("mail"))
rs("linktop")=trim(request("linktop"))
rs("c")=trim(request("cc"))
rs("addtime")=now()
rs("known")=1
if city_oneid="" then city_oneid=0
if city_twoid="" then city_twoid=0
if city_threeid="" then city_threeid=0
rs("city_oneid")=city_oneid
rs("city_one")=city_one
rs("city_two")=city_two
rs("city_twoid")=city_twoid
rs("city_three")=city_three
rs("city_threeid")=city_threeid
rs.Update
response.Write "<script language=javascript>alert('��ӳɹ�!');location.replace('admin_link.asp');</script>"
rs.Close:set rs=nothing

case "edit"
set rs=server.CreateObject("adodb.recordset")
rs.open "select wzname,url,logo,info,img,pass,paixu,city_oneid,city_twoid,city_threeid,city_one,city_two,city_three,pass,mail,linktop,c from DNJY_link where id="&id,conn,1,3
rs(0)=wzname
rs(1)=url
rs("logo")=trim(request("logo"))
rs("info")=trim(request("info"))
rs("img")=trim(request("img"))
if trim(request("pass"))<>"" then
rs("pass")=md5(trim(request("pass")))
end If
rs("mail")=trim(request("mail"))
rs("linktop")=trim(request("linktop"))
rs("c")=trim(request("cc"))
rs("paixu")=trim(request("paixu"))
if city_oneid="" then city_oneid=0
if city_twoid="" then city_twoid=0
if city_threeid="" then city_threeid=0
rs("city_oneid")=city_oneid
rs("city_one")=city_one
rs("city_two")=city_two
rs("city_twoid")=city_twoid
rs("city_three")=city_three
rs("city_threeid")=city_threeid
rs.Update
rs.Close:set rs=Nothing
response.Write "<script language=javascript>alert('�޸ĳɹ�!');location.replace('?adminlink=1&id="&id&"');</script>"

case "del"
Dim Num,objFSO
if action="del" then
If Request.Form("DelID")="" or isnull(Request.Form("DelID")) or isempty(Request.Form("DelID")) then
response.write "<font style='color:red;font-size:12px'>��ѡ��Ҫɾ������Ϣ</font>"
else
Num=request.form("DelID").count
for i=1 to Num
set rs = server.CreateObject ("Adodb.recordset")
rs.open "select * from DNJY_link where id="&request.form("DelID")(i),conn,3,3
do while not rs.eof

logo=trim(rs("logo"))
if left(logo,6)="images" then
Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
if objFSO.fileExists(Server.MapPath("../"&logo)) then
objFSO.DeleteFile(Server.MapPath("../"&logo))
end if
set objfso=nothing
end If

rs.delete
rs.movenext
Loop
Next
Call CloseRs(rs)
end If
response.write "<script language='javascript'>" & chr(13)
		response.write "alert('�ɹ�ɾ����');" & chr(13)
		response.write "window.document.location.href='admin_link.asp';"&chr(13)
		response.write "</script>" & chr(13)
response.end
end If

end select
%>
<script LANGUAGE=JavaScript>
  function textLimitCheck(thisArea, maxLength){
    if (thisArea.value.length > maxLength){
      alert(maxLength + ' ��������. \r�����Ľ��Զ�ȥ��.');
      thisArea.value = thisArea.value.substring(0, maxLength);
      thisArea.focus();
    }
    /*��дspan��ֵ����ǰ��д���ֵ�����*/
    messageCount.innerText = thisArea.value.length;
  }
</script>
<script language = "JavaScript">
<!--
//�ֽ������ƣ����������=========
String.prototype.Length=function(){
temp=this.replace(/([^\x00-\xff])/g,"$1$1");
return temp.length
}

//��֤��ַ��ȷ��
function checkeURL(URL){ 
var str=URL;
//��JavaScript�У�������ʽֻ��ʹ��"/"��ͷ�ͽ���������ʹ��˫����
//�ж�URL��ַ��������ʽΪ:http(s)?://([\w-]+\.)+[\w-]+(/[\w- ./?%&=]*)?
//����Ĵ�����Ӧ����ת���ַ�"\"���һ���ַ�"/"
var Expression=/http(s)?:\/\/([\w-]+\.)+[\w-]+(\/[\w- .\/?%&=]*)?/; 
var objExp=new RegExp(Expression);
if(objExp.test(str)==true){
return true;
}else{
return false;
}
}

//��֤������ȷ��
function checkemail(mail){
var str=mail;
//��JavaScript�У�������ʽֻ��ʹ��"/"��ͷ�ͽ���������ʹ��˫����
var Expression=/\w+([-+.']\w+)*@\w+([-.]\w+)*\.\w+([-.]\w+)*/; 
var objExp=new RegExp(Expression);
if(objExp.test(str)==true){
return true;
}else{
return false;
}
} 

function CheckForm(){   

	if (document.form.wzname.value.length == 0) {
		alert("������վ���ƣ�");
		document.form.wzname.focus();
		return false;
	}
     if ((document.form.wzname.value.Length()>16) || (document.form.wzname.value.Length()<1)) //�ֽ������ƣ����������
     {
	 alert("��վ���Ƴ���Ҫ��1��16�ֽڣ����������룡");
	  document.form.wzname.focus()
	  return false
  }
	if (document.form.url.value.length == 0) {
		alert("������ַ��");
		document.form.url.focus();
		return false;
	}
    var url=form.url.value;
     if(!checkeURL(url)){
     alert("���������ַַ�����Ϲ淶�����������룡");
	  document.form.url.focus()
	  return false
     }

	if (document.form.name.value.length == 0) {
		alert("�����û�����");
		document.form.name.focus();
		return false;
	}
     if ((document.form.name.value.Length()>8) || (document.form.name.value.Length()<1)) //�ֽ������ƣ����������
     {
	 alert("�û�������Ҫ��1��8�ֽڣ����������룡");
	  document.form.name.focus()
	  return false
  }
  <%if adminlink="2" then%>
	if (document.form.pass.value.length == 0) {
		alert("�������룡");
		document.form.pass.focus();
		return false;
	}
     if ((document.form.pass.value.Length()>12) || (document.form.pass.value.Length()<5)) //�ֽ������ƣ����������
     {
	 alert("���볤��Ҫ��5��12�ֽڣ����������룡");
	  document.form.pass.focus()
	  return false
  }
	if (document.form.passs.value.length == 0) {
		alert("����ȷ�����룡");
		document.form.passs.focus();
		return false;
	}
    if(document.form.pass.value != document.form.passs.value) {
	document.form.passs.focus();
	document.form.passs.value = '';
    alert("������������벻ͬ�����������룡");
	return false;
  }
<%else%>
     if (document.form.pass.value.length != 0 && ((document.form.pass.value.Length()>12) || (document.form.pass.value.Length()<5))) //�ֽ������ƣ����������
     {
	 alert("���볤��Ҫ��5��12�ֽڣ����������룡");
	  document.form.pass.focus()
	  return false
  }
	if (document.form.pass.value.length != 0 && document.form.passs.value.length == 0) {
		alert("����ȷ�����룡");
		document.form.passs.focus();
		return false;
	}
    if(document.form.pass.value != document.form.passs.value) {
	document.form.passs.focus();
	document.form.passs.value = '';
    alert("������������벻ͬ�����������룡");
	return false;
  }
  <%end if%>
	return true;
}
// -->
</script>
<HTML>
<HEAD>
<TITLE>�������ӹ���</TITLE>
<META http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK rel="stylesheet" href="../css/style.css" type="text/css">
</HEAD>
<BODY>
<TABLE width="98%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#799AE1">
  <TR> 
    <TD height="20" align="center" bgcolor="#799AE1"><FONT color="#FFFFFF" style="font-size:14px">&nbsp;
    </FONT><FONT color="#FFFFFF" style="font-size:14px">�������ӹ���</FONT></TD>
  </TR>

 <%If adminlink="1" Then
 set rs=server.CreateObject("adodb.recordset")
  rs.Open "select wzname,url,id,addtime,logo,info,mail,linktop,c,paixu,img,name,pass,city_oneid,city_twoid,city_threeid from DNJY_link where id="&id,conn,1,1%>
		  <FORM name="form" method="post" action="?action=edit&id=<%=int(rs("id"))%>&city_oneid=<%=city_oneid%>&city_twoid=<%=city_twoid%>&city_threeid=<%=city_threeid%>" language="javascript" onsubmit="return CheckForm()">
		   <TR> 
    <TD height="20" bgcolor="#799AE1" align="center"><FONT color="#FFFFFF" style="font-size:14px">�� �� �� �� �� ��</FONT></TD>
  </TR>
  <TR> 
    <TD bgcolor="#FFFFFF"> <BR>
	<TABLE width="98%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="#D6DFF7">
            <tr> 
      <td width="188" height="25"align="right">����������</td>
      <td colspan="2">
<%Dim rsi
set rsi=conn.execute("select * from DNJY_city where id>0 and twoid>0 and threeid=0")
%>
      <SCRIPT language = "JavaScript">
var onecount;
onecount=0;
subcat = new Array();
        <%
		dim count:count = 0
        do while not rsi.eof 
        %>
subcat[<%=count%>] = new Array("<%=rsi("city")%>","<%=rsi("id")%>","<%=rsi("twoid")%>");
        <%
        count = count + 1
        rsi.movenext
        loop
        rsi.close
		set rsi=nothing
        %>
onecount=<%=count%>;
</SCRIPT>
<%
set rsi=conn.execute("select * from DNJY_city where id>0 and twoid>0 and threeid>0")
%>
 <SCRIPT language = "JavaScript">
var onecount4;
onecount4=0;
subcat4 = new Array();
        <%
		dim count4:count4 = 0
        do while not rsi.eof 
        %>
subcat4[<%=count4%>] = new Array("<%=rsi("city")%>","<%=rsi("id")%>","<%=rsi("twoid")%>","<%=rsi("threeid")%>");
        <%
        count4 = count4 + 1
        rsi.movenext
        loop
        rsi.close
		set rsi = nothing
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
      <SELECT name="city_one" size="1" id="select2" onChange="changelocation(document.form.city_one.options[document.form.city_one.selectedIndex].value)">
            <OPTION value="" selected>ѡ�����</OPTION>
  <%set rsi=conn.execute("select * from DNJY_city where id>0 and twoid=0 and threeid=0 order by indexid asc")
if rsi.eof or rsi.bof then
response.write "<option value=''>û�з���</option>"
else
do until rsi.eof%>
  <OPTION value="<%=rsi("id")%>" <%if rsi("id")=rs("city_oneid") then%>selected<%end if%>><%=rsi("city")%></OPTION>
 <%rsi.movenext
    loop
	%>
	<%end if
rsi.close
set rsi = nothing
%>
      </SELECT> 
	  <SELECT name="city_two" id="city_two" onChange="changelocation4(document.form.city_two.options[document.form.city_two.selectedIndex].value,document.form.city_one.options[document.form.city_one.selectedIndex].value)">
    <OPTION value="" selected>ѡ�����</OPTION>
   <%
set rsi=conn.execute("select * from DNJY_city where id="&rs("city_oneid")&" and twoid>0 and threeid=0")
if rsi.eof and rsi.bof then
response.write "<option value=''>û�з���</option>"
else
do until rsi.eof%>
  <OPTION value="<%=rsi("twoid")%>" <%if rsi("twoid")=rs("city_twoid") then%>selected<%end if%>><%=rsi("city")%></OPTION>
 <%rsi.movenext
    loop
	%>
	<%end if
rsi.close
set rsi = nothing
%>
	</SELECT>
	     <SELECT name="city_three" id="city_three">
         <OPTION value="" selected>ѡ�����</OPTION>
		<%
set rsi=conn.execute("select * from DNJY_city where id="&rs("city_oneid")&" and twoid="&rs("city_twoid")&" and threeid<>0")
if rsi.eof and rsi.bof then
response.write "<option value=''>û�з���</option>"
else
do until rsi.eof%>
<OPTION value="<%=rsi("threeid")%>" <%if rsi("threeid")=rs("city_threeid") then%>selected<%end if%>><%=rsi("city")%></OPTION>
   <% rsi.movenext
    loop
	%>
<%end if
rsi.close
set rsi = nothing
%>
    </SELECT></td>
    </tr>
  <input type="hidden" name="ttt" id="ttt" value="1">
  <input type="hidden" name="id" id="id" value="<%=rs(2)%>">
    <tr> 
      <td width="188" height="25"align="right">��վ���ƣ�</td>
      <td colspan="2">
<INPUT name="wzname" type="text" id="wzname" value="<%=trim(rs(0))%>" size="20"> <font color="#FF0000">*</font></td>
    </tr>
    <tr> 
      <td width="188" height="25"align="right">��&nbsp;&nbsp;&nbsp;&nbsp;ַ��</td>
      <td colspan="2">
<INPUT name="url" type="text" id="url" value="<%=trim(rs(1))%>" size="35"> <font color="#FF0000">*</font></td>
    </tr>
	    <tr> 
      <td width="188" height="25"align="right">�������䣺</td>
      <td colspan="2">
<INPUT name="mail" type="text" id="mail" value="<%=trim(rs("mail"))%>" size="20"> <font color="#FF0000">����д��ȷ������</font></td>
    </tr>
    <tr> 
      <td width="188" height="25"align="right">�����û�����</td>
      <td colspan="2"><input type="hidden" name="name" value="<%=trim(rs("name"))%>">
<%=trim(rs("name"))%></td>
    </tr>
    <tr> 
      <td width="188" height="25"align="right">�������룺</td>
      <td colspan="2">
<INPUT name="pass" type="password" id="pass" value="" size="20"> <font color="#FF0000">������޸�����������</font></td>
    </tr>
    <tr> 
      <td width="188" height="25"align="right">����ȷ�ϣ�</td>
      <td colspan="2">
<INPUT name="passs" type="password" id="passs" value="" size="20"></td>
    </tr>

    <tr> 
      <td width="188" height="25"align="right">��ʾ��ʽ��</td>
      <td colspan="2">�� 
        <input type="radio" value="0" <%if rs("img")=0 then Response.Write "checked"%>  name="img">���� 
        <input type="radio" value="1" <%if rs("img")=1 then Response.Write "checked"%> name="img">
        ͼ�� ��<font color="#FF0000">ѡ��ͼ��ʱ��ע����дͼ���ַ���ϴ�ͼ�� ��</font></td>
    </tr>
    <tr> 
      <td width="188" height="25"align="right">�Ƿ���ɫ��</td>
      <td colspan="2">�� 
        <input type="radio" value="1" <%if rs("c")=1 then Response.Write "checked"%>  name="cc">�� 
        &nbsp;&nbsp;<input type="radio" value="0" <%if rs("c")=0 then Response.Write "checked"%> name="cc">
        �� �� ��<font color="#FF0000">��ɫ��վ������ʾ��ɫ��</font></td>
    </tr> 
    <tr> 
      <td width="188" height="25"align="right">�Ƿ��ö���</td>
      <td colspan="2">�� 
        <input type="radio" value="1" <%if rs("linktop")=1 then Response.Write "checked"%>  name="linktop">�� 
        &nbsp;&nbsp;<input type="radio" value="0" <%if rs("linktop")=0 then Response.Write "checked"%> name="linktop">
        �� �� ��<font color="#FF0000">�ö���վ������ǰ�棡</font></td>
    </tr> 
      <tr><td width="188" height="25"align="right">��&nbsp;&nbsp;&nbsp;&nbsp;��</td>
      <td colspan="2">
<input style="BACKGROUND-COLOR: #eeeeee; BORDER-BOTTOM: #FFFFFF 1px solid; BORDER-LEFT: #a2a2a2 1px solid; BORDER-RIGHT: #ffffff 1px solid; BORDER-TOP: #a2a2a2 1px solid; COLOR: #333333; FONT-SIZE: 9pt; HEIGHT: 19px;" type="text" name="paixu" size="5" value="<%=rs("paixu")%>" onkeyup="this.value=this.value.replace(/\D/g,'')" onafterpaste="this.value=this.value.replace(/\D/g,'')"> </td>
    </tr>
   <tr> 
      <td width="188" height="25"align="right">��վ˵����</td>
      <td colspan="2"><textarea name="info" cols="50" rows="3" id="info" onkeyUp="textLimitCheck(this, 50);"><%=trim(rs("info"))%></textarea>&nbsp;&nbsp; �� 50 ���ַ� ������ <font color="#CC0000"><span id="messageCount">0</span></font> ����></td>
    </tr>
    <tr> 
      <td width="188" height="25"align="right">��վLOGO��</td>
      <td colspan="2"><input name="logo" type="text" class="input" id="logo" size="50" maxlength="255" value="<%=trim(rs("logo"))%>">û�������� <%If rs("logo")<>"" then%><img src="<%If Left(rs("logo"),7)="http://" then%><%=rs("logo")%><%else%>../<%=rs("logo")%><%End if%>" border=1 width="88"  height="31"><%End if%><br><iframe name="ad" frameborder=0 width="450" height="35" scrolling=no src="../DNJY_upload.asp?ttup=link"></iframe><br>&nbsp;&nbsp;<font color="#ff0000">�����ϴ����ļ����ͣ�gif|jpgs|jpg|bmp|png&nbsp;&nbsp;&nbsp;�ޣ�88*31���أ�500k����</font> </td>
    </tr>


  <tr bgcolor="#eeeeee">
    <td height="35" colspan="10"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td><div align="center">
		
            <input type="submit" value="ȷ��" name="B1" onClick="if(wzname.value=='' || url.value==''){alert('���ƺ͵绰������Ϊ��!');return false;}">
        </div></td>
        <td><div align="center">
            <input type="reset" value="ȡ��" name="B1">
        </div></td>
      </tr>          
        </FORM>
        <%Call CloseRs(rs) %>
      </TABLE>
	<BR></TD>
  </TR>
</TABLE>
<%End if%>
<%If adminlink="2" then%>
<FORM name="form" method="post" action="?action=add&c=<%=city_oneid%>,<%=city_twoid%>,<%=city_threeid%>" language="javascript" onsubmit="return CheckForm()">
<TABLE width="98%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#799AE1">
  <TR> 
    <TD height="20" bgcolor="#799AE1" align="center"><FONT color="#FFFFFF" style="font-size:14px">�� �� �� �� �� ��</FONT></TD>
  </TR>
  <TR> 
    <TD bgcolor="#FFFFFF">
	<BR>
	<TABLE width="98%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="#D6DFF7">
   <TR> 
    <TD bgcolor="#FFFFFF"> <BR>
	<TABLE width="98%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="#D6DFF7">
            <tr> 
      <td width="188" height="25" align="right">����������</td>
      <td colspan="2">
		<%
set rs=conn.execute("select * from DNJY_city where id>0 and twoid>0 and threeid=0")
%>
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
		//dim count4:
		count4 = 0
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
          </SELECT></td>
    </tr>  
    <tr> 
      <td width="188" height="25" align="right">��վ���ƣ�</td>
      <td colspan="2">
<INPUT name="wzname" type="text" id="wzname" value="" size="20"> <font color="#FF0000">*</font></td>
    </tr>
    <tr> 
      <td width="188" height="25" align="right">��&nbsp;&nbsp;&nbsp;&nbsp;ַ��</td>
      <td colspan="2">
    <INPUT name="url" type="text" id="url" value="" size="35"> <font color="#FF0000">*</font></td>
    </tr>
	    <tr> 
      <td width="188" height="25" align="right">�������䣺</td>
      <td colspan="2">
<INPUT name="mail" type="text" id="mail" value="" size="20"> <font color="#FF0000">����д��ȷ������</font></td>
    </tr>
    <tr> 
      <td width="188" height="25" align="right">�����û�����</td>
      <td colspan="2">
<INPUT name="name" type="text" id="name" value="" size="20"> <font color="#FF0000">*</font></td>
    </tr>
    <tr> 
      <td width="188" height="25" align="right">�������룺</td>
      <td colspan="2">
<INPUT name="pass" type="text" id="pass" value="" size="20"> <font color="#FF0000">*</font></td>
    </tr>
    <tr> 
      <td width="188" height="25" align="right">����ȷ�ϣ�</td>
      <td colspan="2">
<INPUT name="passs" type="text" id="passs" value="" size="20"> <font color="#FF0000">*</font></td>
    </tr>
    <tr> 
      <td height="25" align="right">��ʾ��ʽ��</td>
      <td colspan="2"> 
        <input type="radio" name="img" value="0" checked>
        ���� 
        <input name="img" type="radio" value="1" >
        ͼ�� ��<font color="#FF0000">ѡ��ͼ��ʱ��ע����дͼ���ַ���ϴ�ͼ�꣡</font></td>
    </tr>
   <tr> 
      <td height="25" align="right">�Ƿ���ɫ��</td>
      <td colspan="2"> 
        <input type="radio" name="cc" value="1">
        �� 
        &nbsp;&nbsp;<input name="cc" type="radio" value="0" checked >
        �� ��<font color="#FF0000">��ɫ��վ������ʾ��ɫ��</font></td>
    </tr> 	
    <tr> 
      <td height="25" align="right">�Ƿ��ö���</td>
      <td colspan="2"> 
        <input type="radio" name="linktop" value="1">
        �� 
        &nbsp;&nbsp;<input name="linktop" type="radio" value="0" checked >
        �� ��<font color="#FF0000">�ö���վ������ǰ�棡</font></td>
    </tr> 
    <tr> 
      <td width="188" height="25" align="right">��&nbsp;&nbsp;&nbsp;&nbsp;��</td>
      <td colspan="2">
<input style="BACKGROUND-COLOR: #eeeeee; BORDER-BOTTOM: #FFFFFF 1px solid; BORDER-LEFT: #a2a2a2 1px solid; BORDER-RIGHT: #ffffff 1px solid; BORDER-TOP: #a2a2a2 1px solid; COLOR: #333333; FONT-SIZE: 9pt; HEIGHT: 19px;" type="text" name="paixu" size="5" value="" onkeyup="this.value=this.value.replace(/\D/g,'')" onafterpaste="this.value=this.value.replace(/\D/g,'')"> ����Խ��Խ��ǰ</td>
    </tr>
    <tr> 
      <td width="188" height="25" align="right">��վ˵����</td>
      <td colspan="2"><textarea name="info" cols="50" rows="3" id="info" onkeyUp="textLimitCheck(this, 50);"></textarea>&nbsp;&nbsp; �� 50 ���ַ� ������ <font color="#CC0000"><span id="messageCount">0</span></font> ����></td>
    </tr>
    <tr> 
      <td width="188" height="25" align="right">��վLOGO��</td>
      <td colspan="2"><input name="logo" type="text" class="input" id="logo" size="50" maxlength="255">û��������<br><iframe name="logo" frameborder=0 width="450" height="35" scrolling=no src="../DNJY_upload.asp?ttup=link"></iframe><br>&nbsp;&nbsp;<font color="#ff0000">�����ϴ����ļ����ͣ�gif|jpgs|jpg|bmp&nbsp;&nbsp;&nbsp;��jpg��ʽͼƬ�ɻõ���ʾ!&nbsp;&nbsp;�ޣ�88*31���أ�200k����</font> </td>
    </tr>

  <tr bgcolor="#eeeeee">
    <td height="35" colspan="10"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td><div align="center">
		
            <input type="submit" value="ȷ��" name="B1" >
        </div></td>
        <td><div align="center">
            <input type="reset" value="ȡ��" name="B1">
        </div></td>
      </tr>          
        </FORM> 
     
      </TABLE>
	<BR></TD>
  </TR>
</TABLE>
<%End if%>
</BODY> 
</HTML> 
