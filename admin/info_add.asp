<!--#include file="../conn.asp"-->
<!--#include file="../SqlIn.Asp"-->
<!--#include file=cookies.asp-->
<!--#include file=../config.asp-->
<!--#include file="../Include/head.asp"-->
<!--#include file="../Include/Function.asp"-->
<!--#include file="../Include/label.asp"-->
<!--#include file="../Include/err.asp"-->


<%
'=====================================
'���칩����Ϣ��վ����ϵͳ
'��������Ƽ������ҳ�Ʒ
'�ͷ�����http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>
<%call checkmanage("05")
Dim username,count,b,yz
Dim sqlu,rsu
set rsu=server.createobject("adodb.recordset")
sqlu = "select * from [DNJY_user] where username='"&request.cookies("administrator")&"'"
rsu.open sqlu,conn,1,1
if rsu.bof and rsu.eof Then
response.write "��̨����Ա������Ϣ����ǰ̨��ͬ�ʺŵĻ�Ա�������ڻ�Ա��������������Աͬ���Ļ�Ա�ʺţ�"
response.End
End if%>
<html>
<head>
<meta name="GENERATOR" content="Microsoft FrontPage 6.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>�����Ϣ</title>
<!--kindeditor�༭��-->
	<link rel="stylesheet" href="../KindEditor/themes/default/default.css" />
	<link rel="stylesheet" href="../KindEditor/plugins/code/prettify.css" />
	<script charset="utf-8" src="../KindEditor/kindeditor.js"></script>
	<script charset="utf-8" src="../KindEditor/lang/zh_CN.js"></script>
	<script charset="utf-8" src="../KindEditor/plugins/code/prettify.js"></script>
	<script>
		KindEditor.ready(function(K) {
			var editor = K.create('textarea[name="memo"]', {
				cssPath : '../KindEditor/plugins/code/prettify.css',
				uploadJson : '../KindEditor/asp/upload_json.asp?action=infopic',
				afterBlur: function(){this.sync();},//ʧȥ����ִ��this���ֵ
				allowFileManager : true,				
				afterCreate : function() {
					var self = this;
					K.ctrl(document, 13, function() {
						self.sync();
						K('form[name=example]')[0].submit();
					});
					K.ctrl(self.edit.doc, 13, function() {
						self.sync();
						K('form[name=example]')[0].submit();
					});
				}
			});
			prettyPrint();
		});
	</script>
<!--kindeditor�༭��END-->

<link rel="stylesheet" type="text/css" href="../css/css.css">
<script language=javascript src="../Include/mouse_on_title.js"></script>
<script src="../Include/prototype.js"></script>
<style type="text/css">
<!--
body {
	background-color: #E3EBF9;
}
.style1 {	color: #FFFFFF;
	font-weight: bold;
}
-->
</style>
<script language="javascript">
<!--
//��֤������ȷ��
function checkemail(email){
var str=email;
//��JavaScript�У�������ʽֻ��ʹ��"/"��ͷ�ͽ���������ʹ��˫����
var Expression=/\w+([-+.']\w+)*@\w+([-.]\w+)*\.\w+([-.]\w+)*/; 
var objExp=new RegExp(Expression);
if(objExp.test(str)==true){
return true;
}else{
return false;
}
} 


//˵���ֽ������ƣ����������=========
String.prototype.Length=function(){
temp=this.replace(/([^\x00-\xff])/g,"$1$1");
return temp.length
}
//��ʾ�����ַ���


function formcheck(){
var editor=KindEditor.create('#editor');
if (document.myform.city_one.value==0)
	{
	  alert("��ѡ�������")
	  document.myform.city_one.focus()
	  return false
	 }

if (document.myform.type_one.value==0)
	{
	  alert("��ѡ����Ϣһ�����࣡")
	  document.myform.type_one.focus()
	  return false
	 }
if (document.myform.biaoti.value=="" )
	{
	  alert("������Ϣ����")	  
	  document.myform.biaoti.focus()
	  return false
	 }	 
if (document.myform.biaoti.value.length>100)
	{
	  alert("���ⲻ�ܳ���100���ֽ�")
	  document.myform.biaoti.focus()
	  return false
	 }
if (document.myform.keywords.value=="" )
	{
	  alert("������Ϣ�ؼ��ʣ�")
	  document.myform.keywords.focus()
	  return false
	 }
if (document.myform.keywords.value.length>40)
	{
	  alert("�ؼ��ʲ��ܳ���40���ַ�")
	  document.myform.keywords.focus()
	  return false
	 }	 
if (document.myform.leixing.value=="" )
	{
	  alert("������Ϣ����")
	  document.myform.leixing.focus()
	  return false
	 }
	 
//if (document.myform.scjiage.value=="" )
//	{
//	  alert("�����г��۸�")
//	  document.myform.scjiage.focus()
//	  return false
//	 }
//if (isNaN(document.myform.scjiage.value))
//	{
//	  alert("�г��۸�ӦΪ���֣�")
//	  document.myform.scjiage.focus()
//	  return false
//	 }	
if (document.myform.jiage.value=="" )
	{
	  alert("����׼۸�")
	  document.myform.jiage.focus()
	  return false
	 }		 
if (isNaN(document.myform.jiage.value))
	{
	  alert("���׼۸�ӦΪ���֣�")
	  document.myform.jiage.focus()
	  return false
	 }
var editor=KindEditor.create('#editor');
if (editor.isEmpty())
	{
	  alert("��Ϣ˵������Ϊ�գ�")	  
	  return false
	 }
if (document.myform.memo.value.Length()><%=memonum%>)  //�ֽ������ƣ����������
     {
	 alert("��Ϣ˵������Ҫ��<%=memonum%>�ֽ�(<%=memonum/2%>����)�����������룡");
	  return false
  }	
if (document.myform.name.value=="" )
	{
	  alert("������ʵ����")
	  document.myform.name.focus()
	  return false
	 }
	if(document.myform.dianhua.value=="" && document.myform.CompPhone.value=="") 
	{
	    alert("�̶��绰���ƶ��绰����ͬʱΪ��!");
		document.myform.dianhua.focus()
	    return false;
	}
//var sMobile = document.myform.dianhua.value
//if((document.myform.dianhua.value!="") && (!(/^([0\+]\d{2,3}-)?(0[0-9]{2,3}\-)?([2-9][0-9]{6,7})+(\-[0-9]{1,4})?$/.test(sMobile)))) //(���Ҵ���-����-�绰-�ֻ���ʽ(��ѡ)��+086-010-85858585-2538)
//	{
//		alert("���淶�ĵ绰����");
//		document.myform.dianhua.focus();
//		return false;
//	}		
//	var sMobile = document.myform.CompPhone.value
//	if((document.myform.CompPhone.value!="") && (!(/^((\(\d{3}\))|(\d{3}\-))?13[0-9]\d{8}|15[0-9]\d{8}|18[0-9]\d{8}/.test(sMobile)))) //(����13��15��18�Ŷ�)
//	{
//		alert("����������11λ�ֻ��Ż�����ȷ���ֻ���ǰ��λ");
//		document.myform.CompPhone.focus();
//		return false;
//	}
    if((!checkemail(myform.email.value))&&(document.myform.email.value!=""))
	{
    alert("������Email��ַ����ȷ������������!");
    document.myform.email.focus();
    return false;
    }
   if(document.myform.email.value.length>30 )
	{
	    alert("���䲻�ܳ���30���ַ�!");
	    document.myform.email.focus();
	    return false;
	}
    if (document.myform.hfje.value=="" )
	{	  
      alert("���۽���Ϊ�գ��������������Ϊ0��");
	  document.myform.hfje.focus()
	  return false
	 }		
    if (document.myform.b.value=="" )
	{	  
      alert("�ö���ֵ����Ϊ�գ��������������Ϊ0��");
	  document.myform.b.focus()
	  return false
	 }		  
}
// --></script>

</head>
<BODY>

<center>
  <table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#ADAED6" width="800" >

  <TR>
    <TD height="20" bgcolor="#799AE1" align="center" colspan="13"><FONT color="#FFFFFF" style="font-size:14px">����Ա<%=request.cookies("administrator")%>������Ϣ</FONT></TD>
 </TR>
  <table width="800" border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse">
    <tr>

      <td width="800" align="right" valign="top" bgcolor="#E3EBF9"><table width="800"  border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td><form action="info_add_save.asp" name="myform" method="POST" language="javascript" onSubmit="return formcheck()">
            <table width="100%" height="100%" border="0" align="right" cellpadding="6" cellspacing="0" bgcolor="#D6E0F9">
                
 <%dim biaoti,scjiage,jiage,memo,uname,name,email,dianhua,zt,CompPhone,youbian,dizhi,a,dqsj,hfje,homeEliteImage%>              
                <tr>
                  <td height="10" colspan="2"></td>
                </tr>
                        <tr>
                          <td height="25" width="100" align="right"> ���׵�����</td>
                          <td height="25" width="700">
<%set rs=conn.execute("select * from DNJY_city where id>0 and twoid>0 and threeid=0 order by indexid")
%>
      <SCRIPT language = "JavaScript">
var onecount;
onecount=0;
subcat = new Array();
        <%
		//dim count:
		count = 0
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
set rs=conn.execute("select * from DNJY_city where id>0 and twoid>0 and threeid>0 order by indexid")
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
    document.myform.city_two.length = 0; 
	document.myform.city_two.options[0] = new Option('ѡ�����','');
	document.myform.city_three.length = 0; 
	document.myform.city_three.options[0] = new Option('ѡ�����','');
    var locationid=locationid;
    var i;
    for (i=0;i < onecount; i++)
        {
            if (subcat[i][1] == locationid)
            { 
                document.myform.city_two.options[document.myform.city_two.length] = new Option(subcat[i][0], subcat[i][2]);
            }        
        }
        
    }    
	
	function changelocation4(locationid,locationid1)
    {
    document.myform.city_three.length = 0; 
    document.myform.city_three.options[0] = new Option('ѡ�����','');
    var locationid=locationid;
	 var locationid1=locationid1;
    var i;
    for (i=0;i < onecount4; i++)
        {
            if (subcat4[i][2] == locationid)
            { 
			if (subcat4[i][1] == locationid1)
			{
                document.myform.city_three.options[document.myform.city_three.length] = new Option(subcat4[i][0], subcat4[i][3]);
            }        
        }
       } 
    }            
              </SCRIPT>
      <SELECT name="city_one" size="1" class="inputa" id="select2" onChange="changelocation(document.myform.city_one.options[document.myform.city_one.selectedIndex].value)">
            <OPTION value="" selected>ѡ�����</OPTION>
  <%set rs=conn.execute("select * from DNJY_city where id>0 and twoid=0 and threeid=0 order by indexid")
if rs.eof or rs.bof then
response.write "<option value=''>û�з���</option>"
else
do until rs.eof%>
  <OPTION value="<%=rs("id")%>" <%if rs("id")=rsu("city_oneid") then%>selected<%end if%>><%=rs("city")%></OPTION>
 <%rs.movenext
    loop
	%>
	<%end if
rs.close
set rs = nothing
%>
      </SELECT> 
	  <SELECT name="city_two" id="city_two" class="inputa" onChange="changelocation4(document.myform.city_two.options[document.myform.city_two.selectedIndex].value,document.myform.city_one.options[document.myform.city_one.selectedIndex].value)">
    <OPTION value="" selected>ѡ�����</OPTION>
   <%
set rs=conn.execute("select * from DNJY_city where id="&rsu("city_oneid")&" and twoid>0 and threeid=0")
if rs.eof and rs.bof then
response.write "<option value=''>û�з���</option>"
else
do until rs.eof%>
  <OPTION value="<%=rs("twoid")%>" <%if rs("twoid")=rsu("city_twoid") then%>selected<%end if%>><%=rs("city")%></OPTION>
 <%rs.movenext
    loop
	%>
	<%end if
rs.close
set rs = nothing
%>
	</SELECT>
	     <SELECT name="city_three" id="city_three" class="inputa">
         <OPTION value="" selected>ѡ�����</OPTION>
		<%
set rs=conn.execute("select * from DNJY_city where id="&rsu("city_oneid")&" and twoid="&rsu("city_twoid")&" and threeid<>0")
if rs.eof and rs.bof then
response.write "<option value=''>û�з���</option>"
else
do until rs.eof%>
<OPTION value="<%=rs("threeid")%>" <%if rs("threeid")=rsu("city_threeid") then%>selected<%end if%>><%=rs("city")%></OPTION>
   <% rs.movenext
    loop
	%>
<%end if
rs.close
set rs = nothing
%>
    </SELECT>

<font color="#ff0000"> *</font></td>
                        </tr>
						<tr>
                          <td height="25" width="100" align="right">��Ϣ���</td>
                          <td height="25" width="500">
<%set rs=conn.execute("select * from DNJY_type where id>0 and twoid>0 and threeid=0 order by indexid")%>
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
    document.myform.type_two.options[0] = new Option('��Ϣ��������','');
	document.myform.type_three.length = 0; 
    document.myform.type_three.options[0] = new Option('��Ϣ��������','');
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
    document.myform.type_three.options[0] = new Option('��Ϣ��������','');
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
                                 <SELECT name="type_one" size="1" id="select3" class="inputa" onChange="changelocation2(document.myform.type_one.options[document.myform.type_one.selectedIndex].value)">
                                   <OPTION value="" selected>��Ϣһ������</OPTION>
                                   <%set rs=conn.execute("select * from DNJY_type where id>0 and twoid=0 order by indexid asc")
if rs.eof and rs.bof then
response.write "<option value=''>û�з���</option>"
else
do until rs.eof%>
                                   <OPTION value="<%=rs("id")%>" <%if rs("id")=rsu("type_oneid") then%>selected<%end if%>><%=rs("name")%></OPTION>
                                   <% rs.movenext
    loop
	end if
rs.close
set rs= nothing
%>
                                 </SELECT>
                                 <SELECT name="type_two" id="select4" class="inputa" onChange="changelocation3(document.myform.type_two.options[document.myform.type_two.selectedIndex].value,document.myform.type_one.options[document.myform.type_one.selectedIndex].value)">
                                   <OPTION value="" selected>��Ϣ��������</OPTION>
                                   <%
	set rs=conn.execute("select * from DNJY_type where id="&rsu("type_oneid")&" and twoid<>0 and threeid=0")
if rs.eof and rs.bof then
response.write "<option value=''>û�з���</option>"
else
do until rs.eof%>
                                   <OPTION value="<%=rs("twoid")%>" <%if rs("twoid")=rsu("type_twoid") then%>selected<%end if%>><%=rs("name")%></OPTION>
                                   <% rs.movenext
    loop
	end if
rs.close
set rs= nothing%>
                                 </SELECT>
                                 <SELECT name="type_three" id="select5"  class="inputa">
                                   <OPTION value="" selected>��Ϣ��������</OPTION>
                                   <%set rs=conn.execute("select * from DNJY_type where id="&rsu("type_oneid")&" and twoid="&rsu("type_twoid")&" and threeid<>0")
if rs.eof and rs.bof then
response.write "<option value=''>û�з���</option>"
else
do until rs.eof%>
                                   <OPTION value="<%=rs("threeid")%>" <%if rs("threeid")=rsu("type_threeid") then%>selected<%end if%>><%=rs("name")%></OPTION>
                                   <% rs.movenext
    loop
	end if
rs.close
set rs= Nothing
%>
                                 </SELECT> <font color="#ff0000"> *</font></td>
                        </tr>
                <tr>
                  <td height="25" align="right">��Ϣ���⣺</td>
                  <td height="25"><input style="BACKGROUND-COLOR: #ffffff; BORDER-BOTTOM: #FFFFFF 1px solid; BORDER-LEFT: #a2a2a2 1px solid; BORDER-RIGHT: #ffffff 1px solid; BORDER-TOP: #a2a2a2 1px solid; COLOR: #333333; FONT-SIZE: 9pt; HEIGHT: 19px;" type="text" maxlength="100" name="biaoti" size="50" value="">
                    ��<font color="#ff0000"> *</font><font color="#CC5200">30������</font>��
         
        </font> ������ɫ��<font face="����">
          
          <select name="a" class="inputa">
            <option style="COLOR: black" value="0" selected>��ʹ�� </option>
            <option style="background: #000088" value="000088"> </option>
            <option style="background: #0000ff" value="0000ff"> </option>
            <option style="background: #008800" value="008800"> </option>
            <option style="background: #008888" value="008888"> </option>
            <option style="background: #0088ff" value="0088ff"> </option>
            <option style="background: #00a010" value="00a010"> </option>
            <option style="background: #1100ff" value="1100ff"> </option>
            <option style="background: #111111" value="111111"> </option>
            <option style="background: #333333" value="333333"> </option>
            <option style="background: #50b000" value="50b000"> </option>
            <option style="background: #880000" value="880000"> </option>
            <option style="background: #8800ff" value="8800ff"> </option>
            <option style="background: #888800" value="888800"> </option>
            <option style="background: #888888" value="888888"> </option>
            <option style="background: #8888ff" value="8888ff"> </option>
            <option style="background: #aa00cc" value="aa00cc"> </option>
            <option style="background: #aaaa00" value="aaaa00"> </option>
            <option style="background: #ccaa00" value="ccaa00"> </option>
            <option style="background: #ff0000" value="ff0000"> </option>
            <option style="background: #ff0088" value="ff0088"> </option>
            <option style="background: #ff00ff" value="ff00ff"> </option>
            <option style="background: #ff8800" value="ff8800"> </option>
            <option style="background: #ff0005" value="ff0005"> </option>
            <option style="background: #ff88ff" value="ff88ff"> </option>
            <option style="background: #ee0005" value="ee0005"> </option>
            <option style="background: #ee01ff" value="ee01ff"> </option>
            <option style="background: #3388aa" value="3388aa"> </option>
            <option style="background: #000000" value="000000"> </option>
          </select>         
        </font></td>
                </tr>
            <tr>
            <td height="25" width="100" align="right">�� �� �ʣ�</td>
            <td height="25" width="450"><input style="BACKGROUND-COLOR: #ffffff; BORDER-BOTTOM: #FFFFFF 1px solid; BORDER-LEFT: #a2a2a2 1px solid; BORDER-RIGHT: #ffffff 1px solid; BORDER-TOP: #a2a2a2 1px solid; COLOR: #333333; FONT-SIZE: 9pt; HEIGHT: 19px;" type="text" maxlength="100" name="keywords" size="50" value="">
            <font color="#ff0000"> *</font> <input type="button" name="btn" value="�ӱ��⸴��" title="���Ʊ��⵽�ؼ���" onClick="CopyWebbiaoti(document.myform.biaoti.value);"><br>��<font color="#CC5200">��Ϣ�ؼ��ʣ��á�,���ŷָ�����100�ֽ��ڡ�</font>��</td>
              </tr>
                <tr>
                  <td height="25" align="right"> <font color='#0000FF'>��ҳͷ����</font></td>
                  <td height="25"> <input name="homeEliteImage" type="text" id="homeEliteImage" value="" size="50" maxlength="255" readonly onclick="sHomeElite();">
                  <input name="isHomeElite" type="checkbox" id="isHomeElite" value="Yes" onClick="rHomeElite();"> <font color='#FF0000'>ʹ��ͷ�����</font>(��Ѱ��޴˹���)<br><span id="span_createImage"></span></td>
                </tr>
              <tr>
                  <td height="25" align="right"><font color='#0000FF'>ͷ��������</font></td>
				  <td height="25">
                  ��<input name="homeEliteWidth" type="text" id="homeEliteWidth" value="<%=HomeEliteWidth%>" size="4" maxlength="4" />
                  �ߣ�<input name="homeEliteHeight" type="text" id="homeEliteHeight" value="<%=HomeEliteHeight%>" size="4" maxlength="4" />
                  ���壺<input name="homeEliteFontFamily" type="text" id="homeEliteFontFamily" value="<%=HomeEliteFontFamily%>" size="12" maxlength="4" />
                  �����С��<input name="homeEliteFontSize" type="text" id="homeEliteFontSize" value="<%=HomeEliteFontSize%>" size="4" maxlength="4" />
                  ��ɫ��<select name="homeEliteColor" id="homeEliteColor">
                          <option value="000000" selected>��ɫ</option>
                          <option value="000000" style="background-color:#000000"></option>
                          <option value="FFFFFF" style="background-color:#FFFFFF"></option>
                          <option value="008000" style="background-color:#008000"></option>
                          <option value="800000" style="background-color:#800000"></option>
                          <option value="808000" style="background-color:#808000"></option>
                          <option value="000080" style="background-color:#000080"></option>
                          <option value="800080" style="background-color:#800080"></option>
                          <option value="808080" style="background-color:#808080"></option>
                          <option value="FFFF00" style="background-color:#FFFF00"></option>
                          <option value="00FF00" style="background-color:#00FF00"></option>
                          <option value="00FFFF" style="background-color:#00FFFF"></option>
                          <option value="FF00FF" style="background-color:#FF00FF"></option>
                          <option value="FF0000" style="background-color:#FF0000"></option>
                          <option value="0000FF" style="background-color:#0000FF"></option>
                          <option value="008080" style="background-color:#008080"></option>
                        </select>
                </td>
              </tr>

                <tr>
                  <td height="25" align="right"> �������ͣ�</td>
                  <td height="25"><%
Dim rslx,sqllx,x,leixinglx,selectedlx
set rslx=server.createobject("adodb.recordset")
sqllx = "select leixing from DNJY_config "
rslx.open sqllx,conn,1,1
leixinglx=split(rslx("leixing"),"|")
selectedlx=""
response.write "<select name=""leixing""><option value=>����</option>"
for x=0 to ubound(leixinglx)
response.write "<option value="""&leixinglx(x)&""">"&leixinglx(x)&"</option>"
next
response.write "</select>"
rslx.close
set rslx=nothing%> <font color="#ff0000"> *</font></td>
                </tr>
                <!--<tr>
                  <td height="25"> �г��۸�</td>
                  <td height="25"><input style="BACKGROUND-COLOR: #ffffff; BORDER-BOTTOM: #FFFFFF 1px solid; BORDER-LEFT: #a2a2a2 1px solid; BORDER-RIGHT: #ffffff 1px solid; BORDER-TOP: #a2a2a2 1px solid; COLOR: #333333; FONT-SIZE: 9pt; HEIGHT: 19px;" type="text" name="scjiage" size="7" maxlength="6" value="" onkeyup="value=value.replace(/[^\d|^\.]/g,'')"   onbeforepaste="clipboardData.setData('text',clipboardData.getData('text').replace(/[^\d|^\.]/g,''))">
                    Ԫ&nbsp;&nbsp;&nbsp; <font color="#ff0000"> *</font>��0<font color="#CC5200">Ԫ��ʾ����</font>��</td>
                </tr>-->
                <tr>
                  <td height="25" align="right"> ���׼۸�</td>
                  <td height="25"><input style="BACKGROUND-COLOR: #ffffff; BORDER-BOTTOM: #FFFFFF 1px solid; BORDER-LEFT: #a2a2a2 1px solid; BORDER-RIGHT: #ffffff 1px solid; BORDER-TOP: #a2a2a2 1px solid; COLOR: #333333; FONT-SIZE: 9pt; HEIGHT: 19px;" type="text" name="jiage" size="7" maxlength="6" value="0" onkeyup="value=value.replace(/[^\d|^\.]/g,'')"   onbeforepaste="clipboardData.setData('text',clipboardData.getData('text').replace(/[^\d|^\.]/g,''))">
                    Ԫ&nbsp;&nbsp;&nbsp; <font color="#ff0000"> *</font>��0<font color="#CC5200">Ԫ��ʾ����</font>��</td>
                </tr>
	   <tr> 
      <td width="188" height="25" align="right">��Ϣ��ͼ��</td>
      <td><input name="tpname" type="text" class="input" id="tpname" size="50" maxlength="255">��Ҳ����������Ϣ˵������ȡ���ݵ�һ��ͼƬΪ��ͼ��<br><iframe name="tpname" frameborder=0 width="400" height="35" scrolling=no src="../DNJY_upload.asp?ttup=info"></iframe><br>&nbsp;&nbsp;<font color="#ff0000">�����ϴ����ļ����ͣ�gif|jpgs|jpg|bmp|png 500k����</font> </td>
    </tr>   
                <tr>
                  <td height="25" align="right"><p>��Ϣ˵����</td>
                  <td height="350" width="700"><textarea id="editor" name="memo" style="width:670px;height:400px;"></textarea><input type="checkbox" name="T_SaveImg" value="1" /> ����������Զ��ͼƬ������&nbsp;&nbsp;<input type="checkbox" name="AspJpeg" value="1" checked/>���ϴ�ͼƬ��ˮӡ&nbsp;&nbsp;<input type="checkbox" name="T_conpic" value="1" />��ȡ���ݵ�һ��ͼƬΪ��ͼ�������ϴ���ͼ������Ч��</td>
       </tr> 

       <!--<tr>
                  <td height="25" align="right">�����ɫ��</td>
                  <td height="25"><input style="BACKGROUND-COLOR: #ffffff; BORDER-BOTTOM: #FFFFFF 1px solid; BORDER-LEFT: #a2a2a2 1px solid; BORDER-RIGHT: #ffffff 1px solid; BORDER-TOP: #a2a2a2 1px solid; COLOR: #333333; FONT-SIZE: 9pt; HEIGHT: 19px;" type="text" name="wcolor" size="23" maxlength="15" value=""> &lt;=��<FONT color=blue><FONT 
      style="CURSOR: hand" onclick="wcolor.value='��ɫ'" 
      color=#993300>��ɫ</FONT></FONT>����<FONT 
      color=blue><FONT style="CURSOR: hand" onclick='wcolor.value="��ɫ"' 
      color=#993300>��ɫ</FONT></FONT>����<FONT color=blue><FONT 
      style="CURSOR: hand" onclick="wcolor.value='��ɫ'" 
      color=#993300>��ɫ</FONT></FONT>����<FONT color=blue><FONT 
      style="CURSOR: hand" onclick="wcolor.value='��ɫ'" 
      color=#993300>��ɫ</FONT></FONT>����<FONT color=blue><FONT 
      style="CURSOR: hand" onclick="wcolor.value='��ɫ'" 
      color=#993300>��ɫ</FONT></FONT>��</td>
                </tr>
                <tr>
                  <td height="25" align="right">������ɫ��</td>
                  <td height="25"><input style="BACKGROUND-COLOR: #ffffff; BORDER-BOTTOM: #FFFFFF 1px solid; BORDER-LEFT: #a2a2a2 1px solid; BORDER-RIGHT: #ffffff 1px solid; BORDER-TOP: #a2a2a2 1px solid; COLOR: #333333; FONT-SIZE: 9pt; HEIGHT: 19px;" type="text" name="ncolor" size="23" maxlength="15" value="">  &lt;=��<FONT color=blue><FONT 
      style="CURSOR: hand" onclick="ncolor.value='��ɫ'" 
      color=#993300>��ɫ</FONT></FONT>����<FONT 
      color=blue><FONT style="CURSOR: hand" onclick='ncolor.value="��ɫ"' 
      color=#993300>��ɫ</FONT></FONT>����<FONT color=blue><FONT 
      style="CURSOR: hand" onclick="ncolor.value='��ɫ'" 
      color=#993300>��ɫ</FONT></FONT>����<FONT color=blue><FONT 
      style="CURSOR: hand" onclick="ncolor.value='��ɫ'" 
      color=#993300>��ɫ</FONT></FONT>����<FONT color=blue><FONT 
      style="CURSOR: hand" onclick="ncolor.value='��ɫ'" 
      color=#993300>��ɫ</FONT></FONT>��</td>
                </tr>-->
                      <tr>
                        <td height="25" align="right">����������</td>
                        <td width="500" height="25"><select name="days" style="BACKGROUND-COLOR: #FFCCFF; BORDER-BOTTOM: #FFFFFF 1px solid; BORDER-LEFT: #a2a2a2 1px solid; BORDER-RIGHT: #ffffff 1px solid; BORDER-TOP: #a2a2a2 1px solid; COLOR: #333333; FONT-SIZE: 9pt; HEIGHT: 19px;">
						  <option value="1">һ��</option>
						  <option value="2">����</option>
						  <option value="3">����</option>
						  <option value="4">����</option>
						  <option value="5">����</option>
						  <option value="6">����</option>
                          <option value="7">һ����</option>
						  <option value="14">������</option>
						  <option value="21">������</option>
                          <option value="30">һ����</option>
						  <option value="60">������</option>
                          <option value="90">������</option>
					      <option value="180">����</option>
						  <option value="365">һ��</option>
                          <option value="1100">����</option>
                          </select>
                            </font></td>
                      </tr>

<%Dim sqljj,rsjj,Amount
set rsjj=server.createobject("adodb.recordset")
sqljj = "select * from [DNJY_user] where username='"&request.cookies("dick")("username")&"'"
rsjj.open sqljj,conn,1,1
if rsjj.eof Then
Amount=0
else
Amount=rsjj("Amount")
End if
rsjj.close
set rsjj=nothing%>
                     <tr>
                        <td height="25" align="right">���۽�</td>
                        <td  height="-1"><input style="BACKGROUND-COLOR: #ffffff; BORDER-BOTTOM: #FFFFFF 1px solid; BORDER-LEFT: #a2a2a2 1px solid; BORDER-RIGHT: #ffffff 1px solid; BORDER-TOP: #a2a2a2 1px solid; COLOR: #333333; FONT-SIZE: 9pt; HEIGHT: 19px;" type="text" name="hfje" size="23" value="0" onkeyup="this.value=this.value.replace(/\D/g,'')" onafterpaste="this.value=this.value.replace(/\D/g,'')"><font color="#ff0000"> *</font>�� ���ӵľ��۽��뾺�������ӡ������ʽ��ʻ����<%=Amount%>Ԫ</span>��</td>
                      </tr> 
                      <tr> 
                        <td height="25" align="right">��Ϣ�ö���</td>
                        <td height="0" >
                          <input type="text" name="b" size="23" maxlength="40" value="0" style="BACKGROUND-COLOR: #ffffff; BORDER-BOTTOM: #FFFFFF 1px solid; BORDER-LEFT: #a2a2a2 1px solid; BORDER-RIGHT: #ffffff 1px solid; BORDER-TOP: #a2a2a2 1px solid; COLOR: #333333; FONT-SIZE: 9pt; HEIGHT: 19px;" onkeyup="this.value=this.value.replace(/\D/g,'')" onafterpaste="this.value=this.value.replace(/\D/g,'')">
                          <font color="#ff0000"> *</font>���ö�������ҳ������Ϣ����ʾ������Խ������Խ�ߣ�</td>
                      </tr> 
                      <tr> 
                        <td height="25" align="right">�Ķ�Ȩ�ޣ�</td>
                        <td height="0" ><input type="radio" name="Readinfo" value="0" checked>�ο�	<input type="radio" name="Readinfo" value="1">��ͨ��Ա	<input type="radio" name="Readinfo" value="2">VIP��Ա	
                             <br>��������Ϣ��ϸ˵�������緽ʽ����Ա�������ͼ��ݣ�</td>
                      </tr>
					    <!--<tr>
                        <td height="25" align="right">�Ķ�Ȩ�ޣ�</td>
                        <td width="400" height="25"><input type="radio" name="Readinfo" value="1" checked>��ͨ��Ա	<input type="radio" name="Readinfo" value="2">VIP��Ա	
                             <br>��������Ϣ��ϸ˵�������緽ʽ����Ա�������ͼ��ݣ�</td>
                      </tr>--->					  
                      <tr> 
                        <td height="25" align="right">�����֤��</td>
                        <td height="0" >
                          <input type="radio" value="1" name="yz" checked>����ͨ��</font><input type="radio" value="0" name="yz">�ݲ�ͨ��</td>
                      </tr>	
                <tr>
                  <td height="25" align="right">��ʵ������</td>
                  <td height="25"><input style="BACKGROUND-COLOR: #ffffff; BORDER-BOTTOM: #FFFFFF 1px solid; BORDER-LEFT: #a2a2a2 1px solid; BORDER-RIGHT: #ffffff 1px solid; BORDER-TOP: #a2a2a2 1px solid; COLOR: #333333; FONT-SIZE: 9pt; HEIGHT: 19px;" type="text" name="name" size="23" maxlength="15" value="<%=rsu("name")%>"> <font color="#ff0000"> *</font></td>
                </tr>
                       <tr>
                        <td height="25" align="right">�ƶ��绰��</td>
                        <td height="25"><input style="BACKGROUND-COLOR: #ffffff; BORDER-BOTTOM: #FFFFFF 1px solid; BORDER-LEFT: #a2a2a2 1px solid; BORDER-RIGHT: #ffffff 1px solid; BORDER-TOP: #a2a2a2 1px solid; COLOR: #333333; FONT-SIZE: 9pt; HEIGHT: 19px;" type="text" name="CompPhone" size="23" maxlength="40" value="<%=rsu("CompPhone")%>" onkeyup="this.value=this.value.replace(/\D/g,'')" onafterpaste="this.value=this.value.replace(/\D/g,'')"> <font color="#0000ff"> *</font>��<font color="#CC5200">�̶��绰���ƶ��绰������һ</font>��</td>
                      </tr>
                <tr>
                  <td height="25" align="right"> �̶��绰��</td>
                  <td height="25"><input style="BACKGROUND-COLOR: #ffffff; BORDER-BOTTOM: #FFFFFF 1px solid; BORDER-LEFT: #a2a2a2 1px solid; BORDER-RIGHT: #ffffff 1px solid; BORDER-TOP: #a2a2a2 1px solid; COLOR: #333333; FONT-SIZE: 9pt; HEIGHT: 19px;" type="text" name="dianhua" size="23" maxlength="40" value="<%=rsu("dianhua")%>"> <font color="#0000ff"> *</font>��<font color="#CC5200">�̶��绰���ƶ��绰������һ</font>��<br>(���Ҵ���-����-�绰-�ֻ���ʽ(��ֻ������+�绰��绰)��+086-010-85858585-2538)</td>
                </tr>

                <tr>
                  <td height="25" align="right"> ������룺</td>
                  <td height="25"><input style="BACKGROUND-COLOR: #ffffff; BORDER-BOTTOM: #FFFFFF 1px solid; BORDER-LEFT: #a2a2a2 1px solid; BORDER-RIGHT: #ffffff 1px solid; BORDER-TOP: #a2a2a2 1px solid; COLOR: #333333; FONT-SIZE: 9pt; HEIGHT: 19px;" type="text" name="fax" size="23" maxlength="40" value="<%=rsu("fax")%>"> (���Ҵ���-����-�绰-�ֻ���ʽ(��ֻ������+�绰��绰)��+086-010-85858585-2538)</td>
                </tr>
                <tr>
                  <td height="25" align="right">�������䣺</td>
                  <td height="25"><input style="BACKGROUND-COLOR: #ffffff; BORDER-BOTTOM: #FFFFFF 1px solid; BORDER-LEFT: #a2a2a2 1px solid; BORDER-RIGHT: #ffffff 1px solid; BORDER-TOP: #a2a2a2 1px solid; COLOR: #333333; FONT-SIZE: 9pt; HEIGHT: 19px;" type="text" name="email" size="23" maxlength="40" value="<%=rsu("email")%>">
                   ��<font color="#CC5200">���ʼ���ַ��������Ľ�����Ϣ</font>��</td>
                </tr>					  
                      <tr>
                        <td height="25" align="right">QQ �� �룺</td>
                        <td width="500" height="-1"><input style="BACKGROUND-COLOR: #ffffff; BORDER-BOTTOM: #FFFFFF 1px solid; BORDER-LEFT: #a2a2a2 1px solid; BORDER-RIGHT: #ffffff 1px solid; BORDER-TOP: #a2a2a2 1px solid; COLOR: #333333; FONT-SIZE: 9pt; HEIGHT: 19px;" type="text" name="qq" size="23" maxlength="40" value="<%=rsu("qq")%>" onkeyup="this.value=this.value.replace(/\D/g,'')" onafterpaste="this.value=this.value.replace(/\D/g,'')"></td>
                      </tr>
                        <tr>
                          <td height="25" align="right"> ΢�ź��룺</td>
                          <td height="25" width="408"><input name="weixin" type="text" style="BACKGROUND-COLOR: #ffffff; BORDER-BOTTOM: #FFFFFF 1px solid; BORDER-LEFT: #a2a2a2 1px solid; BORDER-RIGHT: #ffffff 1px solid; BORDER-TOP: #a2a2a2 1px solid; COLOR: #333333; FONT-SIZE: 9pt; HEIGHT: 19px;" size="23" maxlength="50" value="<%=rsu("weixin")%>">
                            &nbsp; </td>
                        </tr>						  
                      <tr>
                        <td height="25" align="right">�������룺</td>
                        <td width="500" height="-1"><input style="BACKGROUND-COLOR: #ffffff; BORDER-BOTTOM: #FFFFFF 1px solid; BORDER-LEFT: #a2a2a2 1px solid; BORDER-RIGHT: #ffffff 1px solid; BORDER-TOP: #a2a2a2 1px solid; COLOR: #333333; FONT-SIZE: 9pt; HEIGHT: 19px;" type="text" name="youbian" size="23" maxlength="6" value="<%=rsu("youbian")%>" onkeyup="this.value=this.value.replace(/\D/g,'')" onafterpaste="this.value=this.value.replace(/\D/g,'')"></td>
                      </tr>
                      <tr>
                        <td height="25" width="100" align="right">��˾���ƣ�</td>
                        <td height="25" width="500">
                            <input name="mpname" type="text" style="BACKGROUND-COLOR: #ffffff; BORDER-BOTTOM: #FFFFFF 1px solid; BORDER-LEFT: #a2a2a2 1px solid; BORDER-RIGHT: #ffffff 1px solid; BORDER-TOP: #a2a2a2 1px solid; COLOR: #333333; FONT-SIZE: 9pt; HEIGHT: 19px;" size="50" maxlength="50" value="<%=rsu("mpname")%>">
                          &nbsp; 
                          </td>
                      </tr>					  
                      <tr>
                        <td height="25" width="100" align="right">��ϵ��ַ��</td>
                        <td height="25" width="500">
                            <input name="dizhi" type="text" style="BACKGROUND-COLOR: #ffffff; BORDER-BOTTOM: #FFFFFF 1px solid; BORDER-LEFT: #a2a2a2 1px solid; BORDER-RIGHT: #ffffff 1px solid; BORDER-TOP: #a2a2a2 1px solid; COLOR: #333333; FONT-SIZE: 9pt; HEIGHT: 19px;" size="50" maxlength="100" value="<%=rsu("dizhi")%>">
                          &nbsp; 
                          </td>
                      </tr>

				  
				  
                  <td height="1" colspan="2"></td>
                </tr>
                <tr>
                  <td height="4" colspan="2"></td>
                </tr>
                <tr>
                  <td height="25" colspan="2" align="center"><table width="50%"  border="0" cellspacing="0" cellpadding="0">
                      <tr>
                        <td><div align="center">
                            <input name="I2" type="submit" style="BACKGROUND-COLOR: #ffffff; BORDER-BOTTOM: #a2a2a2 1px solid; BORDER-LEFT: #ffffff 1px solid; BORDER-RIGHT: #a2a2a2 1px solid; BORDER-TOP: #ffffff 1px solid; COLOR: #333333; FONT-SIZE: 9pt; HEIGHT: 19px;" onClick="javascript:return formcheck();" value="ȷ��" border="0">
                        </div></td>
                        <td><div align="center">
                            <input name="I2" type="reset" style="BACKGROUND-COLOR: #ffffff; BORDER-BOTTOM: #a2a2a2 1px solid; BORDER-LEFT: #ffffff 1px solid; BORDER-RIGHT: #a2a2a2 1px solid; BORDER-TOP: #ffffff 1px solid; COLOR: #333333; FONT-SIZE: 9pt; HEIGHT: 19px;" value="���" border="0">
                        </div></td>
                      </tr>
                  </table></td>
                </tr>
              </table>
          </form></td>
        </tr>
      </table></td>
      <td width="4" align="right" valign="top" bgcolor="#e6e6e6"></td>
    </tr>   
  </table>
  </center>
</div>
</body>
</html>
<%rsu.close:set rsu = nothing
Call CloseDB(conn)%>   