<!--#include file="conn.asp"-->
<!--#include file=top.asp-->
<!--#include file="Include/err.asp"-->
<%
'=====================================
'���칩����Ϣ��վ����ϵͳ
'��������Ƽ������ҳ�Ʒ
'�ͷ�����http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>
<%
if not ChkPost() then 
response.write "��ֹվ���ύ�����"
response.end 
end if
if lmkg2="1" then
call errdick(50)
response.end
end If

if addip<>"0" then
if ip<>"" then
'ip=checktext(ip)
addip=split(cstr(ip),"@")
for N=0 to UBound(addip)
if instr(Request.serverVariables("REMOTE_ADDR"),addip(n))>0 then
errdick(43)
response.end
end if
next
end if
end If

if addxinxi=0 then
response.redirect "login.asp"
end if
%>
<%'��ֹ��ҳ����
'Response.Buffer = True 
'Response.ExpiresAbsolute = Now() - 1 
'Response.Expires = 0 
'Response.CacheControl = "no-cache"
%>
<link rel="stylesheet" type="text/css" href="/<%=strInstallDir%>css/css.css">

<!--������֤����ÿ�ʼ-->
<SCRIPT LANGUAGE=javascript>
/*��ʾ��֤�� o start1*/
function get_Code() {
        var Dv_CodeFile = "Dv_GetCode.asp";
        if(document.getElementById("imgid"))
                document.getElementById("imgid").innerHTML = '<img src="'+Dv_CodeFile+'?t='+Math.random()+'" alt="���ˢ����֤��" style="cursor:pointer;border:0;vertical-align:middle;height:30px;" onclick="this.src=\''+Dv_CodeFile+'?t=\'+Math.random()" />'
}
/*o end*/
</script>
<script language="JavaScript" type="text/javascript">
var dvajax_request_type = "GET";
</script>
<script language="JavaScript" src="dv_ajax.js" type="text/javascript"></script>
<script language=javascript src=Include/mouse_on_title.js></script>
<!---������֤����ý���-->
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

//�����ֽ������ƣ����������=========
String.prototype.Length=function(){
temp=this.replace(/([^\x00-\xff])/g,"$1$1");
return temp.length
}
//��ʾ�����ַ���

function kn()
  {
   var ln=myform.biaoti.value.Length();
     num.innerText='������������:'+(<%=biaotinum%>-ln)+'���ַ�'+(<%=biaotinum/2%>-ln/2)+'����';
      //if (ln>=8) form.biaoti.readOnly=true;  //��������������������޷�������޸�
}

//˵���ֽ������ƣ����������=========
String.prototype.Length=function(){
temp=this.replace(/([^\x00-\xff])/g,"$1$1");
return temp.length
}
//��ʾ�����ַ���

function kn2()
  {
   var ln=myform.memo.value.Length();
     num2.innerText='������������:'+(<%=memonum%>-ln)+'���ַ�'+(<%=memonum/2%>-ln/2)+'����';
      //if (ln>=8) form.memo.readOnly=true;  //��������������������޷�������޸�
}

function myform_onsubmit() {
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

if ((document.myform.biaoti.value.Length()><%=biaotinum%>) || (document.myform.biaoti.value.Length()<4)) //�ֽ������ƣ����������
     {
	 alert("��Ϣ���ⳤ��Ҫ��4��<%=biaotinum%>�ֽ�(<%=biaotinum/2%>����)������ܳ�100�ֽ�,���������룡");
	  document.myform.biaoti.focus()
	  return false
  }
if (document.myform.keywords.value=="" )
	{
	  alert("������Ϣ�ؼ��ʣ�")
	  document.myform.keywords.focus()
	  return false
	 }
if (document.myform.keywords.value.length>100)
	{
	  alert("�ؼ��ʲ��ܳ���100���ַ�")
	  document.myform.keywords.focus()
	  return false
	 }	 
if (document.myform.leixing.value=="" )
	{
	  alert("��ѡ�������ͣ�")
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

if (document.myform.memo.value=="" )
	{
	  alert("������Ϣ˵����")
	  document.myform.memo.focus()
	  return false
	 }
if (document.myform.memo.value.Length()><%=memonum%>)  //�ֽ������ƣ����������
     {
	 alert("��Ϣ˵������Ҫ��<%=memonum%>�ֽ�(<%=memonum/2%>����)�����������룡");
	  document.myform.memo.focus()
	  return false
  }
if (document.myform.name.value=="" )
	{
	  alert("������ϵ�ˣ�")
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
//		alert("���淶�ĵ绰���룡");
//		document.myform.dianhua.focus();
//		return false;
//	}	
//var sMobile = document.myform.CompPhone.value
//if((document.myform.CompPhone.value!="") && (!(/^((\(\d{3}\))|(\d{3}\-))?13[0-9]\d{8}|15[0-9]\d{8}|18[0-9]\d{8}/.test(sMobile)))) //(����13��15��18�Ŷ�)
//	{
//		alert("����������11λ�ֻ��Ż�����ȷ���ֻ���ǰ��λ��");
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
<%if codekg1=1 then%>
    if (document.myform.wenti.value=="" )
	{	  
      alert("��֤�𰸲���Ϊ�գ����������룡");
	  document.myform.wenti.focus()
	  return false
	 }
	
    if(document.myform.wenti.value != document.myform.daan.value) 
	{	  
      alert("��֤�𰸲��ԣ����������룡");
	  document.myform.wenti.focus()
	  return false
	 }
 <%end if%>
 <%if codekg2=1 then%>
    if (document.myform.yzm.value=="" )
	{	  
      alert("������֤�벻��Ϊ�գ����������룡");
	  document.myform.yzm.focus()
	  return false
	 }
<%end if%>
 <%if codekg3=1 then%>
    if (document.myform.code.value=="" )
	{	  
      alert("������֤�벻��Ϊ�գ����������룡");
	  document.myform.code.focus()
	  return false
	 }
<%end if%>
<%if codekg4=1 then%>
    if (document.myform.ttdv.value=="" )
	{	  
      alert("��֤���ڲ���Ϊ�գ����������룡");
	  document.myform.ttdv.focus()
	  return false
	 }
<%end if%>
<%if codekg5=1 then%>
    if (document.myform.validate_codee.value=="" )
	{	  
      alert("��ʽ��֤�벻��Ϊ�գ����������룡");
	  document.myform.validate_codee.focus()
	  return false
	 }
<%end if%>
	  
}
// --></script>
</head>

<body topmargin="0" leftmargin="0">

<div align="center">

  <center>
  <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="980" height="428">
    <tr>
      <td width="215" height="356" valign="top"><div align="center">
      <!--#include file=left.asp--></div></td>
	  <td width="5">&nbsp;</td>
      <td width="279" height="356" valign="top" bgcolor="#FFFFFF"><table width="760" border="0" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF" class="dq1">

        <tr>
          <td width="760" valign="top" bgcolor="#FFFFFF"><table width="100%" border="0" align="right" cellpadding="0" cellspacing="0" bordercolor="#111111" style="border-collapse: collapse">
          <form method="POST" name="myform" id="myform" language="javascript" onSubmit="return myform_onsubmit()" action="addtourists_save.asp?<%=C_ID%>">
             
                <tr>
                  <td colspan="3"><div align="center">
                      <table width="100%" height="100%" border="0" align="center" cellpadding="6" cellspacing="0">
                       
                        <tr bgcolor="#FACA38">
                          <td height="1" colspan="2" ><font color="#FF0000">&nbsp;&nbsp;&nbsp;&nbsp;��������<u><%=title%></u>�з�����Ϣ������������и������ع����йط��ɡ����棬���ص��ºͳ���ԭ�򣻽�ֹ���ñ�վ����Σ�����Һ���թ���˵���Ϊ����ֹ����Υ����Ʒ��α����Ʒ����ٽ�����Ϣ����վ����ɾ��Υ����Ϣ�����йز��žٱ���Ȩ������Ϣ������Υ���涨���ܵ�����׷���ߣ������ɷ������Ը�����Ⱥ��������Ϣ���ɾ����</font></td>
                        </tr>
						<TABLE>
                        <TR>
							<TD><TABLE> 
							<tr>
                          <td height="25" width="100%" colspan="2" align="center"><b>ע���Ա�������ö����ϴ�ͼƬ�������ɫ����֤���ߡ���ϸ˵���������ñ༭����<br>����������Ϣ��Ȩ�ޣ�������Ϣ�Զ����룬�Ժ󷢲�����ݣ�����<a href="<%=reg%>"><font color="#0000FF">ע��</font></a>�󷢲���Ϣ��</b><br>��<font color="#ff0000">*</font>Ϊ����&nbsp; </td>     
                        </tr>
                        <tr>
                          <td height="25" width="60"> ���׵�����</td>
                          <td height="25" width="700">
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
          <SELECT name="city_one" size="1" id="select" onChange="changelocation(document.myform.city_one.options[document.myform.city_one.selectedIndex].value)">
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
          <SELECT name="city_two" id="select6" onChange="changelocation4(document.myform.city_two.options[document.myform.city_two.selectedIndex].value,document.myform.city_one.options[document.myform.city_one.selectedIndex].value)">
            <OPTION value="" selected>ѡ�����</OPTION>
          </SELECT>
          <SELECT name="city_three" id="select7">
            <OPTION value="" selected>ѡ�����</OPTION>
          </SELECT>

<font color="#ff0000"> *</font></td>
                        </tr>
						<tr>
                          <td height="25" width="60"> ��Ϣ���</td>
                          <td height="25" width="450"><%
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
          </SELECT> <font color="#ff0000"> *</font></td>
                        </tr>
                        <tr>
                          <td height="25" width="60"> ��Ϣ���⣺<!---��վ��ʶ��Ⱥ��������Ϣ����Ⱥ��������Ϣ���ɾ����---></td>
                          <td height="25" width="450"><input style="BACKGROUND-COLOR: #eeeeee; BORDER-BOTTOM: #FFFFFF 1px solid; BORDER-LEFT: #a2a2a2 1px solid; BORDER-RIGHT: #ffffff 1px solid; BORDER-TOP: #a2a2a2 1px solid; COLOR: #333333; FONT-SIZE: 9pt; HEIGHT: 19px;" type="text" maxlength="100" name="biaoti" size="60" onpropertychange="kn()" value="">
                            <font color="#ff0000"> *</font><br>��<font color="#CC5200"><%=biaotinum%>�ֽڣ�<%=biaotinum/2%>��������</font>��<span id=num>�㻹��������<%=biaotinum%>���ֽڣ�<%=biaotinum/2%>����</span></td>
                        </tr>
                        <tr>
                          <td height="25" width="60"> �� �� �ʣ�</td>
                          <td height="25" width="450"><input style="BACKGROUND-COLOR: #eeeeee; BORDER-BOTTOM: #FFFFFF 1px solid; BORDER-LEFT: #a2a2a2 1px solid; BORDER-RIGHT: #ffffff 1px solid; BORDER-TOP: #a2a2a2 1px solid; COLOR: #333333; FONT-SIZE: 9pt; HEIGHT: 19px;" type="text" maxlength="100" name="keywords" size="60" value="">
                            <font color="#ff0000"> *</font> <input type="button" name="btn" value="�ӱ��⸴��" title="���Ʊ��⵽�ؼ���" onClick="CopyWebbiaoti(document.myform.biaoti.value);"><br>��<font color="#CC5200">��Ϣ�ؼ��ʣ��á�,���ŷָ�����100�ֽ��ڡ�</font>��</td>
                        </tr>
                        <tr>
                          <td height="25" width="71"> �������ͣ�</td>
                          <td height="25" width="408"><%call leixs()%>  <font color="#ff0000"> *</font></td>
                        </tr>
                        <!--<tr>
                          <td height="25" width="60"> �г��۸�</td>
                          <td height="25" width="450">
                              <input style="BACKGROUND-COLOR: #eeeeee; BORDER-BOTTOM: #FFFFFF 1px solid; BORDER-LEFT: #a2a2a2 1px solid; BORDER-RIGHT: #ffffff 1px solid; BORDER-TOP: #a2a2a2 1px solid; COLOR: #333333; FONT-SIZE: 9pt; HEIGHT: 19px;" type="text" name="scjiage" size="10" maxlength="30" value="" onKeyUp="if(isNaN(value)){alert('�۸�ֻ������������');value='';}"   onbeforepaste="clipboardData.setData('text',clipboardData.getData('text').replace(/[^\d|^\.]/g,''))">
                            Ԫ <font color="#ff0000"> *</font>��0<font color="#CC5200">Ԫ��ʾ����</font>��</td>
                        </tr>-->
                        <tr>
                          <td height="25" width="60"> ���׼۸�</td>
                          <td height="25" width="450">
                              <input style="BACKGROUND-COLOR: #eeeeee; BORDER-BOTTOM: #FFFFFF 1px solid; BORDER-LEFT: #a2a2a2 1px solid; BORDER-RIGHT: #ffffff 1px solid; BORDER-TOP: #a2a2a2 1px solid; COLOR: #333333; FONT-SIZE: 9pt; HEIGHT: 19px;" type="text" name="jiage" size="10" maxlength="30" value="" onKeyUp="if(isNaN(value)){alert('�۸�ֻ������������');value='';}"   onbeforepaste="clipboardData.setData('text',clipboardData.getData('text').replace(/[^\d|^\.]/g,''))">
                            Ԫ <font color="#ff0000"> *</font>��0<font color="#CC5200">Ԫ��ʾ����</font>��</td>
                        </tr>
                        <tr>
                          <td height="25">����������</td>
                          <td height="25"><select name="days" style="BACKGROUND-COLOR: #FFCCFF; BORDER-BOTTOM: #FFFFFF 1px solid; BORDER-LEFT: #a2a2a2 1px solid; BORDER-RIGHT: #ffffff 1px solid; BORDER-TOP: #a2a2a2 1px solid; COLOR: #333333; FONT-SIZE: 9pt; HEIGHT: 19px;">
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
                              </td>
                        </tr>
					    <tr>
                        <td height="25">�Ķ�Ȩ�ޣ�</td>
                        <td width="400" height="25"><input type="radio" name="Readinfo" value="0" checked>�ο�	<input type="radio" name="Readinfo" value="1">��ͨ��Ա	<input type="radio" name="Readinfo" value="2">VIP��Ա	
                             <br>��������Ϣ��ϸ˵�������緽ʽ����Ա�������ͼ��ݣ�</td>
                      </tr>						
					    <!--<tr>
                        <td height="25">�Ķ�Ȩ�ޣ�</td>
                        <td width="400" height="25"><input type="radio" name="Readinfo" value="1" checked>��ͨ��Ա	<input type="radio" name="Readinfo" value="2">VIP��Ա	
                             <br>��������Ϣ��ϸ˵�������緽ʽ����Ա�������ͼ��ݣ�</td>
                      </tr>
					    <tr>
                        <td height="25">�����ɫ��</td>
                        <td height="25" width="700"><input style="BACKGROUND-COLOR: #eeeeee; BORDER-BOTTOM: #FFFFFF 1px solid; BORDER-LEFT: #a2a2a2 1px solid; BORDER-RIGHT: #ffffff 1px solid; BORDER-TOP: #a2a2a2 1px solid; COLOR: #333333; FONT-SIZE: 9pt; HEIGHT: 19px;" type="text" name="wcolor" size="23" maxlength="50" value=""> <font color="#ff0000"> *</font> &lt;=��<FONT color=blue><FONT 
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
                        <td height="25">������ɫ��</td>
                        <td height="25" width="700"><input style="BACKGROUND-COLOR: #eeeeee; BORDER-BOTTOM: #FFFFFF 1px solid; BORDER-LEFT: #a2a2a2 1px solid; BORDER-RIGHT: #ffffff 1px solid; BORDER-TOP: #a2a2a2 1px solid; COLOR: #333333; FONT-SIZE: 9pt; HEIGHT: 19px;" type="text" name="ncolor" size="23" maxlength="50" value=""> <font color="#ff0000"> *</font> &lt;=��<FONT color=blue><FONT 
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
                      </tr>--->
</TABLE></TD>
							
                        </TR>
                        </TABLE>

<TABLE>
<TR>
	<TD><TABLE> 
	<tr>
                          <td height="25" width="60"><p>��ϸ˵����<br><%If Filterhtm>0 Then%>(�οͷ�����Ϣ����Html���룬Ϊȷ�����ݲ��ң��ɽ�����ҳ���Ƶ�����ճ���ڼ��±����ٸ��ƹ���)<%End if%></p></td>
                          <td height="25" width="700">
						  <textarea rows="18" name="memo"  cols="90" onpropertychange="kn2()"><%=request("memo")%></textarea><font color="#ff0000"> *</font><br>��<font color="#CC5200"><%=memonum%>�ֽڣ�<%=memonum/2%>��������</font>��<span id=num2>������������<%=memonum%>���ֽڣ�<%=memonum/2%>����</span><br>��ע��Ϊ��Ա������HTML�༭�����������ã���ʾ�����ۣ�</td>
                        </tr>
                        <tr>
                          <td height="25" width="60"> ��ʵ������</td>
                          <td height="25" width="700"><input style="BACKGROUND-COLOR: #eeeeee; BORDER-BOTTOM: #FFFFFF 1px solid; BORDER-LEFT: #a2a2a2 1px solid; BORDER-RIGHT: #ffffff 1px solid; BORDER-TOP: #a2a2a2 1px solid; COLOR: #333333; FONT-SIZE: 9pt; HEIGHT: 19px;" type="text" name="name" size="23" maxlength="15" value="">
                            &nbsp; <font color="#ff0000"> *</font> </td>
                        </tr>
                         <tr>
                          <td height="25">�ƶ��绰��</td>
                          <td width="408" height="25"><input style="BACKGROUND-COLOR: #eeeeee; BORDER-BOTTOM: #FFFFFF 1px solid; BORDER-LEFT: #a2a2a2 1px solid; BORDER-RIGHT: #ffffff 1px solid; BORDER-TOP: #a2a2a2 1px solid; COLOR: #333333; FONT-SIZE: 9pt; HEIGHT: 19px;" type="text" name="CompPhone" size="23" maxlength="20" value="" onKeyUp="if(isNaN(value)){alert('�ƶ��绰ֻ������������');value='';}" onafterpaste="this.value=this.value.replace(/\D/g,'')">  &nbsp; <font color="#0000ff"> *</font>��<font color="#CC5200">�̶��绰���ƶ��绰������һ</font>��</td>
                        </tr>

                        <tr>
                          <td height="25">�̶��绰��</td>
                          <td width="600" height="25"><input style="BACKGROUND-COLOR: #eeeeee; BORDER-BOTTOM: #FFFFFF 1px solid; BORDER-LEFT: #a2a2a2 1px solid; BORDER-RIGHT: #ffffff 1px solid; BORDER-TOP: #a2a2a2 1px solid; COLOR: #333333; FONT-SIZE: 9pt; HEIGHT: 19px;" type="text" name="dianhua" size="23" maxlength="20" value="">
                            &nbsp; <font color="#0000ff"> *</font>��<font color="#CC5200">�̶��绰���ƶ��绰������һ</font>�� <br>(���Ҵ���-����-�绰-�ֻ���ʽ(��ֻ������+�绰��绰)��+086-010-85858585-2538)</td>
                        </tr>
                        <tr>
                          <td height="25">������룺</td>
                          <td width="600" height="25"><input style="BACKGROUND-COLOR: #eeeeee; BORDER-BOTTOM: #FFFFFF 1px solid; BORDER-LEFT: #a2a2a2 1px solid; BORDER-RIGHT: #ffffff 1px solid; BORDER-TOP: #a2a2a2 1px solid; COLOR: #333333; FONT-SIZE: 9pt; HEIGHT: 19px;" type="text" name="fax" size="23" maxlength="20" value="">
                            &nbsp; (���Ҵ���-����-�绰-�ֻ���ʽ(��ֻ������+�绰��绰)��+086-010-85858585-2538)</td>
                        </tr>
                        <tr>
                          <td height="25" width="71">�������䣺</td>                          
                          <td width="520" height="-5"><input name="email" type="text" style="BACKGROUND-COLOR: #eeeeee; BORDER-BOTTOM: #FFFFFF 1px solid; BORDER-LEFT: #a2a2a2 1px solid; BORDER-RIGHT: #ffffff 1px solid; BORDER-TOP: #a2a2a2 1px solid; COLOR: #333333; FONT-SIZE: 9pt; HEIGHT: 19px;" value="" size="23" maxlength="40" >
                            &nbsp; ��<font color="#FF0000">��Ҫ��</font><font color="#CC5200">���ʼ���ַ��������Ľ�����Ϣ������ȷ��д��</font>��</td>
                        </tr>
                        <tr>
                          <td height="25">QQ �� �룺</td>
                          <td width="408" height="-1"><input name="qq" type="text" style="BACKGROUND-COLOR: #eeeeee; BORDER-BOTTOM: #FFFFFF 1px solid; BORDER-LEFT: #a2a2a2 1px solid; BORDER-RIGHT: #ffffff 1px solid; BORDER-TOP: #a2a2a2 1px solid; COLOR: #333333; FONT-SIZE: 9pt; HEIGHT: 19px;"  size="23" maxlength="20" onKeyUp="if(isNaN(value)){alert('QQ����ֻ������������');value='';}" onafterpaste="this.value=this.value.replace(/\D/g,'')">
                            &nbsp;</td>
                        </tr>
                        <tr>
                          <td height="25"> ΢�ź��룺</td>
                          <td height="25" width="408"><input name="weixin" type="text" style="BACKGROUND-COLOR: #eeeeee; BORDER-BOTTOM: #FFFFFF 1px solid; BORDER-LEFT: #a2a2a2 1px solid; BORDER-RIGHT: #ffffff 1px solid; BORDER-TOP: #a2a2a2 1px solid; COLOR: #333333; FONT-SIZE: 9pt; HEIGHT: 19px;" size="23" maxlength="50" value="">
                            &nbsp; </td><!--123:-->
                        </tr>						
                        <tr>
                          <td height="25"> �������룺</td>
                          <td height="25" width="408"><input name="youbian" type="text" style="BACKGROUND-COLOR: #eeeeee; BORDER-BOTTOM: #FFFFFF 1px solid; BORDER-LEFT: #a2a2a2 1px solid; BORDER-RIGHT: #ffffff 1px solid; BORDER-TOP: #a2a2a2 1px solid; COLOR: #333333; FONT-SIZE: 9pt; HEIGHT: 19px;" size="23" maxlength="30" value="" onKeyUp="if(isNaN(value)){alert('��������ֻ������������');value='';}" onafterpaste="this.value=this.value.replace(/\D/g,'')">
                            &nbsp; </td>
                        </tr>
                        <tr>
                          <td height="25"> ��˾���ƣ�</td>
                          <td height="25" width="408"><input name="mpname" type="text" style="BACKGROUND-COLOR: #eeeeee; BORDER-BOTTOM: #FFFFFF 1px solid; BORDER-LEFT: #a2a2a2 1px solid; BORDER-RIGHT: #ffffff 1px solid; BORDER-TOP: #a2a2a2 1px solid; COLOR: #333333; FONT-SIZE: 9pt; HEIGHT: 19px;" size="50" maxlength="50" value="">
                            &nbsp; </td><!--123:-->
                        </tr>						
                        <tr>
                          <td height="25"> ��ϸ��ַ��</td>
                          <td height="25" width="408"><input name="dizhi" type="text" style="BACKGROUND-COLOR: #eeeeee; BORDER-BOTTOM: #FFFFFF 1px solid; BORDER-LEFT: #a2a2a2 1px solid; BORDER-RIGHT: #ffffff 1px solid; BORDER-TOP: #a2a2a2 1px solid; COLOR: #333333; FONT-SIZE: 9pt; HEIGHT: 19px;" size="50" maxlength="100" value="">
                            &nbsp; </td><!--123:-->
                        </tr>
                        <tr>
                          <td height="25"> ɾ�����룺</td>
                          <td height="25" width="408"><input name="delpas" type="text" style="BACKGROUND-COLOR: #eeeeee; BORDER-BOTTOM: #FFFFFF 1px solid; BORDER-LEFT: #a2a2a2 1px solid; BORDER-RIGHT: #ffffff 1px solid; BORDER-TOP: #a2a2a2 1px solid; COLOR: #333333; FONT-SIZE: 9pt; HEIGHT: 19px;" size="23" maxlength="30" value="">
                            &nbsp;���ο�ɾ��������Ϣר�ã����μǣ�</td>
                        </tr>
<%if codekg1=1 Or codekg2=1 Or codekg3=1 Or codekg4=1 Or codekg5=1 then%>
                       <tr>
<td height="25" colspan="2">=======================<b>Ϊ��ֹȺ��������Ϣ����������д������֤��</b>=======================</td>
                       </tr>
 <%End if%>                 
<%if codekg1=1 then
	  '�ʴ�ʽ��֤
dim rscheck
Randomize 

Set rscheck= Server.CreateObject("ADODB.Recordset")
If SystemDatabaseType = 1 Then
sql="select  * from DNJY_wenda where xs=1 order BY newid()"
Else
sql="select  * from DNJY_wenda where xs=1 order BY rnd(-(ID+"& rnd() &"))"
End if
rscheck.open sql,conn,1,1
%>
	        <tr >
        <td >�ʴ���֤��</td>
        <td>���⣺<%=rscheck("wenti")%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;�𰸣�<input type="text" name="wenti" size="12">  <font color="#ff0000"> *</font></td>
      </tr><input type="hidden" name="daan" value="<%=rscheck("daan")%>">
	<%rscheck.close
	set rscheck=Nothing	
	End If
if codekg2=1  then%>
 <tr >
        <td >������֤��</td>
        <td><input type="text" name="yzm" size="4" /><img src="tt_getcode.asp" alt="��֤��,�������?����ˢ����֤��" height="18" style="cursor : pointer;" onClick="this.src='tt_getcode.asp?t='+(new Date().getTime());" />&nbsp;&nbsp; <font color="#ff0000"> *</font> ����֤��,�������?����ˢ����֤�룩</td>
      </tr>
	  <%End if
	if codekg3=1 then%>
      <tr >
        <td >������֤��</td>
        <td><input type="text" name="code" id="code" size="4" maxlength="4" tabindex="6" onFocus="get_Code();this.onfocus=null;" onKeyUp="dv_ajaxcheck('checke_dvcode','code');"  autocomplete="off"/> <font color="#ff0000"> *</font>
    <span id="imgid" style="color:red">�����ȡ��֤��</span><span id="isok_code"></span></td>
      </tr>
<%End if%>
<%if codekg4=1 then%>

<tr>                                                
  <td height="25"> ��֤���ڣ�</td>
      <td height="25" width="408">������ <select name="ttdv">
<option value="" selected="selected">��ѡ��</option>
<option value="1">����һ</option>
<option value="2">���ڶ�</option>
<option value="3">������</option>
<option value="4">������</option>
<option value="5">������</option>
<option value="6">������</option>
<option value="0">������</option>
</select>
       <font color="#ff0000"> *</font></td>
  </tr>  
<%End if%>
<%if codekg5=1 then%>
<tr>                                                
  <td height="25"> ��ʽ��֤��</td>
      <td height="25" width="408"><img src="tt_getcodee.asp" alt="��֤��,�������?����ˢ����֤��" height="10" style="cursor : pointer;" onClick="this.src='tt_getcodee.asp?t='+(new Date().getTime());" /><input name="validate_codee" type="text" tabindex="3" value="" size="4" maxlength="4" onKeyUp="value=value.replace(/[^\d|^\.]/g,'')"   onbeforepaste="clipboardData.setData('text',clipboardData.getData('text').replace(/[^\d|^\.]/g,''))"/> &nbsp;&nbsp;<font color="#ff0000"> *</font> ����֤��,�������?����ˢ����֤�룩
       </td>
    </tr>
<%End if%>
  <tr bgcolor="#DBDBDB">
                          <td height="1" colspan="2"></td>
                        </tr>
                        <tr>
                          <td height="30" colspan="2"><div align="center" style="color: #FF0000">
                              <div align="left" style="color: #FF0000"><% if xinxish=0 or onoff=0 then %><span style="font-weight: bold">������</span>�οͿ��ٷ�����Ҫ�ȹ���Ա��˲��ܷ�������¡�<%Else%><span style="font-weight: bold">������</span>��ǰϵͳ���ÿ��ٷ�������ȹ���Ա��֤���ֱ�ӷ�����</div></div><%end if%></td>
                        </tr>
                        <tr bgcolor="#D9D9D9">
                          <td height="1" colspan="2"></td>
                        </tr>
                        <tr>
                          <td height="10" colspan="2"></td>
                        </tr>
                        <tr>
                          <td height="25" colspan="2" align="center"><table width="65%"  border="0" cellspacing="0" cellpadding="0">
                              <tr>
                                <td><div align="center">
                                    <input class="inputb" name="I2" type="submit" style="BACKGROUND-COLOR: #eeeeee; BORDER-BOTTOM: #a2a2a2 1px solid; BORDER-LEFT: #ffffff 1px solid; BORDER-RIGHT: #a2a2a2 1px solid; BORDER-TOP: #ffffff 1px solid; COLOR: #333333; FONT-SIZE: 9pt; HEIGHT: 19px;"  value="�ύ����" border="0" /></font>&nbsp;&nbsp;��<font color="#ff0000">���������������Ƿ�׼ȷ</font>�� 
                                </div></td>
                                <td><div align="center">
                                    <input class="inputb" name="I2" type="reset" style="BACKGROUND-COLOR: #eeeeee; BORDER-BOTTOM: #a2a2a2 1px solid; BORDER-LEFT: #ffffff 1px solid; BORDER-RIGHT: #a2a2a2 1px solid; BORDER-TOP: #ffffff 1px solid; COLOR: #333333; FONT-SIZE: 9pt; HEIGHT: 19px;" onClick="javascript:return CheckForm();" value="ȡ������" border="0" />
                                </div></td>
                              </tr>
                          </table></td> </form>
</TABLE></TD>
</TR>
</TABLE>

                        </tr>
                      </table>
                  </div></td>
                </tr>               
             
          </table></td>
        </tr>
      </table>
        <div align="center">
        <center>
        </center>
        </div>        </td>
      <td width="4" valign="top" bgcolor="#E6E6E6"></td>
    </tr>
    <tr>
      <td height="26" colspan="3"><!--#include file=footer.asp--></td>
    </tr>
  </table>
  </center>
</div>

</body>
</html>
