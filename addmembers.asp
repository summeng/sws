<!--#include file="conn.asp"-->
<!--#include file=top.asp-->
<!--#include file="Include/err.asp"-->
<!--#include file=usercookies.asp-->
<%
'=====================================
'���칩����Ϣ��վ����ϵͳ
'��������Ƽ������ҳ�Ʒ
'�ͷ�����http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>

<%if not ChkPost() then 
response.write "��ֹվ���ύ�����"
response.end 
end if
username=request.cookies("dick")("username")
if lmkg2="1" then
call errdick(50)
response.end
end if
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

'��ֹ��ҳ����
'Response.Buffer = True 
'Response.ExpiresAbsolute = Now() - 1 
'Response.Expires = 0 
'Response.CacheControl = "no-cache"
%>
<%Dim rsuser,ct,msgrs
Call OpenConn
set rsuser=server.createobject("adodb.recordset")
sql = "select * from [DNJY_user] where username='"&username&"'"
rsuser.open sql,conn,1,3
if rsuser.eof then
errdick(9)
response.end
end If
vip=rsuser("vip")
ct=rsuser("c")
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


//��ʾ�����ַ���
function kn2()
  {
   var ln=myform.memo.value.Length();
     num2.innerText='������������:'+(<%=memonum%>-ln)+'���ַ�'+(<%=memonum/2%>-ln/2)+'����';
      //if (ln>=8) form.memo.readOnly=true;  //��������������������޷�������޸�
}

function length(){
messageCount.innerText = document.addform.request.value.length;//����ı��������ַ�����ҳ��ͼ��������
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
if (document.myform.keywords.value.length>40)
	{
	  alert("�ؼ��ʲ��ܳ���40���ַ�")
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
if (document.myform.hfje.value=="" )
	{	  
      alert("���ѽ���Ϊ�գ��������������Ϊ0��");
	  document.myform.hfje.focus()
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
    if((!checkemail(myform.email.value))&&(document.myform.email.value.length>1))
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
<!--kindeditor�༭��-->
	<link rel="stylesheet" href="KindEditor/themes/default/default.css" />
	<link rel="stylesheet" href="KindEditor/plugins/code/prettify.css" />
	<script charset="utf-8" src="KindEditor/kindeditor.js"></script>
	<script charset="utf-8" src="KindEditor/lang/zh_CN.js"></script>
	<script charset="utf-8" src="KindEditor/plugins/code/prettify.js"></script>
	<script>
		KindEditor.ready(function(K) {
			var editor = K.create('textarea[name="memo"]', {
				cssPath : 'KindEditor/plugins/code/prettify.css',
				uploadJson : 'KindEditor/asp/upload_json.asp?action=infopic',
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
</head>

<body topmargin="0" leftmargin="0">


<div align="center">
  <center>
  <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="980">
    
    <tr>
      <td width="214"  valign="top" bgcolor="#FFFFFF">
      <table width="210" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td  height="7" valign="top"><!--#include file=userleft.asp--></td>
        </tr>
      </table>
      <table width="210" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td  height="7"></td>
        </tr>
      </table>
      <table width="210" border="0" cellspacing="0" cellpadding="0" class="ty1">
        <tr>
          <td  height="7"><script src="<%=adjs_path("ads/js","info5",c1&"_"&c2&"_"&c3)%>"></script></td>
        </tr>
      </table>

</td>

<td width="5">&nbsp;</td>
<td width="760" height="356" bgcolor="#FFFFFF"  valign="top" class="ty1"><div align="center"> 
<%
if username<>"" then
Set msgrs = conn.Execute("select count(*) as cmsg from DNJY_Message where isnew='0' and name='"&username&"'")
if cint(msgrs("cmsg"))>0 Or rsuser("mess")=0 then
response.write "<table><td width=""760""><b>�����µ���Ϣ�����Ķ�����ܷ�����Ϣ����<a target=_top href=messs.asp><font color=""#0000ff"">����Ķ�</font></a>��</b></div></td></table>"
response.end
end if
set msgrs=Nothing
end if%>
          <table width="760" border="0" align="right" cellpadding="0" cellspacing="0" bordercolor="#111111" style="border-collapse: collapse">
          <form method="POST" name="myform" language="javascript" onSubmit="return myform_onsubmit()" action="addmembers_save.asp?<%=C_ID%>">               
                <tr> 
                  <td width="760" colspan="2" ><div align="center">
                    <table width="100%" height="100%" border="0" align="right" cellpadding="0" cellspacing="0" >
                     <TABLE><tr bgcolor="#FACA38">
                        <td height="25" colspan="2"><font color="#FF0000">&nbsp;&nbsp;&nbsp;&nbsp;��������<u><%=title%></u>�з�����Ϣ������������и������ع����йط��ɡ����棬���ص��ºͳ���ԭ�򣻽�ֹ���ñ�վ����Σ�����Һ���թ���˵���Ϊ����ֹ����Υ����Ʒ��α����Ʒ����ٽ�����Ϣ����վ����ɾ��Υ����Ϣ�����йز��žٱ���Ȩ������Ϣ������Υ���涨���ܵ�����׷���ߣ������ɷ������Ը�����Ⱥ��������Ϣ���ɾ����</font></td>
                      </tr>
                      <tr bgcolor="#E8E8E8">
                        <td height="1" colspan="2"></td>
                      </tr></TABLE>
<TABLE>
                    <TR>
						<TD><TABLE >
						<tr>
                        <td height="20" colspan="2" align="center"><b>��<font color="#ff0000"> *</font> Ϊ������</b><p></td>
                      </tr>
                        <tr>
                          <td height="25" width="60"> ���׵�����</td>
                          <td height="25" width="700">
<%Dim rsi
set rsi=conn.execute("select * from DNJY_city where id>0 and twoid>0 and threeid=0 order by indexid")
%>
      <SCRIPT language = "JavaScript">
var onecount;
onecount=0;
subcat = new Array();
        <%
		//dim count:
		count = 0
        do while not rsi.eof 
        %>
subcat[<%=count%>] = new Array("<%=rsi("city")%>","<%=rsi("id")%>","<%=rsi("twoid")%>");
        <%
        count = count + 1
        rsi.movenext
        loop
		Call CloseRs(rsi)
        %>
onecount=<%=count%>;
</SCRIPT>
<%
set rsi=conn.execute("select * from DNJY_city where id>0 and twoid>0 and threeid>0 order by indexid")
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
  <%set rsi=conn.execute("select * from DNJY_city where id>0 and twoid=0 and threeid=0 order by indexid")
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
	  <SELECT name="city_two" id="city_two" class="inputa" onChange="changelocation4(document.myform.city_two.options[document.myform.city_two.selectedIndex].value,document.myform.city_one.options[document.myform.city_one.selectedIndex].value)">
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
	     <SELECT name="city_three" id="city_three" class="inputa">
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
    </SELECT>

<font color="#ff0000"> *</font></td>
                        </tr>

						<tr>
                          <td height="25" width="60"> ��Ϣ���</td>
                          <td height="25" width="450"><%set rsi=conn.execute("select * from DNJY_type where id>0 and twoid>0 and threeid=0 order by indexid")%>
                                 <SCRIPT language = "JavaScript">
var onecount2;
onecount2=0;
subcat2 = new Array();
        <%
		dim count2:count2 = 0
        do while not rsi.eof 
        %>
subcat2[<%=count2%>] = new Array("<%=rsi("name")%>","<%=rsi("id")%>","<%=rsi("twoid")%>");
        <%
        count2 = count2 + 1
        rsi.movenext
        loop
        rsi.close
		set rsi = nothing
        %>
onecount2=<%=count2%>;
                           </SCRIPT>
                                 <%
set rsi=conn.execute("select * from DNJY_type where id>0 and twoid>0 and threeid>0 order by indexid")
%>
                                 <SCRIPT language = "JavaScript">
var onecount3;
onecount3=0;
subcat3 = new Array();
        <%
		dim count3:count3 = 0
        do while not rsi.eof 
        %>
subcat3[<%=count3%>] = new Array("<%=rsi("name")%>","<%=rsi("id")%>","<%=rsi("twoid")%>","<%=rsi("threeid")%>");
        <%
        count3 = count3 + 1
        rsi.movenext
        loop
        rsi.close
		set rsi = nothing
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
                                   <%set rsi=conn.execute("select * from DNJY_type where id>0 and twoid=0 order by indexid asc")
if rsi.eof and rsi.bof then
response.write "<option value=''>û�з���</option>"
else
do until rsi.eof%>
                                   <OPTION value="<%=rsi("id")%>" <%if rsi("id")=rs("type_oneid") then%>selected<%end if%>><%=rsi("name")%></OPTION>
                                   <% rsi.movenext
    loop
	end if
rsi.close
set rsi= nothing
%>
                                 </SELECT>
                                 <SELECT name="type_two" id="select4" class="inputa" onChange="changelocation3(document.myform.type_two.options[document.myform.type_two.selectedIndex].value,document.myform.type_one.options[document.myform.type_one.selectedIndex].value)">
                                   <OPTION value="" selected>��Ϣ��������</OPTION>
                                   <%
	set rsi=conn.execute("select * from DNJY_type where id="&rs("type_oneid")&" and twoid<>0 and threeid=0")
if rsi.eof and rsi.bof then
response.write "<option value=''>û�з���</option>"
else
do until rsi.eof%>
                                   <OPTION value="<%=rsi("twoid")%>" <%if rsi("twoid")=rs("type_twoid") then%>selected<%end if%>><%=rsi("name")%></OPTION>
                                   <% rsi.movenext
    loop
	end if
rsi.close
set rsi= nothing%>
                                 </SELECT>
                                 <SELECT name="type_three" id="select5"  class="inputa">
                                   <OPTION value="" selected>��Ϣ��������</OPTION>
                                   <%set rsi=conn.execute("select * from DNJY_type where id="&rs("type_oneid")&" and twoid="&rs("type_twoid")&" and threeid<>0")
if rsi.eof and rsi.bof then
response.write "<option value=''>û�з���</option>"
else
do until rsi.eof%>
                                   <OPTION value="<%=rsi("threeid")%>" <%if rsi("threeid")=rs("type_threeid") then%>selected<%end if%>><%=rsi("name")%></OPTION>
                                   <% rsi.movenext
    loop
	end if
rsi.close
set rsi= Nothing
%>
                                 </SELECT>
          </SELECT> <font color="#ff0000"> *</font></td>
                        </tr>

                      <tr>
                        <td height="25" width="71"> ��Ϣ���⣺<!---��վ��ʶ��Ⱥ��������Ϣ����Ⱥ��������Ϣ���ɾ����---></td>
                        <td height="25" width="700"><input style="BACKGROUND-COLOR: #eeeeee; BORDER-BOTTOM: #FFFFFF 1px solid; BORDER-LEFT: #a2a2a2 1px solid; BORDER-RIGHT: #ffffff 1px solid; BORDER-TOP: #a2a2a2 1px solid; COLOR: #333333; FONT-SIZE: 9pt; HEIGHT: 19px;" type="text" maxlength="100" name="biaoti" size="65" onpropertychange="kn()" value="">
                          <font color="#ff0000"> *</font><br>��<font color="#CC5200"><%=biaotinum%>�ֽڣ�<%=biaotinum/2%>��������</font>��<span id=num>�㻹��������<%=biaotinum%>���ֽڣ�<%=biaotinum/2%>����</span></td>
                      </tr>
                        <tr>
                          <td height="25" width="60"> �� �� �ʣ�</td>
                          <td height="25" width="700"><input style="BACKGROUND-COLOR: #eeeeee; BORDER-BOTTOM: #FFFFFF 1px solid; BORDER-LEFT: #a2a2a2 1px solid; BORDER-RIGHT: #ffffff 1px solid; BORDER-TOP: #a2a2a2 1px solid; COLOR: #333333; FONT-SIZE: 9pt; HEIGHT: 19px;" type="text" maxlength="100" name="keywords" size="65" onpropertychange="kn()" value="">
                            <font color="#ff0000"> *</font> <input type="button" name="btn" value="�ӱ��⸴��" title="���Ʊ��⵽�ؼ���" onClick="CopyWebbiaoti(document.myform.biaoti.value);"><br>��<font color="#CC5200">��Ϣ�ؼ��ʣ��á�,���ŷָ�����100�ֽ��ڡ�</font>��</td>
                        </tr>
                      <tr>
                        <td height="25" width="71"> �������ͣ�</td>
                        <td height="25" width="700"><%call leixs()%>  <font color="#ff0000"> *</font></td>
                      </tr>
                     <!-- <tr>
                        <td height="25" width="71"> �г��۸�</td>
                        <td height="25" width="700">
                            <input style="BACKGROUND-COLOR: #eeeeee; BORDER-BOTTOM: #FFFFFF 1px solid; BORDER-LEFT: #a2a2a2 1px solid; BORDER-RIGHT: #ffffff 1px solid; BORDER-TOP: #a2a2a2 1px solid; COLOR: #333333; FONT-SIZE: 9pt; HEIGHT: 19px;" type="text" name="scjiage" size="10" maxlength="10" value="" onKeyUp="if(isNaN(value)){alert('�۸�ֻ������������');value='';}"   onbeforepaste="clipboardData.setData('text',clipboardData.getData('text').replace(/[^\d|^\.]/g,''))">
                          Ԫ  <font color="#ff0000"> *</font>��0<font color="#CC5200">Ԫ��ʾ����</font>��</td>
                      </tr>-->
                      <tr>
                        <td height="25" width="71"> ���׼۸�</td>
                        <td height="25" width="700">
                            <input style="BACKGROUND-COLOR: #eeeeee; BORDER-BOTTOM: #FFFFFF 1px solid; BORDER-LEFT: #a2a2a2 1px solid; BORDER-RIGHT: #ffffff 1px solid; BORDER-TOP: #a2a2a2 1px solid; COLOR: #333333; FONT-SIZE: 9pt; HEIGHT: 19px;" type="text" name="jiage" size="10" maxlength="10" value="0" onKeyUp="if(isNaN(value)){alert('�۸�ֻ������������');value='';}"   onbeforepaste="clipboardData.setData('text',clipboardData.getData('text').replace(/[^\d|^\.]/g,''))">
                          Ԫ  <font color="#ff0000"> *</font>��0<font color="#CC5200">Ԫ��ʾ����</font>��</td>
                      </tr>
					    <tr>
                        <td height="25">����������</td>
                        <td width="700" height="25"><select name="days" style="BACKGROUND-COLOR: #FFCCFF; BORDER-BOTTOM: #FFFFFF 1px solid; BORDER-LEFT: #a2a2a2 1px solid; BORDER-RIGHT: #ffffff 1px solid; BORDER-TOP: #a2a2a2 1px solid; COLOR: #333333; FONT-SIZE: 9pt; HEIGHT: 19px;">                            
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
                        <td width="700" height="25"><input type="radio" name="Readinfo" value="0" checked>�ο�	<input type="radio" name="Readinfo" value="1">��ͨ��Ա	<input type="radio" name="Readinfo" value="2">VIP��Ա	
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
                       <tr>
                        <td height="25"></td>
                        <td width="700" height="-1"><font color="#FACA38">������</font><font color="#ff0000">������Ϣ�г�����ѡ������۽�����0Ԫϵͳ���Զ��������ʻ��еȶ�ۿ�</font><font color="#FACA38">������</font><br>&nbsp;&nbsp;&nbsp;&nbsp;������Ϣ������Ŀλ����ʾ��ÿ�����Խ��Խ��ǰ������ݷ����������ʵ����۽�</td></tr>
                     <tr>
                        <td height="25">���۽�</td>
                        <td width="700" height="-1"><input style="BACKGROUND-COLOR: #eeeeee; BORDER-BOTTOM: #FFFFFF 1px solid; BORDER-LEFT: #a2a2a2 1px solid; BORDER-RIGHT: #ffffff 1px solid; BORDER-TOP: #a2a2a2 1px solid; COLOR: #333333; FONT-SIZE: 9pt; HEIGHT: 19px;" type="text" name="hfje" size="8" value="0"  onKeyUp="if(isNaN(value)){alert('���۽��ֻ������������');value='';}" onafterpaste="this.value=this.value.replace(/\D/g,'')"> Ԫ  [�����ʻ����Ϊ<%=rsuser("Amount")%>Ԫ] ��ÿ�����=���۽��·���������
				
						</td></tr>


</TABLE></TD>
</TD>
                    </TR>
</TABLE>                    

<TABLE>
<TR>
	<TD><table>                      <tr>
                        <td height="25" width="71"><p>��ϸ˵����<br><%If Filterhtm>2 Then%>(����Html���룬Ϊȷ�����ݲ��ң��ɽ�����ҳ���Ƶ�����ճ���ڼ��±����ٸ��ƹ���)<%End if%></td>
                        <td height="25" width="650"><textarea id="editor" name="memo" style="width:670px;height:400px;"></textarea><input type="checkbox" name="T_SaveImg" value="1" /> ����������Զ��ͼƬ������&nbsp;&nbsp;<input type="checkbox" name="AspJpeg" value="1" checked/>���ϴ�ͼƬ��ˮӡ&nbsp;&nbsp;<input type="checkbox" name="T_conpic" value="1" />��ȡ���ݵ�һ��ͼƬΪ��ͼ</td>
                      </tr></table>
<table>
                      <tr>
                        <td height="25" width="71">��ʵ������</td>
                        <td height="25" width="700"><input style="BACKGROUND-COLOR: #eeeeee; BORDER-BOTTOM: #FFFFFF 1px solid; BORDER-LEFT: #a2a2a2 1px solid; BORDER-RIGHT: #ffffff 1px solid; BORDER-TOP: #a2a2a2 1px solid; COLOR: #333333; FONT-SIZE: 9pt; HEIGHT: 19px;" type="text" name="name" size="23" maxlength="15" value="<%=rsuser("name")%>"> <font color="#ff0000"> *</font></td>
                      </tr>
                      <tr>
                        <td height="25">�ƶ��绰��</td>
                        <td height="25"><input style="BACKGROUND-COLOR: #eeeeee; BORDER-BOTTOM: #FFFFFF 1px solid; BORDER-LEFT: #a2a2a2 1px solid; BORDER-RIGHT: #ffffff 1px solid; BORDER-TOP: #a2a2a2 1px solid; COLOR: #333333; FONT-SIZE: 9pt; HEIGHT: 19px;" type="text" name="CompPhone" size="23" maxlength="40" value="<%=rsuser("CompPhone")%>" onKeyUp="if(isNaN(value)){alert('�ƶ��绰ֻ������������');value='';}" onafterpaste="this.value=this.value.replace(/\D/g,'')">  <font color="#0000ff"> *</font>��<font color="#CC5200">�̶��绰���ƶ��绰������һ</font>��</td>
                      </tr>
                      <tr>
                        <td height="25">�̶��绰��</td>
                        <td height="25"><input style="BACKGROUND-COLOR: #eeeeee; BORDER-BOTTOM: #FFFFFF 1px solid; BORDER-LEFT: #a2a2a2 1px solid; BORDER-RIGHT: #ffffff 1px solid; BORDER-TOP: #a2a2a2 1px solid; COLOR: #333333; FONT-SIZE: 9pt; HEIGHT: 19px;" type="text" name="dianhua" size="23" maxlength="40" value="<%=rsuser("dianhua")%>"> <font color="#0000ff"> *</font> ��<font color="#CC5200">�̶��绰���ƶ��绰������һ</font>��<br>(���Ҵ���-����-�绰-�ֻ���ʽ(��ֻ������+�绰��绰)��+086-010-85858585-2538)</td>
                      </tr>
                        <tr>
                          <td height="25">������룺</td>
                          <td width="600" height="25"><input style="BACKGROUND-COLOR: #eeeeee; BORDER-BOTTOM: #FFFFFF 1px solid; BORDER-LEFT: #a2a2a2 1px solid; BORDER-RIGHT: #ffffff 1px solid; BORDER-TOP: #a2a2a2 1px solid; COLOR: #333333; FONT-SIZE: 9pt; HEIGHT: 19px;" type="text" name="fax" size="23" maxlength="20" value="<%=rsuser("fax")%>">
                            &nbsp; (���Ҵ���-����-�绰-�ֻ���ʽ(��ֻ������+�绰��绰)��+086-010-85858585-2538)</td>
                        </tr>
                      <tr>
                        <td height="25" width="71" >�������䣺</td>
                        <td  height="-5"><input style="BACKGROUND-COLOR: #eeeeee; BORDER-BOTTOM: #FFFFFF 1px solid; BORDER-LEFT: #a2a2a2 1px solid; BORDER-RIGHT: #ffffff 1px solid; BORDER-TOP: #a2a2a2 1px solid; COLOR: #333333; FONT-SIZE: 9pt; HEIGHT: 19px;" type="text" name="email" size="23" maxlength="40" value="<%=rsuser("email")%>">��<font color="#CC5200">��Ҫ�����ʼ���ַ��������Ľ�����Ϣ������ȷ��д��</font>��</td>
                      </tr>					  
                      <tr>
                        <td height="25">QQ �� �룺</td>
                        <td width="700" height="-1"><input style="BACKGROUND-COLOR: #eeeeee; BORDER-BOTTOM: #FFFFFF 1px solid; BORDER-LEFT: #a2a2a2 1px solid; BORDER-RIGHT: #ffffff 1px solid; BORDER-TOP: #a2a2a2 1px solid; COLOR: #333333; FONT-SIZE: 9pt; HEIGHT: 19px;" type="text" name="qq" size="23" maxlength="20" value="<%=rsuser("qq")%>" onKeyUp="if(isNaN(value)){alert('QQ����ֻ������������');value='';}" onafterpaste="this.value=this.value.replace(/\D/g,'')"></td>
                      </tr>
                        <tr>
                          <td height="25"> ΢�ź��룺</td>
                          <td height="25" width="408"><input name="weixin" type="text" style="BACKGROUND-COLOR: #eeeeee; BORDER-BOTTOM: #FFFFFF 1px solid; BORDER-LEFT: #a2a2a2 1px solid; BORDER-RIGHT: #ffffff 1px solid; BORDER-TOP: #a2a2a2 1px solid; COLOR: #333333; FONT-SIZE: 9pt; HEIGHT: 19px;" size="23" maxlength="50" value="<%=rsuser("weixin")%>">
                            &nbsp; </td><!--123:-->
                        </tr>						  
                      <tr>
                        <td height="25">�������룺</td>
                        <td width="700" height="-1"><input style="BACKGROUND-COLOR: #eeeeee; BORDER-BOTTOM: #FFFFFF 1px solid; BORDER-LEFT: #a2a2a2 1px solid; BORDER-RIGHT: #ffffff 1px solid; BORDER-TOP: #a2a2a2 1px solid; COLOR: #333333; FONT-SIZE: 9pt; HEIGHT: 19px;" type="text" name="youbian" size="23" maxlength="6" value="<%=rsuser("youbian")%>" onKeyUp="if(isNaN(value)){alert('��������ֻ������������');value='';}" onafterpaste="this.value=this.value.replace(/\D/g,'')"></td>
                      </tr> 
                        <tr>
                          <td height="25"> ��˾���ƣ�</td>
                          <td height="25" width="408"><input name="mpname" type="text" style="BACKGROUND-COLOR: #eeeeee; BORDER-BOTTOM: #FFFFFF 1px solid; BORDER-LEFT: #a2a2a2 1px solid; BORDER-RIGHT: #ffffff 1px solid; BORDER-TOP: #a2a2a2 1px solid; COLOR: #333333; FONT-SIZE: 9pt; HEIGHT: 19px;" size="50" maxlength="50" value="<%=rsuser("mpname")%>">
                            &nbsp; </td><!--123:-->
                        </tr>
                      <tr>
                        <td height="25" width="71"> ��ϵ��ַ��</td>
                        <td height="25" width="700">
                            <input name="dizhi" type="text" style="BACKGROUND-COLOR: #eeeeee; BORDER-BOTTOM: #FFFFFF 1px solid; BORDER-LEFT: #a2a2a2 1px solid; BORDER-RIGHT: #ffffff 1px solid; BORDER-TOP: #a2a2a2 1px solid; COLOR: #333333; FONT-SIZE: 9pt; HEIGHT: 19px;" size="50" maxlength="100" value="<%=rsuser("dizhi")%>">
                          &nbsp; 
                         </td>
                      </tr>


				
<%if codekg1=1 Or codekg2=1 Or codekg3=1 Or codekg4=1 Or codekg5=1 then%>
                       <tr>
<td height="25" colspan="2">=====================<b>Ϊ��ֹȺ��������Ϣ����������д������֤��</b>======================</td>
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
        <td>���⣺<%=rscheck("wenti")%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;�𰸣�<input type="text" name="wenti" size="12" value=""> <font color="#ff0000"> *</font></td>
      </tr><input type="hidden" name="daan" value="<%=rscheck("daan")%>">
	<%
	Call CloseRs(rscheck)
	End If
	if codekg2=1  then%>
 <tr >
        <td >������֤��</td>
        <td><input type="text" name="yzm" size="4" /><img src="tt_getcode.asp" alt="��֤��,�������?����ˢ����֤��" height="18" style="cursor : pointer;" onClick="this.src='tt_getcode.asp?t='+(new Date().getTime());" />&nbsp;&nbsp; <font color="#ff0000"> *</font> <img src=images/memo.gif alt="��֤��,�������?����ˢ����֤��"></td>
      </tr>
	  <%End if
	if codekg3=1 then%>
      <tr >
        <td >������֤��</td>
        <td><input type="text" name="code" id="code" size="4" maxlength="4" tabindex="6" onfocus="get_Code();this.onfocus=null;" onkeyup="dv_ajaxcheck('checke_dvcode','code');"  autocomplete="off"/> <font color="#ff0000"> *</font>
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
      <td height="25" width="408"><img src="tt_getcodee.asp" alt="��֤��,�������?����ˢ����֤��" height="10" style="cursor : pointer;" onClick="this.src='tt_getcodee.asp?t='+(new Date().getTime());" /><input name="validate_codee" type="text" tabindex="3" value="" size="4" maxlength="4" onkeyup="value=value.replace(/[^\d|^\.]/g,'')"   onbeforepaste="clipboardData.setData('text',clipboardData.getData('text').replace(/[^\d|^\.]/g,''))"/> &nbsp;&nbsp;<font color="#ff0000"> *</font>  <img src=images/memo.gif alt="��֤��,�������?����ˢ����֤��">
       </td>
    </tr>
<%End if%>
  <tr>
    <td height="4" colspan="2"><table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%">
      <tr>
        <td height="10" colspan="4"></td>
      </tr>
      <tr bgcolor="#FFCC00">
        <td height="5" colspan="4"></td>
      </tr>
      <tr bgcolor="#eeeeee">
        <td height="8" colspan="4"></td>
      </tr>
      <tr>
        <td width="100%" height="10" colspan="4"></td>
      </tr>
      <tr>
        <td width="20%" height="25" align="center"><span style="color: #000000">��Ϣ�����ɫ</span></td>
        <td width="20%" height="25" align="center"><span style="color: #000000">��ǰ��Ϣ�ö�</span></td>
        <td width="22%" height="25" align="center"><font color="#FF0000">�ϴ���Ϣ�����ͼƬ<br>���ϴ����滻����������ȡ����ͼ��</font></td>
        <td width="18%" height="25" align="center"><span style="color: #000000">ֱ��ͨ�����</span></td>
      </tr>
      <tr>
        <td width="20%" height="25" align="center"><font face="����">
          <%if rsuser("a")>0 then%>
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
          <%else%>
          <font color="#999999">û�е���</font>
          <%end if%>
        </font></td>
        <td width="20%" height="25" align="center"><font face="����">
          <%
                          dim b
                          b=rsuser("b")
                          %>
          <%if b>0 then%>
          <select name="b" class="inputa">
            <option value="0">��ʹ��</option>
            <%for i=1 to b%>
            <option value="<%=i%>"><%=i%></option>
            <%next%>
          </select>
          <%else%>
          <font color="#999999">û�е���</font>
          <%end if%>
        </font></td>
        <td width="22%" height="25" align="center"><font face="����">
          <%if ct>0 then%>
          <select name="ct" class="inputa">
            <option selected value="0">��ʹ��</option>
            <option value="1">ʹ��</option>
          </select>
          <%else%>
          <font color="#999999">û�е���</font>
          <%end if%>
        </font></td>
        <td width="18%" height="25" align="center"><font face="����">
        <%if onoff>0 then%>

        <%if vip>0 or vipno>0 then%>
          <%if vip>0 then%>
          <font color="#ff0000">VIP�����</font>
          <%else%>
          <%if rsuser("d")>0 then%>
          <select name="d" class="inputa">
            <option selected value="0">��ʹ��</option>
            <option value="1">ʹ��</option>
          </select>
          <%else%>
          <font color="#999999">û�е���</font>
          <%end if%>
          <%end if%>

          <%else%>
          <font color="#999999">��VIP�����</font>
           <%end if%>
            <%else%>
          <font color="#ff0000">�����Ա��˲�����ʾ������¡�</font> 
           <%end if%>
        </font></td>
      </tr>
    </table></td>
  <tr>
    <td height="10" colspan="2"></td>
  </tr>
  <tr>
    <td colspan="2" align="center"><table width="50%"  border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td height="30"><div align="center">
              ��<font color="#ff0000">��������������׼ȷ��</font>��<input name="I1" type="submit" class="inputb" id="I1" style="BACKGROUND-COLOR: #eeeeee; BORDER-BOTTOM: #a2a2a2 1px solid; BORDER-LEFT: #ffffff 1px solid; BORDER-RIGHT: #a2a2a2 1px solid; BORDER-TOP: #ffffff 1px solid; COLOR: #333333; FONT-SIZE: 9pt; HEIGHT: 19px;"  value="�ύ����" border="0" >&nbsp;&nbsp;
        </div></td>
        <td><div align="center">
          <input class="inputb" name="I2" type="reset" style="BACKGROUND-COLOR: #eeeeee; BORDER-BOTTOM: #a2a2a2 1px solid; BORDER-LEFT: #ffffff 1px solid; BORDER-RIGHT: #a2a2a2 1px solid; BORDER-TOP: #ffffff 1px solid; COLOR: #333333; FONT-SIZE: 9pt; HEIGHT: 19px;"  value="ȡ������" border="0">
        </div></td>
      </tr>
    </table></form></table></TD>  

</TR>
</TABLE>

                    </table>
                  </div>                  </td>
                  
                </tr>
              
              </table>
        </center>
        </div>        </td>
      </tr>
    <tr><%
Call CloseRs(rsuser)
Call CloseRs(rs)
Call CloseDB(conn)%>
      <td height="26" colspan="2"><!--#include file=footer.asp--></td>
    </tr>
  </table>
  </center>
</div>

</body>
</html>
