<!--#include file="conn.asp"-->
<!--#include file="SqlIn.Asp"-->
<!--#include file=top.asp-->
<!--#include file=usercookies.asp-->
<%
'=====================================
'���칩����Ϣ��վ����ϵͳ
'��������Ƽ������ҳ�Ʒ
'�ͷ�����http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>
<link rel="stylesheet" type="text/css" href="/<%=strInstallDir%>css/css.css">
<style type="text/css">
<!--
.style5 {color: #FF0000; font-weight: bold; }
.style6 {color: #FF0000}
-->
</style>
</head>
<body background="img/bg1.gif" leftmargin="0" topmargin="0">
<div align="center">
  <center>
  <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="760" height="426">
    <tr>
      <td width="214" height="354" valign="top" bgcolor="#FFFFFF"><div align="center">
      
        <!--#include file=userleft.asp--></div></td>
		<td width="5">&nbsp;</td>
      <td width="760" height="354" valign="top" bgcolor="#FFFFFF"><table width="99%" height="100%" border="0" align="right" cellpadding="0" cellspacing="0" bordercolor="#111111" style="border-collapse: collapse" class="ty1">
        <tr>
          <td width="99%" height="362" align="center" valign="top"><div align="right">
              <table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0" bordercolor="#111111" style="border-collapse: collapse">
                
                  <tr>
                    <td height="25" bgcolor="#FAFAFA"><span class="style5">���޸��û����ϣ�</span></td>
                    <td bgcolor="#FAFAFA">&nbsp;</td>
                  </tr>
                  <tr bgcolor="#CCCCCC">
                    <td width="456" height="1"><%
dim maillist
username=request.cookies("dick")("username")
set rs = Server.CreateObject("ADODB.RecordSet")
sql="select * from [DNJY_user] where username='"&username&"'" 
rs.open sql,conn,1,1
if rs.eof then
response.write "<br>"
response.write "<li>��������"
response.end
end if
%>

<SCRIPT language=javascript>
<!--
//��ͼ��ע��
function ShowDialog(url, width, height) {
	var arr = showModalDialog(url, window, "dialogWidth:" + width + "px;dialogHeight:" + height + "px;help:no;scroll:no;status:no");
}

function checkeNO(NO){
var str=NO;
//��JavaScript�У�������ʽֻ��ʹ��"/"��ͷ�ͽ���������ʹ��˫����
var Expression=/\d{17}[\d|X]|\d{15}/; 
var objExp=new RegExp(Expression);
if(objExp.test(str)==true){
return true;
}else{
return false;
}
} 


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

function CheckForm()
{
	<%if IsNull(rs("problem")) then%>
        if(document.thisForm.problem.value.length<1)
	{
	    alert("�������ⲻ��Ϊ��!");
		document.thisForm.problem.focus()
	    return false;
	}
        if(document.thisForm.answer1.value.length<1)
	{
	    alert("�𰸲���Ϊ��!");
		document.thisForm.answer1.focus()
	    return false;
	}
	    if(document.thisForm.answer2.value.length<1)
	{
	    alert("ȷ�ϴ𰸲���Ϊ��!");
		document.thisForm.answer2.focus()
	    return false;
	}
	   if(document.thisForm.answer1.value!=document.thisForm.answer2.value)
        {
            alert("��������𰸲�һ��!");
			document.thisForm.answer1.focus()
            return false;
        }
<%end if%>
    if (document.thisForm.city_one.value==0)
	{
	  alert("��ѡ�������")
	  document.thisForm.city_one.focus()
	  return false
	 }
	if(document.thisForm.name.value.length<1)
	{
	    alert("��ʵ��������Ϊ��!");
		document.thisForm.name.focus()
	    return false;
	}


    if((!checkeNO(thisForm.idcard.value)) && (document.thisForm.idcard.value!="")){
    alert("���������֤���벻��ȷ!");
    document.thisForm.idcard.focus();
    return false;
    }
	if(document.thisForm.dianhua.value=="" && document.thisForm.CompPhone.value=="") 
	{
	    alert("�̶��绰���ƶ��绰����ͬʱΪ��!");
		document.thisForm.dianhua.focus()
	    return false;
	}
if(document.thisForm.dianhua.value.length>30 )
	{
	    alert("�̶��绰���ܳ���30���ַ�!");
		document.thisForm.dianhua.focus();
	    return false;
	}
//var sMobile = document.thisForm.dianhua.value
//if((document.thisForm.dianhua.value!="") && (!(/^([0\+]\d{2,3}-)?(0[0-9]{2,3}\-)?([2-9][0-9]{6,7})+(\-[0-9]{1,4})?$/.test(sMobile)))) //(���Ҵ���-����-�绰-�ֻ���ʽ(��ѡ)��+086-010-85858585-2538)
//	{
//		alert("���淶�ĵ绰����");
//		document.thisForm.dianhua.focus();
//		return false;
//	}	
//var sMobile = document.thisForm.CompPhone.value
//if((document.thisForm.CompPhone.value!="") && (!(/^((\(\d{3}\))|(\d{3}\-))?13[0-9]\d{8}|15[0-9]\d{8}|18[0-9]\d{8}/.test(sMobile)))) //(����13��15��18�Ŷ�)
//	{
//		alert("����������11λ�ֻ��Ż�����ȷ���ֻ���ǰ��λ");
//		document.thisForm.CompPhone.focus();
//		return false;
//	}
   if(document.thisForm.email.value.length<1)
	{
	    alert("�����ַ����Ϊ��!");
		document.thisForm.email.focus()
	    return false;
	}   
    if((!checkemail(thisForm.email.value))&&(document.thisForm.email.value!=""))
	{
    alert("������Email��ַ����ȷ������������!");
    document.thisForm.email.focus();
    return false;
    }
   if(document.thisForm.email.value.length>30 )
	{
	    alert("���䲻�ܳ���30���ַ�!");
	    document.thisForm.email.focus();
	    return false;
	}	
//var sMobile = document.thisForm.fax.value
//if((document.thisForm.fax.value!="") && (!(/^([0\+]\d{2,3}-)?(0[0-9]{2,3}\-)?([2-9][0-9]{6,7})+(\-[0-9]{1,4})?$/.test(sMobile)))) //(���Ҵ���-����-�绰-�ֻ���ʽ(��ѡ)��+086-010-85858585-2538)
//	{
//		alert("���淶�Ĵ���绰����");
//		document.thisForm.fax.focus();
//		return false;
//	}
if(document.thisForm.fax.value.length>30 )
	{
	    alert("���治�ܳ���30���ַ�!");
		document.thisForm.fax.focus();
	    return false;
	}
function contain(str,charset)//�ַ����������Ժ���
{
//�����������ִ��в��ܰ���ָ���ַ��е��κ�һ��
//var i;
//for(i=0;i<charset.length;i++)
//if(str.indexOf(charset.charAt(i))>=0)
if(str.indexOf(charset)>=0)//����Ϊ�ִ��в��ܰ���ָ���ַ�������
return true;
return false;
} 
if(contain(document.thisForm.http.value,"http://"))
{ 
alert("��ַǰ��Ҫ��http://");
document.thisForm.http.focus();
return false;
}
if(contain(document.thisForm.http.value,"/"))
{ 
alert("��ַ��Ҫ��/");
document.thisForm.http.focus();
return false;
}     

}
//regexp="^(0[0-9]{2,3}\-)?([2-9][0-9]{6,7})+(\-[0-9]{1,4})?$" 


//-->
                    </SCRIPT></td>
                    <td width="17" height="1"></td>
                  </tr>
                  <tr>
				  <form method="POST" name="thisForm" action="user_data_save.asp?<%=C_ID%>"  language="javascript" onsubmit="return CheckForm()">
                    <td width="760" valign="top"><table width="760" height="126" border="0" align="left" cellpadding="0" cellspacing="0">
                        <tr>
                          <td height="10" colspan="3"></td>
                        </tr>
                       <tr>
                          <td height="30" width="70"><p align="right">ע���ʺţ�</td>
                          <td height="30" width="700">&nbsp;
                              <%=rs("username")%></td>
                        </tr>
				  <%If IsNull(rs("problem")) then%>
                  <tr>
                    <td height="30" width="70"><p align="right">�ܱ����⣺</td>
                    <td  height="30">&nbsp;
                        <input class="inputa"��type="text" name="problem" size="20" maxlength="12"></td>
                  </tr>
                  <tr>
                    <td height="30" width="70"><p align="right">�ܱ��𰸣�</td>
                    <td height="30">&nbsp;
                        <input class="inputa" type="text" name="answer1" size="20"></td>
                  </tr>
                  <tr>
                    <td height="30" width="70"><p align="right">ȷ�ϴ𰸣�</td>
                    <td height="30">&nbsp;
                        <input class="inputa" type="text" name="answer2" size="20"></td>
                  </tr>
				  <%End if%>                        <tr>
                          <td height="30" width="70" ><p align="right">����������</td>
                          <td height="30" width="700">&nbsp;<%Dim rsi
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
        rsi.close
		set rsi=nothing
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
    document.thisForm.city_two.length = 0; 
	document.thisForm.city_two.options[0] = new Option('ѡ�����','');
	document.thisForm.city_three.length = 0; 
	document.thisForm.city_three.options[0] = new Option('ѡ�����','');
    var locationid=locationid;
    var i;
    for (i=0;i < onecount; i++)
        {
            if (subcat[i][1] == locationid)
            { 
                document.thisForm.city_two.options[document.thisForm.city_two.length] = new Option(subcat[i][0], subcat[i][2]);
            }        
        }
        
    }    
	
	function changelocation4(locationid,locationid1)
    {
    document.thisForm.city_three.length = 0; 
    document.thisForm.city_three.options[0] = new Option('ѡ�����','');
    var locationid=locationid;
	 var locationid1=locationid1;
    var i;
    for (i=0;i < onecount4; i++)
        {
            if (subcat4[i][2] == locationid)
            { 
			if (subcat4[i][1] == locationid1)
			{
                document.thisForm.city_three.options[document.thisForm.city_three.length] = new Option(subcat4[i][0], subcat4[i][3]);
            }        
        }
       } 
    }            
              </SCRIPT>
      <SELECT name="city_one" size="1" class="inputa" id="select2" onChange="changelocation(document.thisForm.city_one.options[document.thisForm.city_one.selectedIndex].value)">
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
	  <SELECT name="city_two" id="city_two" class="inputa" onChange="changelocation4(document.thisForm.city_two.options[document.thisForm.city_two.selectedIndex].value,document.thisForm.city_one.options[document.thisForm.city_one.selectedIndex].value)">
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
set rsi = nothing%>
    </SELECT>
                           <font color="#ff0000">*</font> ��Ĭ�ϵ�����</td>
                        </tr>

                        <tr>
                          <td height="30" width="70" ><p align="right">��Ϣ���ࣺ</td>
                          <td height="30" width="700">&nbsp;<%Dim rsj
set rsj=conn.execute("select * from DNJY_type where id>0 and twoid>0 and threeid=0 order by indexid")
%>
      <SCRIPT language = "JavaScript">
var onecount5;
onecount5=0;
subcat5 = new Array();
        <%
		dim count5:
		count5 = 0
        do while not rsj.eof 
        %>
subcat5[<%=count5%>] = new Array("<%=rsj("name")%>","<%=rsj("id")%>","<%=rsj("twoid")%>");
        <%
        count5 = count5 + 1
        rsj.movenext
        loop
        rsj.close
		set rsj=nothing
        %>
onecount5=<%=count5%>;
</SCRIPT>
<%
set rsj=conn.execute("select * from DNJY_type where id>0 and twoid>0 and threeid>0 order by indexid")
%>
 <SCRIPT language = "JavaScript">
var onecount6;
onecount6=0;
subcat6 = new Array();
        <%
		dim count6:count6 = 0
        do while not rsj.eof 
        %>
subcat6[<%=count6%>] = new Array("<%=rsj("name")%>","<%=rsj("id")%>","<%=rsj("twoid")%>","<%=rsj("threeid")%>");
        <%
        count6 = count6 + 1
        rsj.movenext
        loop
        rsj.close
		set rsj = nothing
        %>
onecount6=<%=count6%>;

function changelocation5(locationid)
    {
    document.thisForm.type_two.length = 0; 
	document.thisForm.type_two.options[0] = new Option('ѡ����Ϣ����','');
	document.thisForm.type_three.length = 0; 
	document.thisForm.type_three.options[0] = new Option('ѡ����Ϣ����','');
    var locationid=locationid;
    var i;
    for (i=0;i < onecount5; i++)
        {
            if (subcat5[i][1] == locationid)
            { 
                document.thisForm.type_two.options[document.thisForm.type_two.length] = new Option(subcat5[i][0], subcat5[i][2]);
            }        
        }
        
    }    
	
	function changelocation6(locationid,locationid1)
    {
    document.thisForm.type_three.length = 0; 
    document.thisForm.type_three.options[0] = new Option('ѡ����Ϣ����','');
    var locationid=locationid;
	 var locationid1=locationid1;
    var i;
    for (i=0;i < onecount6; i++)
        {
            if (subcat6[i][2] == locationid)
            { 
			if (subcat6[i][1] == locationid1)
			{
                document.thisForm.type_three.options[document.thisForm.type_three.length] = new Option(subcat6[i][0], subcat6[i][3]);
            }        
        }
       } 
    }            
              </SCRIPT>
      <SELECT name="type_one" size="1" class="inputa" id="select2" onChange="changelocation5(document.thisForm.type_one.options[document.thisForm.type_one.selectedIndex].value)">
            <OPTION value="" selected>ѡ����Ϣ����</OPTION>
  <%set rsj=conn.execute("select * from DNJY_type where id>0 and twoid=0 and threeid=0 order by indexid")
if rsj.eof or rsj.bof then
response.write "<option value=''>û�з���</option>"
else
do until rsj.eof%>
  <OPTION value="<%=rsj("id")%>" <%if rsj("id")=rs("type_oneid") then%>selected<%end if%>><%=rsj("name")%></OPTION>
 <%rsj.movenext
    loop
	%>
	<%end if
rsj.close
set rsj = nothing
%>
      </SELECT> 
	  <SELECT name="type_two" id="type_two" class="inputa" onChange="changelocation6(document.thisForm.type_two.options[document.thisForm.type_two.selectedIndex].value,document.thisForm.type_one.options[document.thisForm.type_one.selectedIndex].value)">
    <OPTION value="" selected>ѡ����Ϣ����</OPTION>
   <%
set rsj=conn.execute("select * from DNJY_type where id="&rs("type_oneid")&" and twoid>0 and threeid=0")
if rsj.eof and rsj.bof then
response.write "<option value=''>û�з���</option>"
else
do until rsj.eof%>
  <OPTION value="<%=rsj("twoid")%>" <%if rsj("twoid")=rs("type_twoid") then%>selected<%end if%>><%=rsj("name")%></OPTION>
 <%rsj.movenext
    loop
	%>
	<%end if
rsj.close
set rsj = nothing
%>
	</SELECT>
	     <SELECT name="type_three" id="type_three" class="inputa">
         <OPTION value="" selected>ѡ����Ϣ����</OPTION>
		<%
set rsj=conn.execute("select * from DNJY_type where id="&rs("type_oneid")&" and twoid="&rs("type_twoid")&" and threeid<>0")
if rsj.eof and rsj.bof then
response.write "<option value=''>û�з���</option>"
else
do until rsj.eof%>
<OPTION value="<%=rsj("threeid")%>" <%if rsj("threeid")=rs("type_threeid") then%>selected<%end if%>><%=rsj("name")%></OPTION>
   <% rsj.movenext
    loop
	%>
<%end if
rsj.close
set rsj = nothing%>
    </SELECT>
                           <font color="#ff0000">*</font> ��Ĭ����Ϣ���ࣩ</td>
                        </tr>

                        <tr>
                          <td height="30" width="70"><p align="right">��ʵ������</td>
                          <td height="30" width="700">&nbsp;
                              <input class="inputa" type="text" maxlength="12" name="name" size="24" value="<%=rs("name")%>"> <font color="#ff0000">*</font></td>
                        </tr>
                        <tr>
                          <td height="30" width="70"><p align="right">�����Ա�</td>
                          <td height="30" width="700">&nbsp;
                              <select class="inputa" id="ctlSex" name="Sex">
                                <option <%if rs("sex")=1 then%>selected<%end if%> value="1">��</option>
                                <option <%if rs("sex")=0 then%>selected<%end if%> value="0">Ů</option>
                            </select></td>
                        </tr>

                        <tr>
                          <td height="30" width="70"><p align="right"> �� �� ֤��</td>
                          <td height="30" width="700">&nbsp;
                              <input class="inputa" type="text" maxlength="20" name="idcard" size="24" value="<%=rs("idcard")%>" > �������⹫����</td>
                        </tr>

                        <tr>
                          <td height="30" width="70" align="right"><div align="right">�ֻ����룺</div> </td>
                          <td height="30" width="700">&nbsp;
                              <input class="inputa" type="text" maxlength="30" name="CompPhone" size="24" value="<%=rs("CompPhone")%>" onkeyup="this.value=this.value.replace(/\D/g,'')" onafterpaste="this.value=this.value.replace(/\D/g,'')"> <font color="#0000ff">*</font> �ֻ��͹̶��绰������һ</td>
                        </tr> 
                        <tr>
                          <td height="30" width="70" align="right"><div align="right">�̶��绰��</div></td>
                          <td height="30" width="700">&nbsp;
                              <input class="inputa" type="text" maxlength="30" name="dianhua" size="24" value="<%=rs("dianhua")%>" > <font color="#0000ff">*</font> (���Ҵ���-����-�绰-�ֻ���ʽ(��ѡ)��+086-010-85858585-2538)</td>
                        </tr>
                        <tr>
                          <td height="30" width="70" align="right"><div align="right">��&nbsp;&nbsp;&nbsp;&nbsp;�棺</div></td>
                          <td height="30" width="700">&nbsp;
                              <input class="inputa" type="text" maxlength="30" name="fax" size="24" value="<%=rs("fax")%>" > (���Ҵ���-����-�绰-�ֻ���ʽ(��ѡ)��+086-010-85858585-2538)</td>
                        </tr>	
                        <tr>
                          <td height="30" width="70"><p align="right">�������䣺</td>
                          <td height="30" width="700">&nbsp;
                              <input class="inputa" type="text" maxlength="50" name="email" size="24" value="<%=rs("email")%>">  ����д��ȷ�ĵ������䣬��������ʼ���<span class="style6">�޷�����</span>��</td>
                        </tr>
                        <tr>
                          <td height="30" width="70" align="right"><div align="right">�� ϵ QQ��</div></td>
                          <td height="30" width="700">&nbsp;
                              <input class="inputa" type="text" maxlength="30" name="qq" size="24" value="<%=rs("qq")%>" onkeyup="this.value=this.value.replace(/\D/g,'')" onafterpaste="this.value=this.value.replace(/\D/g,'')"></td>
                        </tr>
                        <tr>
                          <td height="30" width="70" align="right"><div align="right">΢�ź��룺</div></td>
                          <td height="30" width="700">&nbsp;
                              <input class="inputa" type="text" maxlength="50" name="weixin" size="24" value="<%=rs("weixin")%>" ></td>
                        </tr>
						
                        <tr>
                          <td height="30" width="70" align="right"><div align="right">��˾���ƣ�</div></td>
                          <td height="30" width="700">&nbsp;
                              <input class="inputa" type="text" maxlength="80" name="mpname" size="58" value="<%=rs("mpname")%>" >
                           Ҳ����ҵ��Ƭ����</td>
                        </tr>
                        <tr>
                          <td height="30" width="70" align="right"><div align="right">��˾��ַ��</div></td>
                          <td height="30" width="700">&nbsp;
                              <input class="inputa" type="text" maxlength="100" name="http" size="58" value="<%=rs("http")%>" >
                            �� <span class="style6">www.ip126.com</span> ǰ�治�ܴ�http:// </td>
                        </tr>
                        <tr>
                          <td height="30" width="70" align="right"><div align="right">ͨ�ŵ�ַ��</div></td>
                          <td height="30" width="700">&nbsp;
                              <input class="inputa" type="text" maxlength="50" name="dizhi" size="58" value="<%=rs("dizhi")%>" ></td>
                        </tr>
                        <tr>
                          <td height="30" width="70" align="right"><div align="right">�������룺</div></td>
                          <td height="30" width="700">&nbsp;
                              <input class="inputa" type="text" maxlength="6" name="youbian" size="12" value="<%=rs("youbian")%>" onkeyup="this.value=this.value.replace(/\D/g,'')" onafterpaste="this.value=this.value.replace(/\D/g,'')">
                          </td>
                        </tr>
                        <tr>
                          <td height="30" width="70" align="right"><div align="right">��ͼ��ע��</div></td>
                          <td height="30" width="700">&nbsp;
                              <input class="inputa" name="Emap" type="text" id="Emap" style="width: 150" value="<%=rs("Emap")%>"> <a href=javascript:ShowDialog("user/user_Map.asp?Action=Admin&xy="+document.thisForm.Emap.value+"&Input=Emap",500,350)><font color="#0000FF">��ȡ��γ��</font></a>&nbsp;&nbsp;��ע����VIP��Ա���̹�����������ʾ���ӵ�ͼλ��
                          </td>
                        </tr>
                        <tr>
                          <td height="30" width="70"><p align="right">������Ѷ��</td>
                          <td height="30" width="700">&nbsp;<%if rs("maillist")=1 then%>               
                <input type="radio" name="maillist" value="1" id="maillist" checked>
                 ����&nbsp;&nbsp;&nbsp;&nbsp;
                <input type="radio" name="maillist" value="0" id="maillist">
                ����
                <%else%>                         
                <input type="radio" name="maillist" value="1" id="maillist">
                 ����&nbsp;&nbsp;&nbsp;&nbsp;
                <input type="radio" name="maillist" value="0" id="maillist" checked>
                ����<%end if%>&nbsp;&nbsp;ָ�Ƿ����ʼ����ձ�վ���͵���Ѷ</td>
                        </tr>
                        <tr>
                          <td height="25" colspan="3"><table width="50%"  border="0" cellspacing="0" cellpadding="0">
                              <tr>
                                <td height="10" colspan="2">&nbsp;</td>
                              </tr>
                              <tr>
                                <td height="30"><div align="center">
                                    <input  name="I4" type="submit" class="inputb" value="�ύ�޸�" border="0"  onclick="javascript:return CheckForm();" language="javascript" id="yes">
                                </div> </td>
                                <td><div align="center">
                                    <input  name="I5" type="reset" class="inputb" id="I5" value="ȡ���޸�" border="0" >
                                </div></td>
                              </tr>
                          </table></td>
                        </tr>
                        <tr>
                          <td height="10" colspan="3"><p align="center">��</td>
                        </tr>
                       
                    </table></td>
                    <td width="17"> ��</td>
                  </tr>
                </form>
              </table>
<%Call CloseRs(rs)
%>
          </div></td>
        </tr>
      </table></td>
      <td width="4" valign="top" bgcolor="#e6e6e6"></td>
    </tr>
    <tr>
      <td height="26" colspan="3"><!--#include file=footer.asp--></td>
    </tr>
  </table> 
  </center>
</div>
</body>
</html>
