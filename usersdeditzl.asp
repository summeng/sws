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
function textLimitCheck2(thisArea, maxLength){
    if (thisArea.value.length > maxLength){
      alert(maxLength + ' ��������. \r�����Ľ��Զ�ȥ��.');
      thisArea.value = thisArea.value.substring(0, maxLength);
      thisArea.focus();
    }
    /*��дspan��ֵ����ǰ��д���ֵ�����*/    
	messageCount2.innerText = thisArea.value.length;
  }
</script>
</head>

<body background="img/bg1.gif" leftmargin="0" topmargin="0">

<div align="center">
  <center>
  <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="980" height="426">
    <tr>
      <td width="214" height="354" valign="top" bgcolor="#FFFFFF"><div align="center">
      
        <!--#include file=userleft.asp--></div></td>
		<td width="5">&nbsp;</td>
      <td width="760" height="354" valign="top" bgcolor="#FFFFFF"><table width="99%" height="100%" border="0" align="right" cellpadding="0" cellspacing="0" bordercolor="#111111" style="border-collapse: collapse" class="ty1">
        <tr>
          <td width="100%" height="362" align="center" valign="top"><div align="right">
              <table width="100%" height="100%" border="0" align="center" cellpadding="0" cellspacing="0" bordercolor="#111111" style="border-collapse: collapse">
                <form method="POST" name="thisForm" action="usersdeditzlchk.asp?<%=C_ID%>" onSubmit="return CheckForm();">
                  <tr>
                    <td height="25" bgcolor="#FAFAFA"><span class="style5">������޸ĵ���ע�����</span></td>
                    <td bgcolor="#FAFAFA">&nbsp;</td>
                  </tr>
                  <tr>
                    <td height="25" >һ�����ϵ��������ñ�վƽ̨���Զ�����������̻򷢲���ҵ��Ƭ������ͨ���Լ���������̵����ҵ��Ƭ������Ϣ�͹�����Ӫ��Ʒ��<br>����ΪVIP��Ա���õ������Ȩ�ޡ�<a href="upgradevip.asp?<%=C_ID%>"><font color=#0000ff>��Ҫ����ΪVIP</font></a><br>��������Ҫ�ϴ����֤�ղ���ͨ����ˡ�����δ���ϴ�������<a href="certificate.asp?<%=C_ID%>"><font color=#0000ff>�ϴ�</font></a><br>�������루�޸ģ�������˲���Ч��</td>
                    <td bgcolor="#FAFAFA">&nbsp;</td>
                  </tr> 
                  </tr>
                  <tr>
                    <td height="25" ><b>��˽��֪ͨ��</b><%=rs("notification")%></td>
                    <td bgcolor="#FAFAFA">&nbsp;</td>
                  </tr>   				  
                  <tr bgcolor="#CCCCCC">
                    <td width="760" height="1">
 <%
dim m,sdkg,sdfg,mpkg
username=request.cookies("dick")("username")
set rs=conn.execute("select count(id) from [DNJY_info] where yz=1 and username='"&username&"'")
m=rs(0)
rs.close
set rs = Server.CreateObject("ADODB.RecordSet")
sql="select * from [DNJY_user] where username='"&username&"'" 
rs.open sql,conn,1,1
if rs.eof then
response.write "<br>"
response.write "<li>��������"
response.end
end if
sdkg=rs("sdkg")
sdfg=rs("sdfg")
mpkg=rs("mpkg")
mpfg=rs("mpfg")
%>
<SCRIPT language=javascript>
<!--
//�ֽ������ƣ����������=========
String.prototype.Length=function(){
temp=this.replace(/([^\x00-\xff])/g,"$1$1");
return temp.length
}
//��ʾ�����ַ���

function kn()
  {
   var ln=thisForm.sdname.value.Length();
     num.innerText='������������:'+(30-ln)+'���ַ�';
      //if (ln>=8) form.sdname.readOnly=true;  //��������������������޷�������޸�
}

function kn2()
  {
   var ln=thisForm.mpname.value.Length();
     num2.innerText='������������:'+(30-ln)+'���ַ�';
      //if (ln>=8) form.sdname.readOnly=true;  //��������������������޷�������޸�
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
if (document.thisForm.city_one.value==0)
	{
	  alert("��ѡ�������")
	  document.thisForm.city_one.focus()
	  return false
	 }
if (document.thisForm.type_one.value==0)
	{
	  alert("��ѡ����ҵ���࣡")
	  document.thisForm.type_one.focus()
	  return false
	 }
if ((document.thisForm.sdname.value=="") && (document.thisForm.sdkg.value=="2" )) 
     {
	 alert("����������д��");
	  document.thisForm.sdname.focus()
	  return false
  }	
if (document.thisForm.sdname.value.Length()>30)  //�ֽ������ƣ����������
     {
	 alert("��������Ҫ��30�ֽ��ڣ����������룡");
	  document.thisForm.sdname.focus()
	  return false
  }	 
if ((document.thisForm.sdname.value.length>1) && (document.thisForm.sdfg.value=="0" ))
	{
	    alert("��ѡ��������");
		document.thisForm.sdfg.focus();
	    return false;
	}
if ((document.thisForm.sdname.value.length>1) && (document.thisForm.sdjj.value=="" ))
	{
	    alert("������û����д!");
		document.thisForm.sdjj.focus();
	    return false;
	}
if (document.thisForm.mpname.value=="" && (document.thisForm.mpkg.value=="2" )) 
     {
	 alert("��Ƭ���Ʊ�����д��");
	  document.thisForm.mpname.focus()
	  return false
  }	
if (document.thisForm.mpname.value.Length()>30)  //�ֽ������ƣ����������
     {
	 alert("��Ƭ���Ƴ���Ҫ��30�ֽ��ڣ����������룡");
	  document.thisForm.mpname.focus()
	  return false
  }
  
if ((document.thisForm.mpname.value.length>1) && (document.thisForm.mpfg.value=="0" ))
	{
	    alert("��ѡ����Ƭ���");
		document.thisForm.mpfg.focus();
	    return false;
	}
if ((document.thisForm.mpname.value.length>1) && (document.thisForm.mpjj.value=="" ))
	{
	    alert("��Ƭ���û����д!");
		document.thisForm.mpjj.focus();
	    return false;
	}  
if(document.thisForm.name.value.length<1 )
	{
	    alert("������ϵ��!");
		document.thisForm.name.focus();
	    return false;
	}

if(document.thisForm.dianhua.value.length>30 )
	{
	    alert("�̶��绰���ܳ���30���ַ�!");
		document.thisForm.dianhua.focus();
	    return false;
	}
if(document.thisForm.dianhua.value=="" && document.thisForm.CompPhone.value=="") 
	{
	    alert("�̶��绰���ƶ��绰����ͬʱΪ��!");
		document.thisForm.dianhua.focus()
	    return false;
	}
//var sMobile = document.thisForm.dianhua.value
//if((document.thisForm.dianhua.value!="") && (!(/^(0[0-9]{2,3}\-)?([2-9][0-9]{6,7})+(\-[0-9]{1,4})?$/.test(sMobile)))) //(����-�绰-�ֻ���ʽ��010-85858585-2538)
//	{
//		alert("���淶�ĵ绰����");
//		document.thisForm.dianhua.focus();
//		return false;
//	}	
//	var sMobile = document.thisForm.CompPhone.value
//	if((document.thisForm.CompPhone.value!="") && (!(/^((\(\d{3}\))|(\d{3}\-))?13[0-9]\d{8}|15[0-9]\d{8}|18[0-9]\d{8}/.test(sMobile)))) //(����13��15��18�Ŷ�)
//	{
//		alert("����������11λ�ֻ��Ż�����ȷ���ֻ���ǰ��λ");
//		document.thisForm.CompPhone.focus();
//		return false;
//	}
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
if(document.thisForm.email.value.length>30 )
	{
	    alert("���䲻�ܳ���30���ַ�!");
	    document.thisForm.email.focus();
	    return false;
	}
   if((!checkemail(thisForm.email.value))&&(document.thisForm.email.value!=""))
	{
    alert("������Email��ַ����ȷ������������!");
    document.thisForm.email.focus();
    return false;
    }
if(document.thisForm.youbian.value.length>8 )
	{
	    alert("�ʱ಻�ܳ���8���ַ�!");
		document.thisForm.youbian.focus();
	    return false;
	}
if(document.thisForm.dizhi.value.length>100 )
	{
	    alert("��ַ���ܳ���100���ַ�!");
		document.thisForm.dizhi.focus();
	    return false;
	}
if(document.thisForm.http.value.length>100 )
	{
	    alert("��ַ���ܳ���100���ַ�!");
		document.thisForm.http.focus();
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
if(document.thisForm.zhuying.value.length>50 )
	{
	    alert("��Ӫ��Χ���ܳ���50���ַ�!");
		document.thisForm.zhuying.focus();
	    return false;
	}							

}

//-->
                    </SCRIPT></td>
                    <td width="17" height="1"></td>
                  </tr>
                  <tr>
                    <td width="760" valign="top">
					<table width="760" height="105" border="0" align="left" cellpadding="5" cellspacing="0">
				
                        <tr>
                          <td height="10" colspan="3"><%if m<xinxis then
response.write "�ܱ�Ǹ������������Ϣֻ�� "&m&" ����Ҫ���� "&xinxis&" �������м�ֵ�������ͨ������Ϣ����������̻򷢲���ҵ��Ƭ������Ŭ���ޣ�"%>

<%else %></td>
                        </tr>
	                  <tr><td height="30" width="60" > </td>
                     <td height="30"  colspan="3">(��<font color="#ff0000">*</font>����)</td>
                     </tr>
                        <tr>
                          <td height="30"><p align="right">����������</td>
                          <td height="30" width="700"><%Dim rsi
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
set rsi = nothing
%>
    </SELECT>
                           <font color="#ff0000">*</font></td>
                        </tr>
<tr>
                  <td height="25" align="right"> ��ҵ���</td>
                  <td height="25">
<%set rsi=conn.execute("select * from DNJY_hy_type where id>0 and twoid>0 and threeid=0 order by indexid")%>
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
set rsi=conn.execute("select * from DNJY_hy_type where id>0 and twoid>0 and threeid>0 order by indexid")
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
    document.thisForm.type_two.length = 0; 
    document.thisForm.type_two.options[0] = new Option('ѡ����ҵ��������','');
	document.thisForm.type_three.length = 0; 
    document.thisForm.type_three.options[0] = new Option('ѡ����ҵ��������','');
    var locationid=locationid;
	var i;
    for (i=0;i < onecount2; i++)
        {
            if (subcat2[i][1] == locationid)
            { 
                document.thisForm.type_two.options[document.thisForm.type_two.length] = new Option(subcat2[i][0], subcat2[i][2]);
            }        
        }
       
    }    
     function changelocation3(locationid,locationid1)
    {
    document.thisForm.type_three.length = 0; 
    document.thisForm.type_three.options[0] = new Option('ѡ����ҵ��������','');
    var locationid=locationid;
	 var locationid1=locationid1;
    var i;
    for (i=0;i < onecount3; i++)
        {
            if (subcat3[i][2] == locationid)
            { 
			if (subcat3[i][1] == locationid1)
			{
                document.thisForm.type_three.options[document.thisForm.type_three.length] = new Option(subcat3[i][0], subcat3[i][3]);
            }        
        }
       } 
    }            
	                       </SCRIPT>
                                 <SELECT name="type_one" size="1" id="select3" class="inputa" onChange="changelocation2(document.thisForm.type_one.options[document.thisForm.type_one.selectedIndex].value)">
                                   <OPTION value="" selected>ѡ����ҵһ������</OPTION>
                                   <%set rsi=conn.execute("select * from DNJY_hy_type where id>0 and twoid=0 order by indexid asc")
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
                                 <SELECT name="type_two" id="select4" class="inputa" onChange="changelocation3(document.thisForm.type_two.options[document.thisForm.type_two.selectedIndex].value,document.thisForm.type_one.options[document.thisForm.type_one.selectedIndex].value)">
                                   <OPTION value="" selected>ѡ����ҵ��������</OPTION>
                                   <%
	set rsi=conn.execute("select * from DNJY_hy_type where id="&rs("type_oneid")&" and twoid<>0 and threeid=0")
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
                                   <OPTION value="" selected>ѡ����ҵ��������</OPTION>
                                   <%set rsi=conn.execute("select * from DNJY_hy_type where id="&rs("type_oneid")&" and twoid="&rs("type_twoid")&" and threeid<>0")
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
                      <font color="#ff0000">*</font></td>
                </tr>
                        <tr bgcolor="#BBCEFD">
                          <td height="30"><p align="right">�������ƣ�</td>
                          <td height="30" width="700"><input name="sdname" type="text" class="inputa" value="<%=rs("sdname")%>" size="30" maxlength="30" onpropertychange="kn()"> <font color="#ff0000">*</font>&nbsp;&nbsp;<span id=num>�㻹��������30���ַ���15�����֣�</span>
                           </td>
                        </tr>
                        <tr bgcolor="#BBCEFD">
                          <td height="30" align="right">������</td>
                          <td height="30"><select name="sdfg"  class="inputa">
						    <option  value="0" style="background-color:#F8F4F5 ;color: #FF0000">ѡ��������</option>
                            <option  value="1" <%if sdfg=1 then%>selected<%end if%>>�������</option>
                            <option  value="2" <%if sdfg=2 then%>selected<%end if%>>�ƽ𼾽�</option>
							<option  value="3" <%if sdfg=3 then%>selected<%end if%>>�̲�����</option>
							<option  value="4" <%if sdfg=4 then%>selected<%end if%>>����ǧ��</option>
                          </select>
&nbsp;&nbsp; <a href="img/dqmb.jpg" target="_blank">Ч��Ԥ��</a></td>
                        </tr>
                         <tr bgcolor="#BBCEFD">
                          <td height="30"><p align="right">����״̬��</td>
                          <td height="30" width="384"><label>
                            <select name="sdkg" class="inputa">
                              <option value="0" <%if sdkg=0 then%>selected<%end if%>>�ر�״̬</option>
                              <option value="2" <%if sdkg=2 then%>selected<%end if%>>���뿪��</option>
                                                        </select>
                          ��ǰ״̬��
<%if sdkg=1 then
response.write "<font color=""#ff0000""><strong>����״̬</strong></font>"
elseif sdkg=0 then
response.write "<font color=""#0066FF""><strong>�ر�״̬</strong></font>"
elseif sdkg=2 then
response.write "<font color=""#FF9900""><strong>���뿪��</strong></font>"
end if%>
                          </label></td>
                        </tr>
                        <tr bgcolor="#BBCEFD">
                          <td width="120" height="30" valign="top"><p align="right">
                            <p align="right">�����飺</td>
                          <td width="300" height="30" valign="top"><textarea name="sdjj" cols="65" rows="10" onkeyUp="textLimitCheck(this, 800);"><%=rs("sdjj")%></textarea> �� 800 ���ַ� ������ <font color="#CC0000"><span id="messageCount">0</span></font> ����</td>
                        </tr>
                        <tr bgcolor="#E6F0FA">
                          <td height="30"><p align="right">��Ƭ���ƣ�</td>
                          <td height="30" width="700"><input name="mpname" type="text" class="inputa" value="<%=rs("mpname")%>" size="30" maxlength="30" onpropertychange="kn2()"> <font color="#ff0000">*</font>&nbsp;&nbsp;<span id=num2>�㻹��������30���ַ���15�����֣�</span>
                           </td>
                        </tr>
                        <tr bgcolor="#E6F0FA">
                          <td height="30" align="right">��Ƭ���</td>
                          <td height="30"><select name="mpfg"  class="inputa">
						    <option  value="0" style="background-color:#F8F4F5 ;color: #FF0000">ѡ����ҵ��Ƭ���</option>
                            <option  value="1" <%if mpfg=1 then%>selected<%end if%>>�������</option>
                            <option  value="2" <%if mpfg=2 then%>selected<%end if%>>������ʯ</option>
							<option  value="3" <%if mpfg=3 then%>selected<%end if%>>����Ӳ�</option>
							<option  value="4" <%if mpfg=4 then%>selected<%end if%>>�ƽ𼾽�</option>
							<option  value="5" <%if mpfg=5 then%>selected<%end if%>>��ܰ�ۺ�</option>
                          </select>
&nbsp;&nbsp; <a href="img/dqmb.jpg" target="_blank">Ч��Ԥ��</a></td>
                        </tr>
                         <tr bgcolor="#E6F0FA">
                          <td height="30"><p align="right">��Ƭ״̬��</td>
                          <td height="30" width="384"><label>
                            <select name="mpkg" class="inputa">
                              <option value="0" <%if mpkg=0 then%>selected<%end if%>>�ر�״̬</option>
                              <option value="2" <%if mpkg=2 then%>selected<%end if%>>���뿪��</option>
                                                        </select>
                          ��ǰ״̬��
                          <%
if mpkg=1 then
response.write "<font color=""#ff0000""><strong>����״̬</strong></font>"
elseif mpkg=0 then
response.write "<font color=""#0066FF""><strong>�ر�״̬</strong></font>"
elseif mpkg=2 then
response.write "<font color=""#FF9900""><strong>���뿪��</strong></font>"
end if%>
                          </label></td>
                        </tr>
                        <tr bgcolor="#E6F0FA">
                          <td height="30" valign="top"><p align="right">
                            <p align="right">��Ƭ��飺</td>
                          <td width="300" height="30" valign="top"><textarea name="mpjj" cols="65" rows="10" onkeyUp="textLimitCheck2(this, 800);"><%=rs("mpjj")%></textarea> �� 800 ���ַ� ������ <font color="#CC0000"><span id="messageCount2">0</span></font> ����</td>
                        </tr>
                        <tr>
                          <td height="30" valign="top" align="right">�� ϵ �ˣ�</td>
                          <td height="30" valign="top"><input name="name" type="text" class="inputa"value="<%=rs("name")%>" size="20" maxlength="8"> <font color="#ff0000">*</font></td>
                        </tr>

                        <tr>
                          <td height="30" valign="top" align="right">�̶��绰��</td>
                          <td height="30" valign="top"><input name="dianhua" type="text" class="inputa" value="<%=rs("dianhua")%>" size="20" maxlength="30">
                             <font color="#ff0000">*</font> <font color="#CC5200">(����-�绰-�ֻ���ʽ��010-85858585-2538)</font> </td>
                        </tr>
                        <tr>
                          <td height="30" valign="top" align="right">�ƶ��绰��</td>
                          <td height="30" valign="top"><input name="CompPhone" type="text" class="inputa" value="<%=rs("CompPhone")%>" size="20" maxlength="20" onkeyup="this.value=this.value.replace(/\D/g,'')" onafterpaste="this.value=this.value.replace(/\D/g,'')"> <font color="#CC5200">���̶��绰���ƶ��绰������һ��</font></td>
                        </tr>
                        <tr>
                          <td height="30" valign="top" align="right">������룺</td>
                          <td height="30" valign="top"><input name="fax" type="text" class="inputa" value="<%=rs("fax")%>" size="20" maxlength="30">
                            ͬ��<span class="style6">��</span></td>
                        </tr>
                        <tr>
                          <td height="30" valign="top" align="right">�����ʼ���</td>
                          <td height="30" valign="top"><input name="email" type="text" class="inputa"value="<%=rs("email")%>" size="30" maxlength="30"> <font color="#CC5200">��������Ҫ��Ϣ������ȷ��д����</font> </td>
                        </tr>
                        <tr>
                          <td height="30" valign="top" align="right">����QQ�ţ�</td>
                          <td height="30" valign="top"><input name="qq" type="text" class="inputa" value="<%=rs("qq")%>" size="20" maxlength="15" onkeyup="this.value=this.value.replace(/\D/g,'')" onafterpaste="this.value=this.value.replace(/\D/g,'')">
                            ͬ��<span class="style6">��</span></td>
                        </tr>                        
                        <tr>
                          <td height="30" valign="top" align="right">�������룺</td>
                          <td height="30" valign="top"><input name="youbian" type="text" class="inputa" value="<%=rs("youbian")%>" size="10" maxlength="8" onkeyup="this.value=this.value.replace(/\D/g,'')" onafterpaste="this.value=this.value.replace(/\D/g,'')"></td>
                        </tr>
                        <tr>
                          <td height="30" valign="top" align="right">��ϸ��ַ��</td>
                          <td height="30" valign="top"><input name="dizhi" type="text" class="inputa" value="<%=rs("dizhi")%>" size="50" maxlength="100"></td>
                        </tr>
                        <tr>
                          <td height="30" valign="top" align="right">������վ��</td>
                          <td height="30" valign="top"><input name="http" type="text" class="inputa" value="<%=rs("http")%>" size="50" maxlength="100">
                         <br>���������վ������ַ���� <span class="style6">www.ip126.com</span> ǰ�治�ܴ�http:// </td>
                        </tr>
                        <tr>
                          <td height="30" valign="top" align="right">��Ҫ��Ӫ��</td>
                          <td height="30" valign="top"><input name="zhuying" type="text" class="inputa" value="<%=rs("zhuying")%>" size="50" maxlength="50"></td>
                        </tr>                        
       

                        <tr>
                          <td colspan="3"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
                              
                              <tr>
                                <td><div align="center">
                                    <input  name="I4" type="submit" class="inputb" value="�ύ�޸�" border="0" >
                               <span class="style6">���루�޸ģ�������˲���Ч</span> </div> </td>
                                <td><div align="center">
                                    <input  name="I5" type="reset" class="inputb" id="I5" value="ȡ���޸�" border="0" >
                                </div></td>
                              </tr>
                          </table></td>
                        </tr>
   <%end if %>                     
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
