<!--#include file="../conn.asp"-->
<!--#include file="../config.asp"-->
<!--#include file="../SqlIn.Asp"-->
<!--#include file=cookies.asp-->
<%
'=====================================
'���칩����Ϣ��վ����ϵͳ
'��������Ƽ������ҳ�Ʒ
'�ͷ�����http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>
<%Call OpenConn
set rs = Server.CreateObject("ADODB.RecordSet")
sql="select * from [DNJY_user] where username='"&trim(request("username"))&"'" 
rs.open sql,conn,1,1
if rs.eof then
response.write "<font color=#0000FF>û�д˻�Ա�ʺţ�</font>"
Else%>
<SCRIPT language=javascript>
<!--
//��֤���֤��ȷ��
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
    //if (document.thisForm.city_one.value==0)
	//{
	//  alert("��ѡ�������")
	//  document.thisForm.city_one.focus()
	//  return false
	// }
    //if (document.thisForm.type_one.value==0)
	//{
	 // alert("��ѡ����Ϣ���࣡")
	 // document.thisForm.type_one.focus()
	 // return false
	 //}
    //if(document.thisForm.name.value.length<1)
	//{
	//    alert("��ʵ����û����д!");
	//	document.thisForm.name.focus()
	//    return false;
	//}
	if(((document.thisForm.password.value.length<5) || (document.thisForm.password.value.length>12)) &&(document.thisForm.password.value!=""))
        {
            alert("���볤����5��12�ֽ�!");
			document.thisForm.password.focus()
            return false;
        }
    if((!checkeNO(thisForm.idcard.value)) && (document.thisForm.idcard.value!="")){
    alert("���������֤���벻��ȷ!");
    document.thisForm.idcard.focus();
    return false;
    }
	if(document.thisForm.dianhua.value=="" && document.thisForm.CompPhone.value=="") 
	{
	    alert("�̶��绰���ֻ�����ͬʱΪ��!");
		document.thisForm.dianhua.focus()
	    return false;
	}
// var sMobile = document.thisForm.dianhua.value
//if((document.thisForm.dianhua.value!="") && (!(/^([0\+]\d{2,3}-)?(0[0-9]{2,3}\-)?([2-9][0-9]{6,7})+(\-[0-9]{1,4})?$/.test(sMobile)))) //(���Ҵ���-����-�绰-�ֻ���ʽ(��ѡ)��+086-010-85858585-2538)
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
	if((!checkemail(thisForm.email.value))&&(document.thisForm.email.value!=""))
    {
    alert("������Email��ַ����ȷ������������!");
    document.thisForm.email.focus();
    return false;
    }
}

//-->
</SCRIPT>
<meta http-equiv="Content-Language" content="zh-cn">
<title>�޸�����</title>
<link rel="stylesheet" type="text/css" href="../css/css.css">
<BODY>
<body topmargin="0" leftmargin="0">
<div align="center">
  <center>
            <table width="500" border="0" cellspacing="0" cellpadding="0" style="border-collapse: collapse" bordercolor="#111111">
              <form method="POST" name="thisForm" action="user_editdata_save.asp?username=<%=trim(request("username"))%>">
              <tr> 
                <td width="16" background="../images/obj_waku3_03.gif">
                ��</td>
                <td width="500" bgcolor="#EEEEEE">
                  <table width="500" border="0" cellspacing="0" cellpadding="0" height="297">
                    <tr>
                      <td height="22" bgcolor="#EEEEEE" width="489" colspan="6">
                      <p align="center">�޸��û�<%=rs("username")%>������</td>
                      </tr>
                       <tr>
                          <td height="30" bgcolor="#EEEEEE" width="70" align="right">
                      <p align="right">����������</td>
                      <td height="30" bgcolor="#EEEEEE" width="430" colspan="5"><%Dim rsi
set rsi=conn.execute("select * from DNJY_city where id>0 and twoid>0 and threeid=0 order by indexid")
%>
      <SCRIPT language = "JavaScript">
var onecount;
onecount=0;
subcat = new Array();
        <%
		dim count:
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
                           ��Ĭ�ϵ�����</td>
                        <tr><td height="30" bgcolor="#EEEEEE" width="70" align="right">
                      <p align="right">��Ϣ���ࣺ</td>
                          <td height="30" bgcolor="#EEEEEE" width="410" colspan="5"><%Dim rsj
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
                           ��Ĭ�Ϸ��ࣩ</td>
                        </tr>                        </tr>

                    <tr>
                      <td height="30" bgcolor="#EEEEEE" width="70" align="right">
                      <p align="right">��ʵ������</td>
                      <td height="30" bgcolor="#EEEEEE" width="410" colspan="5">
                      &nbsp;<input type="text" maxlength="12" name="name" size="24" value="<%=rs("name")%>"> </td>
                    </tr> 
                    <tr>
                      <td height="30" bgcolor="#EEEEEE" width="70" align="right">
                      <p align="right">��&nbsp;&nbsp;&nbsp;&nbsp;�룺</td>
                      <td height="30" bgcolor="#EEEEEE" width="410" colspan="5">
                      &nbsp;<input type="text" name="password" size="24" value=""> 5-15λ�����޸�����</td>
                    </tr>

                    <tr> 
                      <td height="30" bgcolor="#EEEEEE" width="70" align="right">
                      <p align="right">�� �� ֤��</td>
                      <td height="30" bgcolor="#EEEEEE" width="410" colspan="5">
                      &nbsp;<input type="text" maxlength="20" name="idcard" size="24" value="<%=rs("idcard")%>" ></td>
                    </tr>
                    <tr> 
                      <td height="30" bgcolor="#EEEEEE" width="70" align="right">
                      <p align="right">��&nbsp;&nbsp;&nbsp;&nbsp;��</td>
                      <td height="30" bgcolor="#EEEEEE" width="410" colspan="5">
                      &nbsp;<select id="ctlSex" name="Sex">
                      <option <%if rs("sex")=1 then%>selected<%end if%> value="1">��</option>
                      <option <%if rs("sex")=0 then%>selected<%end if%> value="0">Ů</option>
                      </select></td>
                    </tr>

                    <tr> 
                      <td height="30" bgcolor="#EEEEEE" width="70" align="right">
                      �ֻ����룺</td>
                      <td height="30" bgcolor="#EEEEEE" width="410" colspan="5">
                      &nbsp;<input type="text" maxlength="30" name="CompPhone" size="24" value="<%=rs("CompPhone")%>" > <font color="#0000ff"> *</font>  �ֻ��͹̶��绰������һ</td>
                    </tr>
                    <tr> 
                      <td height="30" bgcolor="#EEEEEE" width="70" align="right">
                      �̶��绰��</td>
                      <td height="30" bgcolor="#EEEEEE" width="410" colspan="5">
                      &nbsp;<input type="text" maxlength="30" name="dianhua" size="24" value="<%=rs("dianhua")%>" > <font color="#0000ff"> *</font><br> (���Ҵ���-����-�绰-�ֻ���ʽ(��ѡ)��+086-010-85858585-2538)</td>
                    </tr>
                    <tr> 
                      <td height="30" bgcolor="#EEEEEE" width="70" align="right">
                      ��&nbsp;&nbsp;&nbsp;&nbsp;�棺</td>
                      <td height="30" bgcolor="#EEEEEE" width="410" colspan="5">
                      &nbsp;<input type="text" maxlength="30" name="fax" size="24" value="<%=rs("fax")%>" ><br> (���Ҵ���-����-�绰-�ֻ���ʽ(��ѡ)��+086-010-85858585-2538)</td>
                    </tr>
                    <tr> 
                      <td height="30" bgcolor="#EEEEEE" width="70" align="right">
                      <p align="right">QQ ���룺</td>
                      <td height="30" bgcolor="#EEEEEE" width="410" colspan="5">
                      &nbsp;<input type="text" maxlength="50" name="qq" size="24" value="<%=rs("qq")%>"></td>
                    </tr>
                    <tr> 
                      <td height="30" bgcolor="#EEEEEE" width="70" align="right">
                      <p align="right">΢�ź��룺</td>
                      <td height="30" bgcolor="#EEEEEE" width="410" colspan="5">
                      &nbsp;<input type="text" maxlength="50" name="weixin" size="24" value="<%=rs("weixin")%>"></td>
                    </tr>
                    <tr> 
                      <td height="30" bgcolor="#EEEEEE" width="70" align="right">
                      <p align="right">�������䣺</td>
                      <td height="30" bgcolor="#EEEEEE" width="410" colspan="5">
                      &nbsp;<input type="text" maxlength="50" name="email" size="24" value="<%=rs("email")%>"></td>
                    </tr>
                    <tr> 
                      <td height="30" bgcolor="#EEEEEE" width="70" align="right">
                      <p align="right">������Ѷ��</td>
                      <td height="30" bgcolor="#EEEEEE" width="410" colspan="5">
                      &nbsp;<%if rs("maillist")=1 then%>               
                <input type="radio" name="maillist" value="1" id="maillist" checked>
                 ����&nbsp;&nbsp;&nbsp;&nbsp;
                <input type="radio" name="maillist" value="0" id="maillist">
                ����</td>
                <%else%>                         
                <input type="radio" name="maillist" value="1" id="maillist">
                 ����&nbsp;&nbsp;&nbsp;&nbsp;
                <input type="radio" name="maillist" value="0" id="maillist" checked>
                ����</td><%end if%></td>
                    </tr>                    
                    
                    <tr> 
                      <td height="30" bgcolor="#EEEEEE" width="70" align="right">
                      �������룺</td>
                      <td height="30" bgcolor="#EEEEEE" width="410" colspan="5">
                      &nbsp;<input type="text" maxlength="6" name="youbian" size="12" value="<%=rs("youbian")%>" > </td>
                    </tr>
                    <tr> 
                      <td height="30" bgcolor="#EEEEEE" width="70" align="right">
                      ��˾���ƣ�</td>
                      <td height="30" bgcolor="#EEEEEE" width="410" colspan="5">
                      &nbsp;<input type="text" maxlength="100" name="mpname" size="48" value="<%=rs("mpname")%>" > Ҳ����ҵ��Ƭ����</td>
                    </tr>
                    <tr>
                      <td height="30" bgcolor="#EEEEEE" width="70" align="right"> ͨ�ŵ�ַ��</td>
                      <td height="30" bgcolor="#EEEEEE" colspan="2">&nbsp;<input type="text" maxlength="100" name="dizhi" size="48" value="<%=rs("dizhi")%>" ></td>
                   <tr> 
                      <td height="30" bgcolor="#EEEEEE" width="70" align="right">
                      <p align="right">���ͨ����</td>
                      <td height="30" bgcolor="#EEEEEE" width="410" colspan="5">
                      &nbsp;<%if rs("useryz")=1 then%>               
                <input type="radio" name="useryz" value="1" id="useryz" checked>
                 ͨ��&nbsp;&nbsp;&nbsp;&nbsp;
                <input type="radio" name="useryz" value="0" id="useryz">
                ��ͨ��</td>
                <%else%>                         
                <input type="radio" name="useryz" value="1" id="useryz">
                 ͨ��&nbsp;&nbsp;&nbsp;&nbsp;
                <input type="radio" name="useryz" value="0" id="useryz" checked>
                ��ͨ��</td><%end if%></td>
                    </tr>
                    <tr> 
                      <td height="25" bgcolor="#EEEEEE" width="70">
                      <p align="center">��</td>

                      <td height="25" bgcolor="#EEEEEE" width="454">
                      <p align="center">
					  <input type="hidden" name="username" value="<%=rs("username")%>">
                      <input border="0" onClick="javascript:return CheckForm();" src="../images/Login_but.gif" name="I4" type="image"></td>
                    </tr>
                    <tr> 
                      <td height="10" bgcolor="#EEEEEE" width="454" colspan="4">��</td>
                    </tr>
                    <tr> 
                      <td height="10" bgcolor="#EEEEEE" width="454" colspan="4">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; ˵�����ʺŲ��ܰ���������ַ���&quot;'%#&amp;&lt;&gt;�������Լ��ո�ȣ�</td>
                    </tr>
                    <tr> 
                      <td height="20" bgcolor="#EEEEEE" width="454" colspan="4">
                      <p align="left">��</td>
                    </tr>
                  </table>
                </td>
                <td width="17" background="../images/obj_waku3_04.gif">
                ��</td>
              </tr>
            </table>
  </center>
</div>
<%End if
Call CloseRs(rs)
Call CloseDB(conn)%>
