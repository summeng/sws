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
dim username,id
id=trim(request("id"))
if not isnumeric(id) or id="" then
response.write "<li>��������"
response.end
end if
set rs = Server.CreateObject("ADODB.RecordSet")
sql="select username,sdname,sdjj,sdfg,sdkg,mpname,mpjj,mpfg,mpkg,dpdata,http,zhuying,city_oneid,city_twoid,city_threeid,type_oneid,type_twoid,type_threeid,sdgd,mpgd,notification from [DNJY_user] where id="&cstr(id) 
rs.open sql,conn,1,1
%>
<meta http-equiv="Content-Language" content="zh-cn">
<title>�޸��û�������Ƭ����</title>
<link rel="stylesheet" type="text/css" href="../css/css.css">
<BODY>
<body topmargin="0" leftmargin="0">
<SCRIPT language=javascript>
<!--
//�ֽ������ƣ����������=========
String.prototype.Length=function(){
temp=this.replace(/([^\x00-\xff])/g,"$1$1");
return temp.length
}
function CheckForm()
{

     if (document.thisForm.sdname.value.Length()>30) //�ֽ������ƣ����������
     {
	 alert("�������Ƴ���Ҫ��30�ֽ��ڣ����������룡");
	  document.thisForm.sdname.focus()
	  return false
  }	 
if ((document.thisForm.sdname.value.length>1) && (document.thisForm.sdjj.value==""))
	{
	    alert("����������!");
		document.thisForm.sdjj.focus();
	    return false;
	}
	if ((document.thisForm.sdname.value.length>1) && (document.thisForm.sdfg.value=="0" ))
	{
	    alert("��ѡ��������");
		document.thisForm.sdfg.focus();
	    return false;
	}
     if (document.thisForm.mpname.value.Length()>30) //�ֽ������ƣ����������
     {
	 alert("��Ƭ���Ƴ���Ҫ��30�ֽ��ڣ����������룡");
	  document.thisForm.mpname.focus()
	  return false
  }	 
if ((document.thisForm.mpname.value.length>1) && (document.thisForm.mpjj.value==""))
	{
	    alert("������Ƭ���!");
		document.thisForm.mpjj.focus();
	    return false;
	}
if ((document.thisForm.mpname.value.length>1) && (document.thisForm.mpfg.value=="0" ))
	{
	    alert("��ѡ����Ƭ���");
		document.thisForm.mpfg.focus();
	    return false;
    }
if (document.thisForm.type_one.value=="")
	{
	  alert("��ѡ����ҵ���࣡")
	  document.thisForm.type_one.focus()
	  return false
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
if (document.thisForm.notification.value.length>200)
	{
	  alert("���֪ͨ���ݲ��ܳ���200���ַ�")
	  document.thisForm.notification.focus()
	  return false
	 }	
}

//-->
     </SCRIPT>
<div align="center">
  <center><form method="POST" name="thisForm" action="user_editdp_save.asp?id=<%=id%>" onSubmit="return CheckForm();">
            <table width="489" border="0" cellspacing="0" cellpadding="0" style="border-collapse: collapse" bordercolor="#111111">
              <tr> 
                <td width="16" background="../images/obj_waku3_03.gif">
                ��</td>
                <td width="456" bgcolor="#EEEEEE">
                  <table width="454" border="0" cellspacing="0" cellpadding="0" height="177">
                    <tr>
                      <td height="22" bgcolor="#EEEEEE" colspan="4">
                      <p align="center"><b>�޸Ļ�Ա���̺���Ƭ����</b></td>
                    </tr>
                    <tr>
                      <td height="30" bgcolor="#EEEEEE" colspan="4"><HR></td>
                    </tr>
                    <tr>
                      <td height="30" bgcolor="#E3EBF9" colspan="2">
                      <p align="right">�������ƣ�</td>
                      <td height="30" bgcolor="#E3EBF9" colspan="2">
                      &nbsp;
                      <input name="sdname" type="text" id="sdname" value="<%=rs("sdname")%>" size="30" maxlength="17"></td>
                    </tr>

 
                    <tr>
                      <td height="30" bgcolor="#E3EBF9" colspan="2" align="right">������</td>
                      <td height="30" bgcolor="#E3EBF9" colspan="2">&nbsp;
                          <select name="sdfg">
						   <option  value="0" style="background-color:#F8F4F5 ;color: #FF0000">ѡ��������</option>
                            <option  value="1" <%if rs("sdfg")=1 then%>selected<%end if%>>�������</option>
                            <option  value="2" <%if rs("sdfg")=2 then%>selected<%end if%>>�ƽ𼾽�</option>
							<option  value="3" <%if rs("sdfg")=3 then%>selected<%end if%>>�̲�����</option>
							<option  value="4" <%if rs("sdfg")=4 then%>selected<%end if%>>����ǧ��</option>	
                          </select>
 &nbsp;&nbsp; <a href="../img/dqmb.jpg" target="_blank">Ч��Ԥ����</a></td>
                    </tr>
                    <tr> 
                      <td height="30" bgcolor="#E3EBF9" colspan="2">
                      <p align="right">�����飺</td>
                      <td height="30" bgcolor="#E3EBF9" colspan="2">
                      &nbsp;
                      <label>
                      <textarea name="sdjj" cols="50" rows="8" id="sdjj"><%=rs("sdjj")%></textarea>
                      </label></td>
                    </tr>
<tr>
                      <td height="30" bgcolor="#E3EBF9" colspan="2" align="right">������ˣ�</td>
                      <td height="30" bgcolor="#E3EBF9" colspan="2">&nbsp;
                          <select name="sdkg">
                            <option value="0" <%if rs("sdkg")=0 then%>selected<%end if%>>�ر�״̬</option>
                            <option value="1" <%if rs("sdkg")=1 then%>selected<%end if%>>����״̬</option>
			    <option value="2" <%if rs("sdkg")=2 then%>selected<%end if%>>���뿪��</option>
                                                    </select>
��ǰ״̬<%
if rs("sdkg")=1 then
response.write "<font color=""#ff0000""><strong>����״̬</strong></font>"
elseif rs("sdkg")=0 then
response.write "<font color=""#0066FF""><strong>�ر�״̬</strong></font>"
elseif rs("sdkg")=2 then
response.write "<font color=""#FF9900""><strong>���뿪��</strong></font>"
end if%></td>
                    </tr>
<tr>
                      <td height="30" bgcolor="#E3EBF9" colspan="2" align="right">����̶���</td>
                      <td height="30" bgcolor="#E3EBF9" colspan="2">&nbsp;
                          <select name="sdgd">
                            <option value="0" <%if rs("sdgd")=0 then%>selected<%end if%>>δ�̶�</option>
                            <option value="1" <%if rs("sdgd")=1 then%>selected<%end if%>>�ѹ̶�</option>
							</select>
</td>
                    </tr>
                    <tr>
                      <td height="30" bgcolor="#EEEEEE" colspan="4"><HR></td>
                    </tr>					
                    <tr>
                      <td height="30" bgcolor="#CCCCFF" colspan="2">
                      <p align="right">��Ƭ���ƣ�</td>
                      <td height="30" bgcolor="#CCCCFF" colspan="2">
                      &nbsp;
                      <input name="mpname" type="text" id="mpname" value="<%=rs("mpname")%>" size="30" maxlength="17"></td>
                    </tr>
 
                    <tr>
                      <td height="30" bgcolor="#CCCCFF" colspan="2" align="right">��Ƭ���</td>
                      <td height="30" bgcolor="#CCCCFF" colspan="2">&nbsp;
                          <select name="mpfg">
						   <option  value="0" style="background-color:#F8F4F5 ;color: #FF0000">ѡ����Ƭ���</option>
                            <option  value="1" <%if rs("mpfg")=1 then%>selected<%end if%>>�������</option>
                            <option  value="2" <%if rs("mpfg")=2 then%>selected<%end if%>>������ʯ</option>
							<option  value="3" <%if rs("mpfg")=3 then%>selected<%end if%>>����Ӳ�</option>
							<option  value="4" <%if rs("mpfg")=4 then%>selected<%end if%>>�ƽ𼾽�</option>
							<option  value="5" <%if rs("mpfg")=5 then%>selected<%end if%>>��ܰ�ۺ�</option>
                          </select>
 &nbsp;&nbsp; <a href="../img/dqmb.jpg" target="_blank">Ч��Ԥ����</a></td>
                    </tr> 
                    <tr> 
                      <td height="30" bgcolor="#CCCCFF" colspan="2">
                      <p align="right">��Ƭ��飺</td>
                      <td height="30" bgcolor="#CCCCFF" colspan="2">
                      &nbsp;
                      <label>
                      <textarea name="mpjj" cols="50" rows="8" id="mpjj"><%=rs("mpjj")%></textarea>
                      </label></td>
                    </tr>
<tr>
                      <td height="30" bgcolor="#CCCCFF" colspan="2" align="right">��Ƭ��ˣ�</td>
                      <td height="30" bgcolor="#CCCCFF" colspan="2">&nbsp;
                          <select name="mpkg">
                            <option value="0" <%if rs("mpkg")=0 then%>selected<%end if%>>�ر�״̬</option>
                            <option value="1" <%if rs("mpkg")=1 then%>selected<%end if%>>����״̬</option>
			                <option value="2" <%if rs("mpkg")=2 then%>selected<%end if%>>���뿪��</option>
                                                    </select>
��ǰ״̬<%
if rs("mpkg")=1 then
response.write "<font color=""#ff0000""><strong>����״̬</strong></font>"
elseif rs("mpkg")=0 then
response.write "<font color=""#0066FF""><strong>�ر�״̬</strong></font>"
elseif rs("mpkg")=2 then
response.write "<font color=""#FF9900""><strong>���뿪��</strong></font>"
end if%></td>
                    </tr>
<tr>
                      <td height="30" bgcolor="#CCCCFF" colspan="2" align="right">��Ƭ�̶���</td>
                      <td height="30" bgcolor="#CCCCFF" colspan="2">&nbsp;
                          <select name="mpgd">
                            <option value="0" <%if rs("mpgd")=0 then%>selected<%end if%>>δ�̶�</option>
                            <option value="1" <%if rs("mpgd")=1 then%>selected<%end if%>>�ѹ̶�</option>
							</select>
</td>
                    </tr>
                    <tr>
                      <td height="30" bgcolor="#EEEEEE" colspan="4"><HR></td>
                    </tr>
<tr>
                  <td height="30" bgcolor="#EEEEEE" colspan="2"><p align="right">������ҵ���</td>
                  <td height="30" bgcolor="#EEEEEE" colspan="2">
<%Dim rsi
set rsi=conn.execute("select * from DNJY_hy_type where id>0 and twoid>0 and threeid=0 order by indexid")%>
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
                    <tr> 
                      <td height="30" bgcolor="#EEEEEE" colspan="2">
                      <p align="right">�� &nbsp;Ӫ��</td>
                      <td height="30" bgcolor="#EEEEEE" colspan="2">
                      &nbsp;
                      <label>
                      <input name="zhuying" type="text" class="inputa" value="<%=rs("zhuying")%>" size="50" maxlength="50">
                      </label></td>
                    </tr>

                    <tr>
                      <td height="30" bgcolor="#EEEEEE" colspan="2">
                      <p align="right">��˾��վ��</td>
                      <td height="30" bgcolor="#EEEEEE" colspan="2">
                      &nbsp;
                      <input name="http" type="text" class="inputa" id="http" value="<%=rs("http")%>" size="30" maxlength="50"> ǰ�治�ܴ�http://</td>
                    </tr>            					
                     <tr>
                      <td height="30" bgcolor="#EEEEEE" colspan="4"><HR></td>
                    </tr>                   
                      <tr> 
                      <td height="30" bgcolor="#ffffff" colspan="2">
                      <p align="right">�������֪ͨ��</td>
                      <td height="30" bgcolor="#ffffff" colspan="2">
                      &nbsp;
                      <label>
                      <textarea name="notification" cols="50" rows="5" id="notification"><%=rs("notification")%></textarea>
                      </label>(��200�ַ�)</td>
                    </tr>                  
                    <tr>
                      <td height="30" bgcolor="#ffffff" colspan="2">
                      <p align="right">�Ƿ�������ʱ�䣺</td>
                      <td height="30" bgcolor="#ffffff" colspan="2">
                      &nbsp;
                      <input type="radio" name="gxsj" value="1" id="switch" >
                 ����&nbsp;&nbsp;&nbsp;&nbsp;
                <input type="radio" name="gxsj" value="0" id="switch" checked>������<br></td>
                    </tr>                        
                    <tr>
                      <td height="30" bgcolor="#ffffff" colspan="2">
                      <p align="right">���ʱ�䣺</td>
                      <td height="30" bgcolor="#ffffff" colspan="2">
                      &nbsp;
                      <%=rs("dpdata")%></td>
                    </tr>
					<input type="hidden" name="username" value="<%=rs("username")%>">
                    <tr> 
                      <td height="25" bgcolor="#ffffff" width="110">
                      <p align="center">��</td>
                      <td height="25" bgcolor="#ffffff" colspan="2">
                      <p align="right">
                      ��</td>
                      <td height="25" bgcolor="#ffffff" width="280">
                      <p align="center">
                      <input border="0" src="../images/Login_but.gif" name="I4" type="image">
                &nbsp;&nbsp;&nbsp;&nbsp;<a href="user_certificate.asp?id=<%=cstr(id)%>" target=blank><font color=#0000ff>�鿴���ϴ���֤��</font></a></td>
                    </tr>
                    <tr> 
                      <td height="10" bgcolor="#ffffff" colspan="4">&nbsp;</td>
                    </tr>
                  </table>
                </td>
                <td width="17" background="../images/obj_waku3_04.gif"></td>
              </tr>
            </table></form>
  </center>
</div>
<%Call CloseRs(rs)
Call CloseDB(conn)%>
