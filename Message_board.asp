<!--#include file="conn.asp"-->
<!--#include file="SqlIn.Asp"-->
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
<%if not ChkPost() then 
response.write "<center>��ֹվ���ύ�����"
response.end 
end if
if lmkg2="1" then
call errdick(50)
Call CloseDB(conn)
response.end
end if
dim ts,dick,strUserName,mybook
dim ThisPage,Pagesize,Allrecord,Allpage,tj
mybook=request.querystring("mybook")
username=request.cookies("dick")("username")
strUserName = "��ע���û�"
If request.cookies("dick")("username")<>"" Then
strUserName=request.cookies("dick")("username")
End if
ts=trim(request("ts"))
if request("page")="" then
  ThisPage=1		
else
  ThisPage=request("page")
end if
%>
<link rel="stylesheet" type="text/css" href="/<%=strInstallDir%>css/css.css">
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
<script language="javascript">
<!--
function formcheck(){

	   if (document.form.memo.value=="")
	{
      alert("�������ݲ���Ϊ�գ����������룡")
	  document.form.memo.focus()
	  return false
	 }

 <%if codekg1=1 then%>
    if (document.form.wenti.value=="" )
	{	  
      alert("��֤�𰸲���Ϊ�գ����������룡");
	  document.form.wenti.focus()
	  return false
	 }
	
    if(document.form.wenti.value != document.form.daan.value) 
	{	  
      alert("��֤�𰸲��ԣ����������룡");
	  document.form.wenti.focus()
	  return false
	 }
 <%end if%>
 <%if codekg2=1 then%>
    if (document.form.yzm.value=="" )
	{	  
      alert("������֤�벻��Ϊ�գ����������룡");
	  document.form.yzm.focus()
	  return false
	 }
<%end if%>
 <%if codekg3=1 then%>
    if (document.form.code.value=="" )
	{	  
      alert("������֤�벻��Ϊ�գ����������룡");
	  document.form.code.focus()
	  return false
	 }
<%end if%>
<%if codekg4=1 then%>
    if (document.form.ttdv.value=="" )
	{	  
      alert("��֤���ڲ���Ϊ�գ����������룡");
	  document.form.ttdv.focus()
	  return false
	 }
<%end if%>
<%if codekg5=1 then%>
    if (document.form.validate_codee.value=="" )
	{	  
      alert("��ʽ��֤�벻��Ϊ�գ����������룡");
	  document.form.validate_codee.focus()
	  return false
	 }
<%end if%>	  
}
// --></script>

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
<!---������֤����ý���-->
</head>

<body>
<table width="980" border="0" align="center" cellpadding="0" cellspacing="0">

  <tr>
    <td width="215" height="100%" valign="top"><!--#include file=left.asp--></td>
    <td width="5">&nbsp;</td>
    <td width="760" valign="top" bgcolor="#FFFFFF">
	<table width="760" border="0" align="right" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF" class="ty1">

      <tr>
        <td  bgcolor="#FFFFFF"><table width="760" height="166" border="0" cellpadding="0" cellspacing="0" class="dq2" >
<%set rs=server.createobject("adodb.recordset")
if  strUserName="��ע���û�" then
sql="select * from DNJY_gbook where hf=1 and known=0"&ttccbook&" order by fbsj "&DN_OrderType&""
Else
if mybook="ok" Then
sql="select * from DNJY_gbook where username='"&strUserName&""&ttccbook&"' order by fbsj "&DN_OrderType&""
Else
sql="select * from DNJY_gbook where (hf=1 and known=0) or username='"&strUserName&"'"&ttccbook&" order by fbsj "&DN_OrderType&""
end If
end if
rs.open sql,conn,1,1
if rs.eof Then
Response.Write "<p><br>����"
IF c1>0 Then Response.Write conn.Execute("Select city From DNJY_city Where id="&c1&" and twoid="&c2&" and threeid="&c3&"")(0)
Response.Write "�û�����!"
else
rs.Pagesize=7
Pagesize=rs.Pagesize
Allrecord=rs.Recordcount
Allpage=rs.Pagecount
if ThisPage<1 then                           
ThisPage=1                           
end if
rs.move (ThisPage-1)*Pagesize
%>
<%Dim num1,bg
do while not rs.eof
num1 = Num1+1 '//�ж�������ż������ɫ
If num1 Mod 2=0 Then
bg="F1F9FE"
Else
bg="F9F6FF"
End if			%>

            <tr bgcolor="#<%=bg%>">
              <td width="110" height="22" background="img/22.gif"><p style="margin-top: 0; margin-bottom: 0"> <font color="#FF0000"> <strong> &nbsp;</strong><%If rs("username")<>"" then%><%=rs("username")%><%else%>�ο�<%End if%>�ʣ�</font></td>
              <td width="650" height="26" background="img/22.gif"><font color="#800000"><b><%=rs("gbook1")%></b></font></td>
            </tr>
            <tr bgcolor="#<%=bg%>">
              <td width="110" valign="top" height="19"><p style="margin-top: 0; margin-bottom: 0"><strong><font color="#009900"> &nbsp;</font></strong><font color="#009900">����Ա��</font></td>
              <td height="27"><font color="#808080"><%If rs("hf")=1 then%><%=rs("gbook2")%>(�ظ�ʱ�䣺<font color="#666666"><%=rs("hfsj")%></font>)<%else%><font color="#FF0000">��δ�ظ�</font><%End if%></font></td>
            </tr>

            <%
        tj=tj+1
        if tj>=Pagesize then exit do
        rs.movenext
        loop
        %>

            <tr>
              <td colspan="2" height="6"><table width="100%" border="0" cellpadding="0" cellspacing="0" class="dq2" style="border-collapse: collapse">
                  <tr>
                    <td height="25" width="151"><p align="center"> ����&nbsp;<font color="#CC5200"><%=Allrecord%></font>&nbsp;����¼</td>
                    <td height="25" width="126"><p align="center">�� <font color="#CC5200"><%=Allpage%></font> ҳ</td>
                    <td height="25" width="118"><p align="center">�����ǵ� <font color="#CC5200"><%=ThisPage%></font> ҳ</td>
                    <td height="25" width="187"><p align="center">
                        <%
if ThisPage<2 then     
response.write "<font color=""#808080"">��ҳ</font>&nbsp;"
response.write "<font color=""#808080"">��һҳ</font>&nbsp;"     
else     
response.write "<a href=?page=1&ts="&ts&">��ҳ</a>&nbsp;"
response.write "<a href=?page="&ThisPage-1&"&ts="&ts&">��һҳ</a>&nbsp;"     
end if
if Allpage-ThisPage<1 then     
response.write "<font color=""#808080"">��һҳ</font>&nbsp;"
response.write "<font color=""#808080"">βҳ</font>&nbsp;"  
else     
response.write "<a href=?page="&(ThisPage+1)&"&ts="&ts&">��һҳ</a>&nbsp;"   
response.write "<a href=?page="&Allpage&"&ts="&ts&">βҳ</a>&nbsp;"     
end if
%>
                      </td>
                  </tr>
              </table></td>
            </tr><%end if%>
            <tr>
              <td height="27" colspan="2">&nbsp;</td>
            </tr>
            <tr bgcolor="#FAFAFA">
              <td height="27" colspan="2" align="center"><span class="style1">�����Խ���վ�����û�֮���վ����ѯ����</span></td>
            </tr>
            <tr bgcolor="#CCCCCC">
              <td height="1" colspan="2"></td>
            </tr>
          </table>         

            <table width="760" border="0" align="center" cellPadding="0" cellSpacing="0" bordercolor="#111111" class="font_10_e_blue" style="border-collapse: collapse">
              <tr>
                <td vAlign="top" width="760"><table class="font_10_e_black" cellSpacing="0" cellPadding="3" width="100%" align="center" border="0">
  
<form id="applyform" method="post" name="form" action="Message_board_save.asp?<%=C_ID%>" language="javascript" onsubmit="return formcheck()">
                      <tr>
                        <td width="100"><p align="left">��������</td>
                        <td width="600">
<%set rs=conn.execute("select * from DNJY_city where id>0 and twoid>0 and threeid=0")%>
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
</td>
                      </tr>
                      <tr>
                        <td width="100"><p align="left">��������</td>
                        <td width="600"><select size="1" name="gbook">
                            <option value="0">�� ��ͨ���� ��</option>
                            <option  <%if ts="1" then%> selected <%end if%> value="1">�� Ͷ������ ��</option>
                        </select></td>
                      </tr>
				   <%if  strUserName<>"��ע���û�" then%>	 
                      <tr>
                        <td width="100"><p align="left">�Ƿ񹫿�</td>
                        <td width="600"><input type="radio" name="known" value="0" id="known" checked>
                 ����&nbsp;&nbsp;&nbsp;&nbsp;
                <input type="radio" name="known" value="1" id="known" >
                ������ </td>
                      </tr>					     
                    <%end if%>
                      <tr>
                        <td width="100"><p align="left">��������</td>
                        <td width="600"><font color="#FF0000">��������Ϊ100��</font></td>
                      </tr>
                      <tr>
                        <td><p align="center"> </td>
                        <td><textarea rows="8" name="memo" cols="60" onkeyUp="textLimitCheck(this, 300);"><%=request("memo")%></textarea>&nbsp;&nbsp; <br>�� 300 ���ַ� ������ <font color="#CC0000"><span id="messageCount">0</span></font> ����</td>
                      </tr>
					     </tr>
 <%if codekg1=1 Or codekg2=1 Or codekg3=1 Or codekg4=1 Or codekg5=1 then%>
                       <tr>
<td height="25" colspan="2">=======================<b>Ϊ��ֹȺ��������Ϣ����������д������֤��</b>=======================</td>
                       </tr>
 <%End If
 if codekg1=1 then
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
                        <tr>
                          <td height="25" width="100">�ʴ���֤��</td>
                          <td height="25" width="600">���⣺<%=rscheck("wenti")%>&nbsp;&nbsp;&nbsp;�𰸣�<input type="text" name="wenti" size="12"></td>      <input type="hidden" name="daan" value="<%=rscheck("daan")%>">                    
                        </tr>
	<%rscheck.close
	set rscheck=Nothing
	End If
	if codekg2=1  then%>
                        <tr>
                          <td height="25" width="100">������֤��</td>
                          <td height="25" width="600"> <input type="text" name="yzm" size="4" /><img src="tt_getcode.asp" alt="��֤��,�������?����ˢ����֤��" height="10" style="cursor : pointer;" onClick="this.src='tt_getcode.asp?t='+(new Date().getTime());" />&nbsp;&nbsp; ����֤��,�������?����ˢ����֤�룩</td>                          
                        </tr>
   <%End if
	if codekg3=1 then%>						
                        <tr>
                          <td height="25" width="100">������֤��</td>
                          <td height="25" width="600"><input type="text" name="code" id="code" size="4" maxlength="4" tabindex="6" onfocus="get_Code();this.onfocus=null;" onkeyup="dv_ajaxcheck('checke_dvcode','code');"  autocomplete="off"/>
    <span id="imgid" style="color:red">�����ȡ��֤��</span><span id="isok_code"></span></td>                          
                        </tr> 
   <%End if%>
<%if codekg4=1 then%>					
                        <tr>
                          <td height="25" width="100">������֤��</td>
                          <td height="25" width="600">������ <select name="ttdv">
<option value="" selected="selected">��ѡ��</option>
<option value="1">����һ</option>
<option value="2">���ڶ�</option>
<option value="3">������</option>
<option value="4">������</option>
<option value="5">������</option>
<option value="6">������</option>
<option value="0">������</option>
</select>&nbsp;&nbsp; ����ѡ����ȷ�����ڼ���</td>                          
                        </tr> 
<%End if%>
<%if codekg5=1 then%>					
                        <tr>
                          <td height="25" width="100">��ʽ��֤��</td>
                          <td height="25" width="600"><img src="tt_getcodee.asp" alt="��֤��,�������?����ˢ����֤��" height="10" style="cursor : pointer;" onClick="this.src='tt_getcodee.asp?t='+(new Date().getTime());" /><input name="validate_codee" type="text" tabindex="3" value="" size="4" maxlength="4" onkeyup="value=value.replace(/[^\d|^\.]/g,'')"   onbeforepaste="clipboardData.setData('text',clipboardData.getData('text').replace(/[^\d|^\.]/g,''))"/> &nbsp;&nbsp; ����֤��,�������?����ˢ����֤�룩</td>                          
                        </tr> 
<%End if%>		
 <tr>
                        <td width="50">��</td>
                        <td><table width="100%"  border="0" cellspacing="0" cellpadding="0">
                            <tr>
                              <td height="30"><div align="center">
                                  <input class="inputb" border="0" name="I1" type="submit"   value="ȷ������">
                              </div></td>
                              <td><div align="center">
                                  <input name="I2" type="reset" class="inputb" id="I2" value="ȡ������" border="0">
                              </div></td>
                            </tr>
                        </table></td>
                      </tr>
                    </form>
                  <!---------------------->
                </table></td>
              </tr>
          </table></td>
      </tr>
    </table></td>
    <td width="4" bgcolor="#E6E6E6"></td>
  </tr>
</table>
</body>
</html>
<!--#include file=footer.asp-->

