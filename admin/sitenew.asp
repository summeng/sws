<!--#include file="../conn.asp"-->
<!--#include file="../SqlIn.Asp"-->
<!--#include file=../config.asp-->
<!--#include file=cookies.asp-->
<%
'=====================================
'���칩����Ϣ��վ����ϵͳ
'��������Ƽ������ҳ�Ʒ
'�ͷ�����http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>
<%call checkmanage("05")
dim javastr1,rsjs,fs,f,action,hs,ls,btw,hg,bgw,iii,i
action=request.querystring("action")
If trim(request("hs"))<>"" Then hs=CInt(trim(request("hs")))
If trim(request("ls"))<>"" Then ls=CInt(trim(request("ls")))
If trim(request("bgw"))<>"" Then btw=CInt(trim(request("btw")))
If trim(request("hg"))<>"" Then hg=CInt(trim(request("hg")))
If trim(request("bgw"))<>"" Then bgw=CInt(trim(request("bgw")))
If action="Production" Then
	javastr1="document.write('"
	Call OpenConn
             set rsjs=server.createobject("adodb.recordset")
			 sql="select * from DNJY_info where yz=1 and fsohtm=1"
			 if jiaoyi=0 Then sql=sql&" and zt="&jiaoyi
             if overdue=0 Then sql=sql&" and dqsj>"&DN_Now&""
             sql=sql&" order by fbsj "&DN_OrderType&""
			 rsjs.open sql,conn,3,1
	'-------����ľ�����������������ɸ�ʽ--------------------------------
	iii=0
	i=0
	javastr1=javastr1+"<table border=0 cellpadding=0 cellspacing=0  width="&bgw&">"
	do while not rsjs.bof and not rsjs.eof
	i=i+1
	IF iii Mod ls=0 then javastr1=javastr1&  "<tr height="""&hg&""">"
	javastr1=javastr1+"<td width="""&bgw&""">��<a href=""http://"&web&"/"&strInstallDir&"html/"&server.HTMLEncode(trim(rsjs("id")&" "))&".html"" target=""_blank"">"
	if len(trim(rsjs("biaoti")))>btw then
	javastr1=javastr1+""&server.htmlencode(trim(mid(rsjs("biaoti"),1,btw-2)&" "))&"...</a></td>"
	Else
	javastr1=javastr1+""&server.htmlencode(trim(rsjs("biaoti")&" "))&"</a></td>"	
	End if
	iii=iii+1
	IF iii Mod ls=0 then javastr1=javastr1&  "</tr>"
    if i>=hs then exit do 
	rsjs.movenext
	Loop	
	IF iii Mod ls<>0 then
	Do While not iii Mod ls=0
	javastr1=javastr1&  "<td></td>"
	iii=iii+1
	Loop
	javastr1=javastr1&  "</tr>"
	End IF
	javastr1=javastr1&  "</table>"
'-------�������ɸ�ʽ����----------------------------------------

			  rsjs.close
			  set rsjs=nothing	  

javastr1=javastr1+"');"&vbcrlf
		Set fs = CreateObject("Scripting.FileSystemObject")
		
		Set f = fs.CreateTextFile(Server.MapPath("/")&"/"&""&strInstallDir&"sitenew/sitenew.js", True)
		f.write(javastr1)
		f.close
		Set f = nothing
		Set fs = Nothing
		Call CloseDB(conn)
		response.Write "<script language=javascript>alert('JS����������Ϣ�ļ����ɳɹ�!');location.replace('sitenew.asp');</script>"

End if%>
<link href="../css/style.css" rel="stylesheet" type="text/css">
<BODY>
<style type="text/css">
<!--
.style1 {color: #FFFFFF;
	font-weight: bold;
}
.style2 {color: #FF0000}
-->
</style>
<table width="100%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="#6699cc">

  <TR>
    <TD height="20" bgcolor="#799AE1" align="center" colspan="13"><FONT color="#FFFFFF" style="font-size:14px">�� �� �� �� �� �� �� ��</FONT></TD>
 </TR>
  <FORM name="form" method="post" action="?action=Production" >
  
    <tr> 
      <td height="30" colspan="3" align="center">&nbsp;</td>
    </tr>

                      <tr>
                        <td height="25">����������<input style="BACKGROUND-COLOR: #eeeeee; BORDER-BOTTOM: #FFFFFF 1px solid; BORDER-LEFT: #a2a2a2 1px solid; BORDER-RIGHT: #ffffff 1px solid; BORDER-TOP: #a2a2a2 1px solid; COLOR: #333333; FONT-SIZE: 9pt; HEIGHT: 19px;" type="text" name="hs" size="10" value="10" onKeyUp="this.value=this.value.replace(/\D/g,'')" onafterpaste="this.value=this.value.replace(/\D/g,'')"> �� </td>
                      </tr> 
                     <tr>
                        <td height="25">����������<input style="BACKGROUND-COLOR: #eeeeee; BORDER-BOTTOM: #FFFFFF 1px solid; BORDER-LEFT: #a2a2a2 1px solid; BORDER-RIGHT: #ffffff 1px solid; BORDER-TOP: #a2a2a2 1px solid; COLOR: #333333; FONT-SIZE: 9pt; HEIGHT: 19px;" type="text" name="ls" size="10" value="1" onKeyUp="this.value=this.value.replace(/\D/g,'')" onafterpaste="this.value=this.value.replace(/\D/g,'')"> �� </td>
                      </tr>
                      </tr> 
                     <tr>
                        <td height="25">���ⳤ�ȣ�<input style="BACKGROUND-COLOR: #eeeeee; BORDER-BOTTOM: #FFFFFF 1px solid; BORDER-LEFT: #a2a2a2 1px solid; BORDER-RIGHT: #ffffff 1px solid; BORDER-TOP: #a2a2a2 1px solid; COLOR: #333333; FONT-SIZE: 9pt; HEIGHT: 19px;" type="text" name="btw" size="10" value="20" onKeyUp="this.value=this.value.replace(/\D/g,'')" onafterpaste="this.value=this.value.replace(/\D/g,'')"> �ַ� </td>
                      </tr>
                     <tr>
                        <td height="25">����иߣ�<input style="BACKGROUND-COLOR: #eeeeee; BORDER-BOTTOM: #FFFFFF 1px solid; BORDER-LEFT: #a2a2a2 1px solid; BORDER-RIGHT: #ffffff 1px solid; BORDER-TOP: #a2a2a2 1px solid; COLOR: #333333; FONT-SIZE: 9pt; HEIGHT: 19px;" type="text" name="hg" size="10" value="10" onKeyUp="this.value=this.value.replace(/\D/g,'')" onafterpaste="this.value=this.value.replace(/\D/g,'')">  </td>
                      </tr> 	
                     <tr>
                        <td height="25">����ȣ�<input style="BACKGROUND-COLOR: #eeeeee; BORDER-BOTTOM: #FFFFFF 1px solid; BORDER-LEFT: #a2a2a2 1px solid; BORDER-RIGHT: #ffffff 1px solid; BORDER-TOP: #a2a2a2 1px solid; COLOR: #333333; FONT-SIZE: 9pt; HEIGHT: 19px;" type="text" name="bgw" size="10" value="250" onKeyUp="this.value=this.value.replace(/\D/g,'')" onafterpaste="this.value=this.value.replace(/\D/g,'')">  </td>
                      </tr> 						  
    <tr> 
      <td height="30" colspan="3" align="center"> 
        <input type="submit" name="Submit" value="�ύ" class="input" >
        
        <input type="reset" name="Submit2" value="����" class="input"> 
 </td>
    </tr>
  </form>
  <tr> 
    <td height="30" colspan="3" align="center">&nbsp;</td>
  </tr>
</table>
<table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#ADAED6" width="100%">
  
    <tr bgcolor="#BDBEDE">
      <td height="28" style="border-bottom-style: solid; border-bottom-width: 1"><span class="style1"><span style="font-size: 9pt"><span style="font-size: 9pt"><font size="3"><font size="2">&nbsp;</font></font>��ҳ�������˵��</span></span></span></td>
    </tr>
    <tr bgcolor="#FFFFFF">
      <td style="border-top-style: solid; border-top-width: 1"><table width="98%"  border="0" align="center" cellpadding="0" cellspacing="0">
        <tr>
          <td><p>���ɵ�JS��sitenewĿ¼�µ���ص���js�ļ���վ�ڻ�վ���ֱ��ʹ�����������ã������õ���Ϣ�Ǿ�̬�ģ�������Ϣ����Ҫ�ֹ�����<br>���ļ������ϵ���ID�ſɵ��õ�����Ϣ��sitenew<font color='red'>c1</font>_<font color='red'>c2</font>_<font color='red'>c3</font>.js������c1��c2��c3�ֱ�Ϊ����������ID�š������Ƿ�վ����Ա��������Ӧ�������ļ���</p>
           
            <p class="style2">
              <textarea name="textarea" cols="88" rows="6"><!--���뿪ʼ-->
<SCRIPT language=JavaScript src="http://<%=web%>/<%=strInstallDir%>sitenew/sitenew.js" type=text/JavaScript></SCRIPT>
<!--�������-->
              </textarea>
            </p>
            <p>JS��������ҪFSO֧��</p></td>
        </tr>
      </table></td>
    </tr>
	<table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#ADAED6" width="100%">
  
    <tr bgcolor="#BDBEDE">
      <td height="28" style="border-bottom-style: solid; border-bottom-width: 1">
<b>����Ԥ����</b>&nbsp;&nbsp;<input type=button value=ˢ��Ч�� onClick="location.reload()"><p>
<!--���뿪ʼ-->
<SCRIPT language=JavaScript src="../sitenew/sitenew.js" type=text/JavaScript></SCRIPT>
<!--�������-->
</td>
        </tr>
      </table
></table>

