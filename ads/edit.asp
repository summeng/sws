<!--#include file="../conn.asp"-->
<!--#include file="../config.asp"-->
<!--#include file=cookies.asp-->
<%
'=====================================
'���칩����Ϣ��վ����ϵͳ
'��������Ƽ������ҳ�Ʒ
'�ͷ�����http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>
<%call checkmanage("13")%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>COAD System</title>
<link href="../css/style.css" rel="stylesheet" type="text/css">
<script language=javascript src=../Include/mouse_on_title.js></script>
<SCRIPT LANGUAGE="JavaScript" src="images/js.js"></SCRIPT>
<SCRIPT LANGUAGE="JavaScript" src="images/Time_selector.js"></SCRIPT>
<script language="JavaScript" type="text/JavaScript">
//��վlogo�ϴ�
function JM_wu(ob){
 ob.style.display="none";
}
function JM_you(ob){
 ob.style.display="";
}
function uppic(model,frmname) {
//ȷ���ϴ��ļ�������ѡ�������������жϱ�Ų��Ա������
id=document.form.ADID.value;
id1=document.form.city_one.value;
id2=document.form.city_two.value;
id3=document.form.city_three.value;
if (id1=='') id1=0
if (id2=='') id2=0
if (id3=='') id3=0
if (id=='') {
 alert("������д���ID�ţ�");
 document.form.ADID.focus();
return false;}
window.open("../ads/DNJY_upload_ad.asp?ttup=adv&fupname="+id+model+"_"+id1+"_"+id2+"_"+id3+"&frmname="+frmname,"blank_","scrollbars=yes,resizable=no,top=40,left=150,width=500,height=150")
}

function CheckAll(form) {
 for (var i=0;i<form.elements.length;i++) {
 var e = form.elements[i];
 if (e.name != 'chkall') e.checked = form.chkall.checked; 
 }
 }
</script>
<style type="text/css">
<!--
body {
	background-color: #F7F3F7;
}
-->
</style>
<style type="text/css">
<!--
.style1 {color: #FFFFFF;
	font-weight: bold;
}
-->
</style>
</head>
<body>
<%Call OpenConn
Dim id,city_oneid,city_twoid,city_threeid,ADhost,ADhost1,ADhost2,ADhost3,ADhost4,ADhost5,ADhost6,ADhost7,ADhost8,ADhost9,ADLink
If Request.QueryString("id")<> "" Then
city_oneid=TRIM(Request.Form("city_one"))
city_twoid=TRIM(Request.Form("city_two"))
city_threeid=TRIM(Request.Form("city_three"))
If city_oneid="" Then city_oneid=0
If city_twoid="" Then city_twoid=0
If city_threeid="" Then city_threeid=0
id=Request.QueryString("id")
Set rs = Server.CreateObject("ADODB.Recordset")
sql="select * from [DNJY_ad] where id<>"&Request.QueryString("id")&" and ADID='"&DangerEncode(Request.Form("ADID"))&"' and city_oneid="&city_oneid&" and city_twoid="&city_twoid&" and city_threeid="&city_threeid&""
rs.Open sql,conn,1,3
If not rs.eof Then
	Response.Write ("<script>alert(' ��������! \n\n ���������ID�ظ�,��ʹ������ID !');history.back();</script>")
	Call CloseRs(rs)
	Call CloseDB(conn)
	Response.end
End If
Call CloseRs(rs)
'��/���¹������
Set rs = Server.CreateObject("ADODB.Recordset")
sql="select ADID,ADType,ADSrc,ADHeight,ADWidth,adsjs,ADLink,ADAlt,ADStopViews,ADStopHits,ADStopDate,ADNote,ADViews,ADHits,ADjs,ADkg,city_oneid,city_twoid,city_threeid,Advertisement_host from [DNJY_ad] where id=" & id
rs.Open sql,conn,1,3
'�Ƿ����
If Request.Form("ChangeAD") <> "" Then
ADLink=DangerEncode(Request.Form("ADLink"))
If Left(ADLink,7)<>"http://" Then ADLink="http://"+ADLink
	rs("ADType")=DangerEncode(Request.Form("ADType"))
	rs("ADSrc")=DangerEncode(Request.Form("ADSrc"))
	If TRIM(Request.Form("ADHeight"))<>"" Then
	rs("ADHeight")=TRIM(Request.Form("ADHeight"))
	Else
	rs("ADHeight")=0
	End if
	If TRIM(Request.Form("ADWidth"))<>"" Then
	rs("ADWidth")=TRIM(Request.Form("ADWidth"))
	Else
	rs("ADWidth")=0
	End if
	rs("ADLink")=ADLink
	rs("ADAlt")=DangerEncode(Request.Form("ADAlt"))
	rs("ADStopViews")=TRIM(Request.Form("ADStopViews"))
	rs("ADStopHits")=TRIM(Request.Form("ADStopHits"))
	rs("ADStopDate")=TRIM(Request.Form("ADStopDate"))
	rs("ADNote")=DangerEncode(Request.Form("ADNote"))
	rs("ADjs")=TRIM(Request.Form("ADjs"))
	rs("ADkg")=TRIM(Request.Form("ADkg"))
	rs("city_oneid")=city_oneid
    rs("city_twoid")=city_twoid
    rs("city_threeid")=city_threeid
	rs("adsjs")=""&DangerEncode(Request.Form("ADID"))&"_"&city_oneid&"_"&city_twoid&"_"&city_threeid&".js"
	rs("Advertisement_host")=Request.Form("ADhost1")&"|"&Request.Form("ADhost2")&"|"&Request.Form("ADhost3")&"|"&Request.Form("ADhost4")&"|"&Request.Form("ADhost5")&"|"&Request.Form("ADhost6")&"|"&Request.Form("ADhost7")&"|"&Request.Form("ADhost8")&"|"&Request.Form("ADhost9")
	If Request.Form("ADRESET")="YES" Then
		rs("ADViews")=0
		rs("ADHits")=0
	End If
	rs.Update
	
	Response.Write "<script language=javascript>setTimeout(""alert('�޸Ĺ��ɹ���');" &"location.replace('createjs.asp?page="&request("page")&"&ADID="&rs("ADID")&"&city_oneid="&rs("city_oneid")&"&city_twoid="&rs("city_twoid")&"&city_threeid="&rs("city_threeid")&"')"",0)</script>"

End If
If Err <> 0 Then
			Response.Write "<font color=red size=2>����:"&Err.Description
			Response.End
End If
%>
<table width="100%" border="0" align="center" cellpadding="3" cellspacing="1" bgcolor="#6699cc">
  <form name=form action=edit.asp?id=<%=id%>  method=post onsubmit="return chkinput()">
    <tr align="center" bgcolor="#6699cc">
      <td height="28" colspan="2"><div align="center" ><font color="#FFFFFF"><b>�޸Ĺ��</b></font></div></td>
    </tr>
    <tr bgcolor="#FFFFFF">
      <td width="15%" align="right" bgcolor="#FAFAFA"><div align="center">���&nbsp;I&nbsp;D��</div></td>
      <td width="400">
      <INPUT   maxLength=14 size=20 name="dddd" value="<%=rs("ADID")%>" DISABLED> <font color="#FF0000">*</font></td>
    </tr><input type="hidden" name="ADID" value="<%=rs("ADID")%>">
    <tr bgcolor="#FFFFFF">
      <td width="15%" align="right" bgcolor="#FAFAFA"><div align="center">����������</div></td>
      <td width="400"><%
Dim rsi
set rsi=server.createobject("adodb.recordset")
set rsi=conn.execute("select * from DNJY_city where id>0 and twoid>0 and threeid=0")
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
                                     <%set rsi=conn.execute("select * from DNJY_city where id>0 and twoid=0 order by indexid asc")
if rsi.eof and rsi.bof then
response.write "<option value=''>û�з���</option>"
else
do until rsi.eof%>
                                     <OPTION value="<%=rsi("id")%>" <%if rsi("id")=rs("city_oneid") then%>selected<%end if%>><%=rsi("city")%></OPTION>
                                     <% rsi.movenext
    loop
end if
rsi.close
set rsi = nothing
%>
                                   </SELECT>
                                   <SELECT name="city_two" id="city_two" onChange="changelocation4(document.form.city_two.options[document.form.city_two.selectedIndex].value,document.form.city_one.options[document.form.city_one.selectedIndex].value)">
                                     <OPTION value="" selected>ѡ�����</OPTION>
                                     <%set rsi=conn.execute("select * from DNJY_city where id="&rs("city_oneid")&" and twoid<>0 and threeid=0")
if rsi.eof and rsi.bof then
response.write "<option value=''>û�з���</option>"
else
do until rsi.eof%>
                                     <OPTION value="<%=rsi("twoid")%>" <%if rsi("twoid")=rs("city_twoid") then%>selected<%end if%>><%=rsi("city")%></OPTION>
                                     <% rsi.movenext
    loop
	end if
rsi.close
set rsi = nothing%>
                                   </SELECT>
                                   <SELECT name="city_three" id="city_three">
                                     <OPTION value="" selected>ѡ�����</OPTION>
                                     <%set rsi=conn.execute("select * from DNJY_city where id="&rs("city_oneid")&" and twoid="&rs("city_twoid")&" and threeid<>0")
if rsi.eof and rsi.bof then
response.write "<option value=''>û�з���</option>"
else
do until rsi.eof%>
                                     <OPTION value="<%=rsi("threeid")%>" <%if rsi("threeid")=rs("city_threeid") then%>selected<%end if%>><%=rsi("city")%></OPTION>
                                     <% rsi.movenext
    loop
	end if
rsi.close
set rsi = nothing%>
                                   </SELECT></td>
    </tr>
    <tr bgcolor="#FFFFFF">
      <td width="15%" height="25" align="right" bgcolor="#FaFaFa"><div align="center">������ͣ�</div></td>
      <td><select name="ADType" size="1" class="input1" onChange="ChangeType(this.options[this.selectedIndex].value)">
          <option <%If rs("ADType")=1 Then Response.Write"selected"%> value="1">��ͨͼƬ</option>
          <option <%If rs("ADType")=2 Then Response.Write"selected"%> value="2">����������ʾ</option>
          <option <%If rs("ADType")=3 Then Response.Write"selected"%> value="3">���¸�����ʾ - ��</option>
          <option <%If rs("ADType")=4 Then Response.Write"selected"%> value="4">���¸�����ʾ - ��</option>
          <option <%If rs("ADType")=5 Then Response.Write"selected"%> value="5">ȫ��Ļ������ʧ</option>
          <option <%If rs("ADType")=6 Then Response.Write"selected"%> value="6">��ͨ��ҳ�Ի��� </option>
          <option <%If rs("ADType")=7 Then Response.Write"selected"%> value="7">���ƶ�͸���Ի��� </option>
          <option <%If rs("ADType")=8 Then Response.Write"selected"%> value="8">���´���</option>
          <option <%If rs("ADType")=9 Then Response.Write"selected"%> value="9">�����´���</option>
          <option <%If rs("ADType")=10 Then Response.Write"selected"%> value="10">����ʽ��棨�̶���</option>
          <option <%If rs("ADType")=11 Then Response.Write"selected"%> value="11">����ʽ��棨������</option>
		  <option <%If rs("ADType")=12 Then Response.Write"selected"%> value="12">���������</option>  
		  <option <%If rs("ADType")=13 Then Response.Write"selected"%> value="13">��ͨ���ֹ��</option>
          <option style="background-color:#F8F4F5 ;color: #FF0000" <%If rs("ADType")=14 Then Response.Write"selected"%> value="14">��д������</option>
		  <option <%If rs("ADType")=15 Then Response.Write"selected"%> value="15">FlashͼƬ���</option>
		  <option <%If rs("ADType")=16 Then Response.Write"selected"%> value="16">��Ƶ���</option>
      </select> ��ѡ������������ʾ��������л������б��޷�������</td>
    </tr>
    <tr bgcolor="#FFFFFF">
      <td width="15%" height="25" align="right" bgcolor="#FaFaFa"><div align="center">ͼƬ��ַ��</div></td>
      <td><%If rs("ADType")= 6 Then
    	Response.Write ("<textarea rows=3 name=ADSrc cols=27 class=input2>"&server.HTMLencode(rs("ADSrc"))&"</textarea>")
	else%>
          <INPUT name="ADSrc" type="text" class="input1" value="<%=rs("ADSrc")%>" size=30>
          <%end if%> <INPUT alt="�뵥����������ϴ�ͼƬ<br>����дͼƬ����ַ����<font color=blue>http://www.ip126.com/adv/mp3.jpg</font><br>����дվ�ڵ�ͼƬ·�����ļ�������<font color=blue>adv/001.jpg</font>" TYPE="button" value="���ͼƬ�ϴ�" onclick="javascript:uppic('','ADSrc');">
      *<a href="#" onclick=openhelp("ext") title="����鿴����"><font color="#0000FF">�ϴ���ͨͼƬ��FLASHͼƬ</font></a>������д��Ƶ�ļ���ַ <a href="javascript:void(0)" onclick = "document.getElementById('light').style.display='block';document.getElementById('fade').style.display='block'" title="��˲鿴��Ƶ��ʽҪ��"><font color="#FF0000">[��Ƶ������˵��]</font></a></td>
    </tr>


<!--����ҳ�����ݿ�ʼ-->
<div id="light" class="white_content">
<div class="white_content_top"><span class="white_content_top_span1">��Ƶ������˵��</span><span class="white_content_top_span2"><a href="javascript:void(0)" onclick = "document.getElementById('light').style.display='none';document.getElementById('fade').style.display='none'">�ر�</a></span></div><span class="fRed">&nbsp;&nbsp;&nbsp;&nbsp;1����Ƶ�ļ���ַ��������ַ�����һ�����·������Ƶ�ļ����ڱ�����������ַ��ʽ��/ads/pic/abc.wmv�����Ǿ���·������Ƶ�ļ��ɷ��ڱ���������Ҳ��Զ�̣���ַ��ʽΪhttp://www.abc.com/video/abc.wmv��<br>&nbsp;&nbsp;&nbsp;&nbsp;2����Ƶһ��ϴ�������ڱ�����������ֱ���ϴ�������FTP�ϴ���/ads/picĿ¼�¡�<br>&nbsp;&nbsp;&nbsp;&nbsp;3����Ƶ���Ҫ����Ա��밲װWindows Media Player 9.0���ϰ汾�����������������ţ�һ��xp�����ϰ汾����ϵͳ���߱�����<br>&nbsp;&nbsp;&nbsp;&nbsp;4����Ƶ��ʽ���ݣ�wmv��mpg��avi,rmvb�ȣ�����wmv��ʽ�ļ����СЩ��</div>
<div id="fade" class="black_overlay"></div>
<!--����ҳ�����ݽ���-->

    <tr bgcolor="#FFFFFF">
      <td width="15%" height="25" align="right" bgcolor="#FaFaFa"><div align="center">ͼƬ���</div></td>
      <td><INPUT name=ADWidth type="text" class="input1" onkeypress="return Num();" value="<%=rs("ADWidth")%>" size=8 maxlength=4>
        ��
      <INPUT name=ADHeight type="text" class="input1" onkeypress="return Num();" value="<%=rs("ADHeight")%>" size=8 maxlength=4> ͼƬ(Flash)���|�������|��Ƶ�����|ǰһ������������͹�����ֹ����ٶȣ�һ����5</td>
    </tr>
    <tr bgcolor="#FFFFFF" style="visibility:hide;">
      <td width="15%" height="25" align="right" bgcolor="#FaFaFa"><div align="center">���ӵ�ַ��</div></td>
      <td><INPUT name=ADLink type="text" class="input1" value="<%=rs("ADLink")%>" size=30 maxlength=150> ���ӵ�ַ|�������ӵ�ַ|������������ӵ�ַ</td>
    </tr>
    <tr bgcolor="#FFFFFF">
      <td width="15%" height="25" align="right" bgcolor="#FaFaFa"><div align="center">��ʾ���֣�</div></td>
      <td><INPUT name=ADAlt type="text" class="input1" value="<%=rs("ADAlt")%>" size=30 maxlength=50> �����ͣ��ʾ|��������͹������</td>
    </tr>
    <tr bgcolor="#FFFFFF">
      <td width="15%" height="25" align="right" bgcolor="#FaFaFa"><div align="center">��д���룺</div></td>
      <td><textarea name="ADjs" cols="70" rows="5" id="ADjs"><%=rs("ADjs")%></textarea></td>
    </tr>
    <tr bgcolor="#FFFFFF">
      <td width="15%" height="25" align="right" bgcolor="#FaFaFa"><div align="center">Ͷ�����ƣ�</div></td>
      <td><INPUT name=ADStopViews type="text" class="input1" onkeypress="return Num();" value="<%=rs("ADStopViews")%>" size=8 maxlength=10>
        ��
          <INPUT name=ADStopHits type="text" class="input1" onkeypress="return Num();" value="<%=rs("ADStopHits")%>" size=8 maxlength=10>
        ��
        <INPUT name=ADStopDate type="text" class="input1" value="<%=rs("ADStopDate")%>" size=12 maxlength=19 onclick="setday(this)">
      <a href="#" onclick=openhelp("stop") title="����鿴����">��ʾ����������ڣ�</a></td>
    </tr>
    <tr bgcolor="#FFFFFF">
      <td width="15%" height="25" align="right" bgcolor="#FaFaFa"><div align="center">����ͳ�ƣ�</div></td>
      <td><input type="checkbox" name="ADRESET" value="YES">
      ������ʾ�͵������</td>
    </tr>
    <tr bgcolor="#FFFFFF">
      <td width="15%" height="25" align="right" bgcolor="#FaFaFa"><div align="center">��ע�ͣ�</div></td>
      <td><INPUT name=ADNote type="text" class="input1" value="<%=rs("ADNote")%>" size=30 maxlength=100>
      ��ע����ʾ�ڹ����</td>
    <tr bgcolor="#FFFFFF">
      <td width="15%" height="25" align="right" bgcolor="#FaFaFa"><div align="center">�Ƿ���</div></td>
      <td><%if rs("adkg")=1 then%>               
                <input type="radio" name="adkg" value="1" id="adkg" checked>
                 �&nbsp;&nbsp;&nbsp;&nbsp;
                <input type="radio" name="adkg" value="0" id="adkg">
                ��ͣ</td>
                <%else%>                         
                <input type="radio" name="adkg" value="1" id="adkg">
                 �&nbsp;&nbsp;&nbsp;&nbsp;
                <input type="radio" name="adkg" value="0" id="adkg" checked>
                ��ͣ<%end if%></td>
    </tr>
<%
ADhost=rs("Advertisement_host")
ADhost=split(ADhost,"|")
ADhost1=ADhost(0)
ADhost2=ADhost(1)
ADhost3=ADhost(2)
ADhost4=ADhost(3)
ADhost5=ADhost(4)
ADhost6=ADhost(5)
ADhost7=ADhost(6)
ADhost8=ADhost(7)
ADhost9=ADhost(8)
%>	
    <tr bgcolor="#FFFFFF">
      <td width="15%" align="right" bgcolor="#FAFAFA"><div align="center">�������Ϣ��</div></td>
      <td height="17">����<INPUT name=ADhost1 value="<%=ADhost1%>" type="text" class="input1" size=10 maxlength=30>&nbsp;&nbsp;�绰<INPUT name=ADhost2  value="<%=ADhost2%>" type="text" class="input1" size=20 maxlength=30>&nbsp;&nbsp;�ֻ�<INPUT name=ADhost3 value="<%=ADhost3%>"  type="text" class="input1" size=20 maxlength=30>&nbsp;&nbsp;��ѶQQ<INPUT name=ADhost4 value="<%=ADhost4%>"  type="text" class="input1" size=20 maxlength=30>&nbsp;&nbsp;����<INPUT name=ADhost5 value="<%=ADhost5%>"  type="text" class="input1" size=30 maxlength=30><br>��ַ<INPUT name=ADhost6 value="<%=ADhost6%>"  type="text" class="input1" size=38 maxlength=30>&nbsp;&nbsp;�ʱ�<INPUT name=ADhost7 value="<%=ADhost7%>"  type="text" class="input1" size=20 maxlength=30>&nbsp;&nbsp;����<INPUT name=ADhost8 value="<%=ADhost8%>"  type="text" class="input1" size=17 maxlength=30> Ԫ&nbsp;&nbsp;��ע<INPUT name=ADhost9 value="<%=ADhost9%>"  type="text" class="input1" size=30 maxlength=30></td>
    </tr>	
    <tr bgcolor="#eeeeee">
      <td height="35" colspan="2" align="center">        <table width="100%"  border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td><div align="center">
			<input type="hidden" name="page" value="<%=request("page")%>">
              <input name="ChangeAD" type="submit" class="input1" value="ȷ���޸�">
            </div></td>
            <td><div align="center">
              <input name="ChangeAD" type="reset" class="input1" value="��ԭ����">
            </div></td>
          </tr>
        </table>
		
      </td>
    </tr>
  </form>
</table>
<CENTER><button onClick="document.location.reload()">ˢ�¿�Ч��</button> ���Ч����ʾ���������ʾ�����鿴�Ƿ�Ϊ���ڻ���ͣ��<br>
<%if rs("ADType")=2 Or rs("ADType")=3 Or rs("ADType")=4 Or rs("ADType")=5 Or rs("ADType")=6 Or rs("ADType")=7 Or rs("ADType")=8 Or rs("ADType")=9 Or rs("ADType")=10 Or rs("ADType")=11 Or rs("ADType")=12 Or rs("ADType")=13 Or rs("ADType")=14 Or rs("ADType")=15 then%>
<script src="/<%=strInstallDir%>ads/JS/<%=rs("ADID")%>_<%=rs("city_oneid")%>_<%=rs("city_twoid")%>_<%=rs("city_threeid")%>.js"></script>
<%ElseIf rs("ADType")=16 then
Response.Write "��Ƶ��治Ԥ��"
else%>
 <IMG SRC="<%=server.HTMLencode(rs("ADSrc"))%>" BORDER="0" ALT="">
 <%end if%> </CENTER>
<br>
<%
    'Response.Write "<Script language=javascript>ChangeType('"&rs("ADType")&"')</Script>"
	rs.Close
	set rs=nothing
	conn.Close
	set conn=nothing
	
Else
	Response.Write "û��ָ��Ҫ�༭��ID��"
End If

Rem ���˿��ܳ�����ķ���
Function DangerEncode(fString)
If not isnull(fString) Then
    fString = Replace(fString, CHR(13), "")
    fString = Replace(fString, CHR(10), "")
	fString = replace(fString, "'", """")
    fString = Trim(fString)
    DangerEncode = fString
End If
End Function
%>
</body>
</html>
       <form name='reg' action='regid.asp' method='post' target='CheckReg'>
          <input type='hidden' name='ADID' value=''>
        </form>
      <script language=javascript>
function checkreg()
{
  if (document.form.ADID.value=="")
	{
	alert("�������û�����");
	document.form.ADID.focus();
	return false;
	}
  else
    {
	document.reg.ADID.value=document.form.ADID.value;
	var popupWin = window.open('regid.asp', 'CheckReg', 'scrollbars=no,width=300,height=100');
	document.reg.submit();
	}
}
</script>