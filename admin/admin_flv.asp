<!--#include file="../conn.asp"-->
<!--#include file="../SqlIn.Asp"-->
<!--#include file="../config.asp"-->
<!--#include file=cookies.asp-->
<!--#include file=../ads/Html2Js.asp-->
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
<title>FLV��Ƶ�������</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" type="text/css" href="../css/css.css">
<SCRIPT language="javascript" src="admin.js"></SCRIPT>
<IFRAME width="260" height="165" id="colourPalette" src="nc_selcolor.htm" style="visibility:hidden; position: absolute; left: 0px; top: 0px;border:1px gray solid" frameborder="0" scrolling="no" ></IFRAME>
<BODY>
<link href="../css/style.css" rel="stylesheet" type="text/css">
</head>

<%Call OpenConn
dim action,pic,n,JsCode,texts,filesurl
Dim flvconfig,flvconfig0,flvconfig1,flvconfig2,flvconfig3,flvconfig4,flvconfig5,flvconfig6,flvconfig7,flvconfig8,flvconfig9,flvconfig10,flvconfig11,flvconfig12,flvconfig13,flvconfig14,flvconfig15,flvconfig16,flvconfig17,flvconfig18,flvconfig19,flvconfig20,flvconfig21,flvconfig22,flvconfig23,flvconfig24,flvconfig25,flvconfig26,flvconfig27,flvconfig28
action=request("ok")
if action="" then 
Set rs = conn.Execute("select * from DNJY_adhd") 
flvconfig=rs("flv")
flvconfig=split(flvconfig,"|")
flvconfig0=flvconfig(0)
flvconfig1=flvconfig(1)
flvconfig2=flvconfig(2)
flvconfig3=flvconfig(3)
flvconfig4=flvconfig(4)
flvconfig5=flvconfig(5)
flvconfig6=flvconfig(6)
flvconfig7=flvconfig(7)
flvconfig8=flvconfig(8)
flvconfig9=flvconfig(9)
flvconfig10=flvconfig(10)
flvconfig11=flvconfig(11)
flvconfig12=flvconfig(12)
flvconfig13=flvconfig(13)
flvconfig14=flvconfig(14)
flvconfig15=flvconfig(15)
flvconfig16=flvconfig(16)
flvconfig17=flvconfig(17)
flvconfig18=flvconfig(18)
flvconfig19=flvconfig(19)
flvconfig20=flvconfig(20)
flvconfig21=flvconfig(21)
flvconfig22=flvconfig(22)
flvconfig23=flvconfig(23)
flvconfig24=flvconfig(24)
flvconfig25=flvconfig(25)
flvconfig26=flvconfig(26)
flvconfig27=flvconfig(27)
flvconfig28=flvconfig(28)
%>

	<table width="98%" border="1"  style="border-collapse: collapse; border-style: dotted; border-width: 0px"  bordercolor="#333333" cellspacing="0" cellpadding="2">
	<form action=admin_flv.asp method=post name=form>
	
	<tr class=backs><td colspan=3 class=td height=18 align="center">FLV��Ƶ�������</td></tr>

	<tr><td width=20% align=right height="18">��Ƶһ����</td><td> <input type=text value="<%=flvconfig0%>" name=flvconfig0 size=40 maxlength=100></td></tr>
	<tr><td width=20% align=right height="18">��Ƶһ�ļ���ַ</td><td> <input type=text value="<%=flvconfig1%>" name=flvconfig1 size=40 maxlength=100>
	 <INPUT TYPE="button" value="�ϴ�FLV��Ƶ�ļ�һ"  onClick="window.open('../ads/DNJY_upload_ad.asp?ttup=flv&fupname=flv1&frmname=flvconfig1','blank_','scrollbars=yes,resizable=no,top=40,left=150,width=500,height=150')"> <br><FONT color=#AB3801>��ʽ������·����http://www.ip126.com/UploadFiles/flv/flv1.flv�������·�����ʺϸ�Ŀ¼�°�װ����/UploadFiles/flv/flv1.flv</font></td></tr>		
	
	<tr><td width=20% align=right height="18">��Ƶ������</td><td> <input type=text value="<%=flvconfig2%>" name=flvconfig2 size=40 maxlength=100></td></tr>
	<tr><td width=20% align=right height="18">��Ƶ���ļ���ַ</td><td> <input type=text value="<%=flvconfig3%>" name=flvconfig3 size=40 maxlength=100>
	  <INPUT TYPE="button" value="�ϴ�FLV��Ƶ�ļ���"  onClick="window.open('../ads/DNJY_upload_ad.asp?ttup=flv&fupname=flv2&frmname=flvconfig3','blank_','scrollbars=yes,resizable=no,top=40,left=150,width=500,height=150')"> <br><FONT color=#AB3801>��ʽ������·����http://www.ip126.com/UploadFiles/flv/flv2.flv�������·�����ʺϸ�Ŀ¼�°�װ����/UploadFiles/flv/flv2.flv</font></td></tr>
	
	<tr><td width=20% align=right height="18">��Ƶ������</td><td> <input type=text value="<%=flvconfig4%>" name=flvconfig4 size=40 maxlength=100></td></tr>
	<tr><td width=20% align=right height="18">��Ƶ���ļ���ַ</td><td> <input type=text value="<%=flvconfig5%>" name=flvconfig5 size=40 maxlength=100>
	  <INPUT TYPE="button" value="�ϴ�FLV��Ƶ�ļ���"  onClick="window.open('../ads/DNJY_upload_ad.asp?ttup=flv&fupname=flv3&frmname=flvconfig5','blank_','scrollbars=yes,resizable=no,top=40,left=150,width=500,height=150')"> <br><FONT color=#AB3801>��ʽ������·����http://www.ip126.com/UploadFiles/flv/flv3.flv�������·�����ʺϸ�Ŀ¼�°�װ����/UploadFiles/flv/flv3.flv</font></td></tr>

	<tr><td width=20% align=right height="18">��Ƶ�ı���</td><td> <input type=text value="<%=flvconfig6%>" name=flvconfig6 size=40 maxlength=100></td></tr>
	<tr><td width=20% align=right height="18">��Ƶ���ļ���ַ</td><td> <input type=text value="<%=flvconfig7%>" name=flvconfig7 size=40 maxlength=100>
	  <INPUT TYPE="button" value="�ϴ�FLV��Ƶ�ļ���"  onClick="window.open('../ads/DNJY_upload_ad.asp?ttup=flv&fupname=flv4&frmname=flvconfig7','blank_','scrollbars=yes,resizable=no,top=40,left=150,width=500,height=150')"> <br><FONT color=#AB3801>��ʽ������·����http://www.ip126.com/UploadFiles/flv/flv4.flv�������·�����ʺϸ�Ŀ¼�°�װ����/UploadFiles/flv/flv4.flv</font></td></tr>	 

	<tr><td width=20% align=right height="18">��Ƶ�����</td><td> <input type=text value="<%=flvconfig8%>" name=flvconfig8 size=40 maxlength=100></td></tr>
	<tr><td width=20% align=right height="18">��Ƶ���ļ���ַ</td><td> <input type=text value="<%=flvconfig9%>" name=flvconfig9 size=40 maxlength=100>
	 <INPUT TYPE="button" value="�ϴ�FLV��Ƶ�ļ���"  onClick="window.open('../ads/DNJY_upload_ad.asp?ttup=flv&fupname=flv5&frmname=flvconfig9','blank_','scrollbars=yes,resizable=no,top=40,left=150,width=500,height=150')"> <br><FONT color=#AB3801>��ʽ������·����http://www.ip126.com/UploadFiles/flv/flv5.flv�������·�����ʺϸ�Ŀ¼�°�װ����/UploadFiles/flv/flv5.flv</font></td></tr>	

	<tr><td width=20% align=right height="18"></td></tr>	
	<tr><td width=20% align=right height="18">��������</td><td colspan=3>
��������� <input type=text value="<%=flvconfig10%>" name=flvconfig10 size=4 maxlength=3>����<br>
�������߶� <input type=text value="<%=flvconfig11%>" name=flvconfig11 size=4 maxlength=3>����<br>
�Ƿ��Զ�����<input type="radio" value="1" <%if flvconfig12="1" then%>checked<%end if%> name="flvconfig12">�Զ�
<input type="radio" value="0" <%if flvconfig12="0" then%>checked<%end if%> name="flvconfig12">���Զ�<br>
�Ƿ���������<input type="radio" value="1" <%if flvconfig13="1" then%>checked<%end if%> name="flvconfig13">����
<input type="radio" value="0" <%if flvconfig13="0" then%>checked<%end if%> name="flvconfig13">������ <br>
�Ƿ��������<input type="radio" value="1" <%if flvconfig14="1" then%>checked<%end if%> name="flvconfig14">���
<input type="radio" value="0" <%if flvconfig14="0" then%>checked<%end if%> name="flvconfig14">˳��<br>
�Ƿ���ʾʱ��<input type="radio" value="1" <%if flvconfig15="1" then%>checked<%end if%> name="flvconfig15">��ʾ
<input type="radio" value="0" <%if flvconfig15="0" then%>checked<%end if%> name="flvconfig15">����ʾ<br>
��������ʾ&nbsp;&nbsp;<input type="radio" value="1" <%if flvconfig16="1" then%>checked<%end if%> name="flvconfig16">һֱ��ʾ <input type="radio" value="2" <%if flvconfig16="2" then%>checked<%end if%> name="flvconfig16">�����ͣʱ��ʾ <input type="radio" value="3" <%if flvconfig16="3" then%>checked<%end if%> name="flvconfig16">�����ͣ����ʾ <input type="radio" value="0" <%if flvconfig16="0" then%>checked<%end if%> name="flvconfig16">����ʾ<br>
���ſ�������ɫ <input type=text value="<%=flvconfig17%>" name=flvconfig17 size=8 maxlength=6 style="background:<%=flvconfig17%>" onClick="Getcolor(this)">(16�������ֱ�ʾ ����000033)<br>
������������ɫ <input type=text value="<%=flvconfig18%>" name=flvconfig18 size=8 maxlength=6 style="background:<%=flvconfig18%>" onClick="Getcolor(this)">(16�������ֱ�ʾ ����ffffff)<br>
����ͼ����ɫ&nbsp;&nbsp; <input type=text value="<%=flvconfig19%>" name=flvconfig19 size=8 maxlength=6 style="background:<%=flvconfig19%>" onClick="Getcolor(this)">(16�������ֱ�ʾ ����66ff00)<br>
�����ͣ����ɫ <input type=text value="<%=flvconfig20%>" name=flvconfig20 size=8 maxlength=6 style="background:<%=flvconfig20%>" onClick="Getcolor(this)">(16�������ֱ�ʾ ����FFFFFF)<br>
ӰƬ����ʱ��&nbsp;&nbsp; <input type=text value="<%=flvconfig21%>" name=flvconfig21 size=8 maxlength=6>(��λ:��)<br>
Ĭ����������&nbsp;&nbsp; <input type=text value="<%=flvconfig22%>" name=flvconfig22 size=8 maxlength=6>(0-100 ����ֵ)<br>

����LOGO��ַ &nbsp;&nbsp;<input type=text value="<%=flvconfig23%>" name=flvconfig23 size=40 maxlength=100>(Ӣ��)<br>
ͼƬLOGO��ַ &nbsp;&nbsp;<input type=text value="<%=flvconfig24%>" name=flvconfig24 size=40 maxlength=100>(��дLOGO��ַ,֧��ͼƬ��ʽ��swf��ʽ)<br>
<!--����ǰswf�ļ���ַ<input type=text value="<%=flvconfig25%>" name=flvconfig25 size=40 maxlength=100>(ӰƬ��ʼ����ǰ���ⲿ��ȡswf�ļ�������)<br>
���ź�swf�ļ���ַ<input type=text value="<%=flvconfig26%>" name=flvconfig26 size=40 maxlength=100>(ӰƬ���Ž�������ⲿ��ȡswf�ļ���ӰƬ��Ϣ��)-->
</td></tr>
	
	<tr><td colspan=3>
<INPUT name="ok" TYPE="hidden" value="ok"><INPUT name=action TYPE="submit" value="�������ò�����JS�ļ�"></td></tr>
</form></table>
	<%end if%>

<%              
if action="ok" then

	if trim(request.form("flvconfig0"))="" or trim(request.form("flvconfig1"))=""  then
	response.write "<script language='javascript'>"
	response.write "alert('��һ����Ƶ���һ��Ҫ��д��������ȷ����������һҳ��');"
	response.write "location.href='javascript:history.go(-1)';"
	response.write "</script>"
	response.end
	end if

	if trim(request.form("flvconfig2"))<>"" and trim(request.form("flvconfig3"))=""  then
	response.write "<script language='javascript'>"
	response.write "alert('��Ƶ�������������Ƶ��ַ��Ӧ��������ȷ����������һҳ��');"
	response.write "location.href='javascript:history.go(-1)';"
	response.write "</script>"
	response.end
	end if

	if trim(request.form("flvconfig4"))<>"" and trim(request.form("flvconfig5"))=""  then
	response.write "<script language='javascript'>"
	response.write "alert('��Ƶ�������������Ƶ��ַ��Ӧ��������ȷ����������һҳ��');"
	response.write "location.href='javascript:history.go(-1)';"
	response.write "</script>"
	response.end
	end if

	if trim(request.form("flvconfig6"))<>"" and trim(request.form("flvconfig7"))=""  then
	response.write "<script language='javascript'>"
	response.write "alert('��Ƶ�ı����������Ƶ��ַ��Ӧ��������ȷ����������һҳ��');"
	response.write "location.href='javascript:history.go(-1)';"
	response.write "</script>"
	response.end
	end if

	if trim(request.form("flvconfig8"))<>"" and trim(request.form("flvconfig9"))=""  then
	response.write "<script language='javascript'>"
	response.write "alert('��Ƶ������������Ƶ��ַ��Ӧ��������ȷ����������һҳ��');"
	response.write "location.href='javascript:history.go(-1)';"
	response.write "</script>"
	response.end
	end if

flvconfig0=request.form("flvconfig0")
flvconfig1=request.form("flvconfig1")
flvconfig2=request.form("flvconfig2")
flvconfig3=request.form("flvconfig3")
flvconfig4=request.form("flvconfig4")
flvconfig5=request.form("flvconfig5")
flvconfig6=request.form("flvconfig6")
flvconfig7=request.form("flvconfig7")
flvconfig8=request.form("flvconfig8")
flvconfig9=request.form("flvconfig9")
flvconfig10=request.form("flvconfig10")
flvconfig11=request.form("flvconfig11")
flvconfig12=request.form("flvconfig12")
flvconfig13=request.form("flvconfig13")
flvconfig14=request.form("flvconfig14")
flvconfig15=request.form("flvconfig15")
flvconfig16=request.form("flvconfig16")
flvconfig17=Right(request.form("flvconfig17"),6)
flvconfig18=Right(request.form("flvconfig18"),6)
flvconfig19=Right(request.form("flvconfig19"),6)
flvconfig20=Right(request.form("flvconfig20"),6)
flvconfig21=request.form("flvconfig21")
flvconfig22=request.form("flvconfig22")
flvconfig23=request.form("flvconfig23")
flvconfig24=request.form("flvconfig24")
flvconfig25=request.form("flvconfig25")
flvconfig26=request.form("flvconfig26")
flvconfig27=request.form("flvconfig27")
flvconfig28=request.form("flvconfig28")

flvconfig=flvconfig0+"|"+flvconfig1+"|"+flvconfig2+"|"+flvconfig3+"|"+flvconfig4+"|"+flvconfig5+"|"+flvconfig6+"|"+flvconfig7+"|"+flvconfig8+"|"+flvconfig9+"|"+flvconfig10+"|"+flvconfig11+"|"+flvconfig12+"|"+flvconfig13+"|"+flvconfig14+"|"+flvconfig15+"|"+flvconfig16+"|"+flvconfig17+"|"+flvconfig18+"|"+flvconfig19+"|"+flvconfig20+"|"+flvconfig21+"|"+flvconfig22+"|"+flvconfig23+"|"+flvconfig24+"|"+flvconfig25+"|"+flvconfig26+"|"+flvconfig27+"|"+flvconfig28

     Set rs=Server.CreateObject("ADODB.Recordset")
	 sql="select * from DNJY_adhd"
	 rs.open sql,conn,1,3
 	 rs("flv")=flvconfig 	 
	rs.update

'��ȡ����
flvconfig=rs("flv")
flvconfig=split(flvconfig,"|")

texts=flvconfig0
If flvconfig2<>"" Then texts=texts+"|"+flvconfig2
If flvconfig4<>"" Then texts=texts+"|"+flvconfig4
If flvconfig6<>"" Then texts=texts+"|"+flvconfig6
If flvconfig8<>"" Then texts=texts+"|"+flvconfig8
filesurl=flvconfig1
If flvconfig3<>"" Then filesurl=filesurl+"|"+flvconfig3
If flvconfig5<>"" Then filesurl=filesurl+"|"+flvconfig5
If flvconfig7<>"" Then filesurl=filesurl+"|"+flvconfig7
If flvconfig9<>"" Then filesurl=filesurl+"|"+flvconfig9

'����js����
JsCode="<SCRIPT type=text/javascript>" & vbCrLf & _
"var swf_width="""&flvconfig10 &"""//���������" & vbCrLf & _
"var swf_height="&flvconfig11 &"//�������߶�" & vbCrLf & _
"var texts='"&texts&"'//��������������ʹ��|�ֿ�,������filesurl���һ��" & vbCrLf & _
"var filesurl='"&filesurl&"'//flv�ļ����Ե�ַ���������ʹ��|�ֿ�,������texts���һ��" & vbCrLf & _
"var config='&IsAutoPlay="&flvconfig12&"&IsContinue="&flvconfig13&"&IsRandom="&flvconfig14&"&IsShowTime="&flvconfig15&"&IsShowBar="&flvconfig16&"&BarColor=0x"&flvconfig17&"&TextColor=0x"&flvconfig18&"&GlowColor=0x"&flvconfig19&"&IconColor=0x"&flvconfig20&"&BarPosition=1&BufferTime="&flvconfig21&"&DefaultVolume="&flvconfig22&"&LogoText="&flvconfig23 &"&LogoUrl="&flvconfig24&"&BeginSwf="&flvconfig25&"&EndSwf="&flvconfig26&"'" & vbCrLf & _
"document.write('<object classid=""clsid:d27cdb6e-ae6d-11cf-96b8-444553540000"" codebase=""http://fpdownload.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=6,0,0,0"" width=""'+ swf_width +'"" height=""'+ swf_height +'"">');" & vbCrLf & _
"document.write('<param name=""movie"" value=""/"&strInstallDir&"img/flv.swf""><param name=""quality"" value=""high"">');" & vbCrLf & _
"document.write('<param name=""menu"" value=""false""><param name=""allowFullScreen"" value=""true"" />');" & vbCrLf & _
"document.write('<param name=""FlashVars"" value=""vcastr_file='+filesurl+'&vcastr_title='+texts+'&vcastr_config='+config+'"">');" & vbCrLf & _
"document.write('<embed src=""/"&strInstallDir&"img/flv.swf"" allowFullScreen=""true"" FlashVars=""vcastr_file='+filesurl+'&vcastr_title='+texts+'&vcastr_config='+config+'"" menu=""false"" quality=""high"" width=""'+ swf_width +'"" height=""'+ swf_height +'"" type=""application/x-shockwave-flash"" pluginspage=""http://www.macromedia.com/go/getflashplayer""/>');" & vbCrLf & "document.write('</object>'); " & vbCrLf & _
"</SCRIPT>" & vbCrLf & _
""

JsCode=Html2Js(JsCode)
JsCode="document.write('"&JsCode&"')"
JsCode = "<!--" & vbCrLf & JsCode & vbCrLf &"//-->" 
dim fs,f
Set fs = CreateObject("Scripting.FileSystemObject")
Set f = fs.CreateTextFile(Server.MapPath("/")&"/"&""&strInstallDir&"UploadFiles/js/flv.js", True)
f.write(JsCode)
f.close
Set f = nothing
Set fs = Nothing


Call CloseRs(rs)
Call CloseDB(conn)

	response.write "<script language='javascript'>"
	response.write "alert('FLV������óɹ�������js���룡');"
	response.write "location.href='admin_flv.asp';"
	response.write "</script>"
	response.end
end if%>

</body>
<TABLE align="center">
<TR>
	<TD><button onClick="document.location.reload()">ˢ�¿�Ч��</button>Ч��Ԥ�� :<br>ע��1������������߱�����.flv��Ƶ�������ܲ��ţ������ܲ���Ҫ�����ã������<a href="http://www.ip126.com/Article/fwq/940.html" target="_blank"><FONT color=#0000ff>Flv��ʽ��Ƶ�Ľ������</font></a>��<br>2����Ƶ��ַ�ͱ���Ҫ��ʵ����������û�о����գ��������<br>
<span class="f14b" >
<script src="../UploadFiles/js/flv.js"></script>


<%Call CloseRs(rs)
Call CloseDB(conn)%></TD>
	<TD>���÷�ʽ��&lt;script src="UploadFiles/js/flv.js">&lt;/script></TD>
</TR>
</TABLE><br>&nbsp;<p>
</html>

