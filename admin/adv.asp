<!--#include file="../conn.asp"-->
<!--#include file="../SqlIn.Asp"-->
<!--#include file="../config.asp"-->
<!--#include file="../Include/Function.asp"-->
<!--#include file=cookies.asp-->
<!--#include file=../ads/Html2Js.asp-->
<%
'=====================================
'���칩����Ϣ��վ����ϵͳ
'��������Ƽ������ҳ�Ʒ
'�ͷ�����http://www.ip126.com
'����֧����̳ http://www.ip126.com/bbs
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>
<%call checkmanage("13")%>
<html>
<head>
<title>��վ����</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" type="text/css" href="../css/css.css">
<BODY>
<link href="../css/style.css" rel="stylesheet" type="text/css">
<SCRIPT language="javascript" src="admin.js"></SCRIPT>
<script language="javascript">
function changedbtype(dbtype){
  var accesstr=document.getElementById("accesstr");
  var sqltr=document.getElementById("sqltr");
  if(dbtype==0){
  	accesstr.style.display = '';
	sqltr.style.display = 'none';
  }else{
  	accesstr.style.display = 'none';
	sqltr.style.display = '';
  }
}
</script>
<IFRAME width="260" height="165" id="colourPalette" src="nc_selcolor.htm" style="visibility:hidden; position: absolute; left: 0px; top: 0px;border:1px gray solid" frameborder="0" scrolling="no" ></IFRAME>
</head>

<%Call OpenConn
dim action,pic,n
action=request("ok")
if action="" then 
Set rs = conn.Execute("select * from DNJY_adhd") 
%>

	<table width="98%" border="1"  style="border-collapse: collapse; border-style: dotted; border-width: 0px"  bordercolor="#333333" cellspacing="0" cellpadding="2">
	<form action=adv.asp method=post name=form>
	
	<tr class=backs><td colspan=3 class=td height=18 align="center"><b>�õƹ��B���ã���ҳ�õƹ��B���ֵ�����ȫ����ʾ���ڻ��������п�������ʾA��B��</b></td></tr>
	
	<tr><td width=20% align=right height="18">ͼƬһ��·�����ļ���</td><td> <input type=text value="<%=rs("pic1")%>" name=pic1 size=40 maxlength=100>
	<INPUT TYPE="button" value="����ϴ�ͼƬ"  onClick="window.open('../ads/DNJY_upload_ad.asp?ttup=slideb&fupname=g1&frmname=pic1','blank_','scrollbars=yes,resizable=no,top=40,left=150,width=500,height=150')"> <%=rs("width")%>*<%=rs("height")%>����</td> 
	<td  align="center" rowspan="3"><a href=<%=rs("pic1")%> target=_blank><img src=<%=rs("pic1")%> border=0  alt="���Ԥ��Ч��" width=180 height=110></a>
	</td></tr>	
	<tr><td width=20% align=right height="18">����</td><td> <input type=text value="<%=rs("pic1_lnk")%>" name=pic1_lnk size=40 maxlength=100></td></tr>		
	<tr><td width=20% align=right height="18">˵��</td><td> <input type=text value="<%=rs("tit1")%>" name=tit1 size=40 maxlength=100></td></tr>	
	<tr><td width=20% align=right height="18">ͼƬ����·�����ļ���</td><td> <input type=text value="<%=rs("pic2")%>" name=pic2 size=40 maxlength=100>
	<INPUT TYPE="button" value="����ϴ�ͼƬ"  onClick="window.open('../ads/DNJY_upload_ad.asp?ttup=slideb&fupname=g2&frmname=pic2','blank_','scrollbars=yes,resizable=no,top=40,left=150,width=500,height=150')"> <%=rs("width")%>*<%=rs("height")%>����
	</td>
	<td  align="center" rowspan="3"><a href=<%=rs("pic2")%> target=_blank><img src=<%=rs("pic2")%> border=0  alt="���Ԥ��Ч��" width=180 height=110></a>
	</td>	</tr>	
	<tr><td width=20% align=right height="18">����</td><td> <input type=text value="<%=rs("pic2_lnk")%>" name=pic2_lnk size=40 maxlength=100></td></tr>			
	<tr><td width=20% align=right height="18">˵��</td><td> <input type=text value="<%=rs("tit2")%>" name=tit2 size=40 maxlength=100></td></tr>	
	<tr><td width=20% align=right height="18">ͼƬ����·�����ļ���</td><td><input type=text value="<%=rs("pic3")%>" name=pic3 size=40 maxlength=100>
	<INPUT TYPE="button" value="����ϴ�ͼƬ"  onClick="window.open('../ads/DNJY_upload_ad.asp?ttup=slideb&fupname=g3&frmname=pic3','blank_','scrollbars=yes,resizable=no,top=40,left=150,width=500,height=150')"> <%=rs("width")%>*<%=rs("height")%>����
	</td>
	<td  align="center" rowspan="3"><a href=<%=rs("pic3")%> target=_blank><img src=<%=rs("pic3")%> border=0  alt="���Ԥ��Ч��" width=180 height=110></a>
	</td>	</tr>
	<tr><td width=20% align=right height="18">����</td><td> <input type=text value="<%=rs("pic3_lnk")%>" name=pic3_lnk size=40 maxlength=100></td></tr>
	<tr><td width=20% align=right height="18">˵��</td><td> <input type=text value="<%=rs("tit3")%>" name=tit3 size=40 maxlength=100></td></tr>		                
	<tr><td width=20% align=right height="18">ͼƬ�ĵ�·�����ļ���</td><td> <input type=text value="<%=rs("pic4")%>" name=pic4 size=40 maxlength=100>
	<INPUT TYPE="button" value="����ϴ�ͼƬ"  onClick="window.open('../ads/DNJY_upload_ad.asp?ttup=slideb&fupname=g4&frmname=pic4','blank_','scrollbars=yes,resizable=no,top=40,left=150,width=500,height=150')"> <%=rs("width")%>*<%=rs("height")%>����
	</td>
	<td  align="center" rowspan="3"><a href=<%=rs("pic4")%> target=_blank><img src=<%=rs("pic4")%> border=0  alt="���Ԥ��Ч��" width=180 height=110></a>
	</td>	</tr>
	<tr><td width=20% align=right height="18">����</td><td> <input type=text value="<%=rs("pic4_lnk")%>" name=pic4_lnk size=40 maxlength=100></td></tr>		                
	<tr><td width=20% align=right height="18">˵��</td><td> <input type=text value="<%=rs("tit4")%>" name=tit4 size=40 maxlength=100></td></tr>

	<tr><td width=20% align=right height="18">ͼƬ���·�����ļ���</td><td> <input type=text value="<%=rs("pic5")%>" name=pic5 size=40 maxlength=100>
	<INPUT TYPE="button" value="����ϴ�ͼƬ"  onClick="window.open('../ads/DNJY_upload_ad.asp?ttup=slideb&fupname=g5&frmname=pic5','blank_','scrollbars=yes,resizable=no,top=40,left=150,width=500,height=150')"> <%=rs("width")%>*<%=rs("height")%>����
	</td>
	<td  align="center" rowspan="3"><a href=<%=rs("pic5")%> target=_blank><img src=<%=rs("pic5")%> border=0  alt="���Ԥ��Ч��" width=180 height=110></a>
	</td>	</tr>
	<tr><td width=20% align=right height="18">����</td><td> <input type=text value="<%=rs("pic5_lnk")%>" name=pic5_lnk size=40 maxlength=100></td></tr>		                
	<tr><td width=20% align=right height="18">˵��</td><td> <input type=text value="<%=rs("tit5")%>" name=tit5 size=40 maxlength=100></td></tr>
	<tr><td width=20% align=right height="18"></td></tr>	
	<tr><td width=20% align=right height="18">ͼƬ���</td><td colspan=3> <input type=text value="<%=rs("width")%>" name=width size=10 maxlength=4> ���� ҪΪ��������</td></tr>
	<tr><td width=20% align=right height="18">ͼƬ�߶�</td><td colspan=3> <input type=text value="<%=rs("height")%>" name=height size=10 maxlength=4>  ���� ҪΪ��������</td></tr>
	<tr><td width=20% align=right height="18">�Ƿ���ʾ���ֱ�ǩ</td><td colspan=3> <input type=text value="<%=rs("show_text")%>" name=show_text size=10 maxlength=1 <%=onKeyUp%>>  ҪΪ����������1��ʾ 0����ʾ��</td></tr>
	<tr><td width=20% align=right height="18">���ֱ�ǩ��ɫ</td><td colspan=3> <input type=text value="<%=rs("txtcolor")%>" name=txtcolor size=10 maxlength=7 style="background:<%=rs("txtcolor")%>" onClick="Getcolor(this)">  ҪΪ��ɫRGB����</td></tr>
	<tr><td width=20% align=right height="18">���ֱ�ǩ��ɫ</td><td colspan=3> <input type=text value="<%=rs("bgcolor")%>" name=bgcolor size=10 maxlength=7 style="background:<%=rs("bgcolor")%>" onClick="Getcolor(this)">  ҪΪ��ɫRGB����</td></tr>
	<tr><td width=20% align=right height="18">ͼƬͣ��ʱ��</td><td colspan=3> <input type=text value="<%=rs("stop_time")%>" name=stop_time size=10 maxlength=4 <%=onKeyUp%>>  ҪΪ����������ÿ1000Ϊ1���ӣ�</td></tr>
	<tr><td width=20% align=right height="18">��ťλ��</td><td colspan=3> <input type=text value="<%=rs("button_pos")%>" name=button_pos size=10 maxlength=1 <%=onKeyUp%>> ҪΪ����������1�� 2�� 3�� 4�£�</td></tr>
	<tr><td colspan=3>
<INPUT name="ok" TYPE="hidden" value="ok"><INPUT name=action TYPE="submit" value="�������ò�����js����"></td></tr>
</form></table>
	<%end if%>

<%              
if action="ok" then

	if trim(request.form("pic1"))="" or trim(request.form("pic2"))="" or trim(request.form("pic3"))="" or trim(request.form("pic4"))="" or trim(request.form("pic5"))=""  then
	response.write "<script language='javascript'>"
	response.write "alert('���Ź���ͼƬ��·�����ļ���������д��������ȷ����������һҳ��');"
	response.write "location.href='javascript:history.go(-1)';"
	response.write "</script>"
	response.end
	end if

	pic=request.form("pic1")
	pic=split(pic,".")
	n=UBound(pic)
	if LCase(pic(n))<>"gif" and LCase(pic(n))<>"jpg" then
	response.write "<script language='javascript'>"
	response.write "alert('ͼƬһ�ĸ�ʽ����ȷ����֧��JPG��GIF��ʽ��ͼƬ��������ȷ����������һҳ��');"
	response.write "location.href='javascript:history.go(-1)';"
	response.write "</script>"
	response.end
	end if

	pic=request.form("pic2")
	pic=split(pic,".")
	n=UBound(pic)
	if LCase(pic(n))<>"gif" and LCase(pic(n))<>"jpg" then
	response.write "<script language='javascript'>"
	response.write "alert('ͼƬ���ĸ�ʽ����ȷ����֧��JPG��GIF��ʽ��ͼƬ��������ȷ����������һҳ��');"
	response.write "location.href='javascript:history.go(-1)';"
	response.write "</script>"
	response.end
	end if

	pic=request.form("pic3")
	pic=split(pic,".")
	n=UBound(pic)
	if LCase(pic(n))<>"gif" and LCase(pic(n))<>"jpg" then
	response.write "<script language='javascript'>"
	response.write "alert('ͼƬ���ĸ�ʽ����ȷ����֧��JPG��GIF��ʽ��ͼƬ��������ȷ����������һҳ��');"
	response.write "location.href='javascript:history.go(-1)';"
	response.write "</script>"
	response.end
	end if

	pic=request.form("pic4")
	pic=split(pic,".")
	n=UBound(pic)
	if LCase(pic(n))<>"gif" and LCase(pic(n))<>"jpg" then
	response.write "<script language='javascript'>"
	response.write "alert('ͼƬ�ĵĸ�ʽ����ȷ����֧��JPG��GIF��ʽ��ͼƬ��������ȷ����������һҳ��');"
	response.write "location.href='javascript:history.go(-1)';"
	response.write "</script>"
	response.end
	end if

	pic=request.form("pic5")
	pic=split(pic,".")
	n=UBound(pic)
	if LCase(pic(n))<>"gif" and LCase(pic(n))<>"jpg" then
	response.write "<script language='javascript'>"
	response.write "alert('ͼƬ��ĸ�ʽ����ȷ����֧��JPG��GIF��ʽ��ͼƬ��������ȷ����������һҳ��');"
	response.write "location.href='javascript:history.go(-1)';"
	response.write "</script>"
	response.end
	end if
	
     Set rs=Server.CreateObject("ADODB.Recordset")
	 sql="select * from DNJY_adhd"
	 rs.open sql,conn,1,3
 	 rs("pic1")=trim(request.form("pic1"))
 	 rs("pic1_lnk")=trim(request.form("pic1_lnk"))
 	 rs("tit1")=trim(request.form("tit1"))

 	 rs("pic2")=trim(request.form("pic2"))
 	 rs("pic2_lnk")=trim(request.form("pic2_lnk"))
 	 rs("tit2")=trim(request.form("tit2"))

 	 rs("pic3")=trim(request.form("pic3"))
 	 rs("pic3_lnk")=trim(request.form("pic3_lnk"))
 	 rs("tit3")=trim(request.form("tit3"))

 	 rs("pic4")=trim(request.form("pic4"))
 	 rs("pic4_lnk")=trim(request.form("pic4_lnk"))
 	 rs("tit4")=trim(request.form("tit4"))
 	 
 	 rs("pic5")=trim(request.form("pic5"))
 	 rs("pic5_lnk")=trim(request.form("pic5_lnk"))
 	 rs("tit5")=trim(request.form("tit5")) 
 	 	 
 	 rs("width")=trim(request.form("width"))
 	 rs("height")=trim(request.form("height")) 
 	 rs("show_text")=trim(request.form("show_text"))
 	 rs("txtcolor")=trim(request.form("txtcolor")) 
 	 rs("bgcolor")=trim(request.form("bgcolor"))
 	 rs("stop_time")=trim(request.form("stop_time")) 
 	 rs("button_pos")=trim(request.form("button_pos"))
 
	rs.update

'����js���뿪ʼ
Dim JsCode
JsCode="<script type=text/javascript>" & vbCrLf &_ 
"var pic_width="&rs("width")&"; //ͼƬ���" & vbCrLf &_ 
"var pic_height="&rs("height")&"; //ͼƬ�߶�" & vbCrLf &_ 
"var button_pos="&rs("button_pos")&"; //��Ťλ�� 1�� 2�� 3�� 4��" & vbCrLf &_ 
"var stop_time="&rs("stop_time")&"; //ͼƬͣ��ʱ��(1000Ϊ1����)" & vbCrLf &_ 
"var show_text="&rs("show_text")&"; //�Ƿ���ʾ���ֱ�ǩ 1��ʾ 0����ʾ" & vbCrLf &_ 
"var txtcolor="""&Mid(rs("txtcolor"),2)&"""; //����ɫ" & vbCrLf &_ 
"var bgcolor="""&Mid(rs("bgcolor"),2)&"""; //����ɫ" & vbCrLf &_ 
"var imag=new Array();" & vbCrLf &_ 
"var link=new Array();" & vbCrLf &_ 
"var text=new Array();" & vbCrLf &_ 
"imag[1]="""&rs("pic1")&""";" & vbCrLf &_ 
"link[1]="""&rs("pic1_lnk")&""";" & vbCrLf &_ 
"text[1]="""&rs("tit1")&""";" & vbCrLf &_ 
"imag[2]="""&rs("pic2")&""";" & vbCrLf &_ 
"link[2]="""&rs("pic2_lnk")&""";" & vbCrLf &_ 
"text[2]="""&rs("tit2")&""";" & vbCrLf &_ 
"imag[3]="""&rs("pic3")&""";" & vbCrLf &_ 
"link[3]="""&rs("pic3_lnk")&""";" & vbCrLf &_ 
"text[3]="""&rs("tit3")&""";" & vbCrLf &_ 
"imag[4]="""&rs("pic4")&""";" & vbCrLf &_ 
"link[4]="""&rs("pic4_lnk")&""";" & vbCrLf &_ 
"text[4]="""&rs("tit4")&""";" & vbCrLf &_ 
"imag[5]="""&rs("pic5")&""";" & vbCrLf &_ 
"link[5]="""&rs("pic5_lnk")&""";" & vbCrLf &_ 
"text[5]="""&rs("tit5")&""";" & vbCrLf &_ 
"//�ɱ༭���ݽ���" & vbCrLf &_ 
"var swf_height=show_text==1?pic_height+20:pic_height;" & vbCrLf &_ 
"var pics="""", links="""", texts="""";" & vbCrLf &_ 
"for(var i=1; i<imag.length; i++){" & vbCrLf &_ 
"	pics=pics+(""|""+imag[i]);" & vbCrLf &_ 
"	links=links+(""|""+link[i]);" & vbCrLf &_ 
"	texts=texts+(""|""+text[i]);" & vbCrLf &_ 
"}" & vbCrLf &_ 
"pics=pics.substring(1);" & vbCrLf &_ 
"links=links.substring(1);" & vbCrLf &_ 
"texts=texts.substring(1);" & vbCrLf &_ 
"document.write('<object classid=""clsid:d27cdb6e-ae6d-11cf-96b8-444553540000"" codebase=""http://fpdownload.macromedia.com/pub/shockwave/cabs/flash/swflash.cabversion=6,0,0,0"" width=""'+ pic_width +'"" height=""'+ swf_height +'"">');" & vbCrLf &_ 
"document.write('<param name=""movie"" value=""/"&strInstallDir&"images/focus.swf"">');" & vbCrLf &_ 
"document.write('<param name=""quality"" value=""high""><param name=""wmode"" value=""opaque"">');" & vbCrLf &_ 
"document.write('<param name=""FlashVars""  value=""pics='+pics+'&links='+links+'&texts='+texts+'&pic_width='+pic_width+'&pic_height='+pic_height+'&show_text='+show_text+'&txtcolor='+txtcolor+'&bgcolor='+bgcolor+'&button_pos='+button_pos+'&stop_time='+stop_time+'"">');" & vbCrLf &_ 
"document.write('<embed src=""/"&strInstallDir&"images/focus.swf"" FlashVars=""pics='+pics+'&links='+links+'&texts='+texts+'&pic_width='+pic_width+'&pic_height='+pic_height+'&show_text='+show_text+'&txtcolor='+txtcolor+'&bgcolor='+bgcolor+'&button_pos='+button_pos+'&stop_time='+stop_time+'"" quality=""high"" width=""'+ pic_width +'"" height=""'+ swf_height +'"" allowScriptAccess=""sameDomain"" type=""application/x-shockwave-flash"" pluginspage=""http://www.macromedia.com/go/getflashplayer"" />');" & vbCrLf &_ 
"document.write('</object>');" & vbCrLf &_ 
"</script>" & vbCrLf &_ 
""

JsCode=Html2Js(JsCode)
JsCode="document.write('"&JsCode&"')"
JsCode = "<!--" & vbCrLf & JsCode & vbCrLf &"//-->" 
dim fs,f
Set fs = CreateObject("Scripting.FileSystemObject")
Set f = fs.CreateTextFile(Server.MapPath("/")&"/"&""&strInstallDir&"UploadFiles/js/hd.js", True)
f.write(JsCode)
f.close
Set f = nothing
Set fs = Nothing
'����js�������
Call CloseRs(rs)
Call CloseDB(conn)
	response.write "<script language='javascript'>"
	response.write "alert('�õ�ͼƬ���óɹ���������JS���룡');"
	response.write "location.href='adv.asp';"
	response.write "</script>"
	response.end
end if%>

</body>
<TABLE align="center">
<TR>
	<TD><button onClick="document.location.reload()">ˢ�¿�Ч��</button>Ч��Ԥ����<br>
<script src="../UploadFiles/js/hd.js"></script>	
	
<%Call CloseRs(rs)
Call CloseDB(conn)%></TD>
	<TD>���÷�ʽ��&lt;script src="UploadFiles/js/hd.js">&lt;/script></TD>
</TR>
</TABLE>
<p></html>

