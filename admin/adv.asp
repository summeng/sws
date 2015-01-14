<!--#include file="../conn.asp"-->
<!--#include file="../SqlIn.Asp"-->
<!--#include file="../config.asp"-->
<!--#include file="../Include/Function.asp"-->
<!--#include file=cookies.asp-->
<!--#include file=../ads/Html2Js.asp-->
<%
'=====================================
'天天供求信息网站管理系统
'天天网络科技工作室出品
'客服中心http://www.ip126.com
'技术支持论坛 http://www.ip126.com/bbs
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>
<%call checkmanage("13")%>
<html>
<head>
<title>网站设置</title>
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
	
	<tr class=backs><td colspan=3 class=td height=18 align="center"><b>幻灯广告B设置（首页幻灯广告B不分地区，全局显示，在基本参数中可设置显示A或B）</b></td></tr>
	
	<tr><td width=20% align=right height="18">图片一的路径和文件名</td><td> <input type=text value="<%=rs("pic1")%>" name=pic1 size=40 maxlength=100>
	<INPUT TYPE="button" value="浏览上传图片"  onClick="window.open('../ads/DNJY_upload_ad.asp?ttup=slideb&fupname=g1&frmname=pic1','blank_','scrollbars=yes,resizable=no,top=40,left=150,width=500,height=150')"> <%=rs("width")%>*<%=rs("height")%>像素</td> 
	<td  align="center" rowspan="3"><a href=<%=rs("pic1")%> target=_blank><img src=<%=rs("pic1")%> border=0  alt="点击预览效果" width=180 height=110></a>
	</td></tr>	
	<tr><td width=20% align=right height="18">链接</td><td> <input type=text value="<%=rs("pic1_lnk")%>" name=pic1_lnk size=40 maxlength=100></td></tr>		
	<tr><td width=20% align=right height="18">说明</td><td> <input type=text value="<%=rs("tit1")%>" name=tit1 size=40 maxlength=100></td></tr>	
	<tr><td width=20% align=right height="18">图片二的路径和文件名</td><td> <input type=text value="<%=rs("pic2")%>" name=pic2 size=40 maxlength=100>
	<INPUT TYPE="button" value="浏览上传图片"  onClick="window.open('../ads/DNJY_upload_ad.asp?ttup=slideb&fupname=g2&frmname=pic2','blank_','scrollbars=yes,resizable=no,top=40,left=150,width=500,height=150')"> <%=rs("width")%>*<%=rs("height")%>像素
	</td>
	<td  align="center" rowspan="3"><a href=<%=rs("pic2")%> target=_blank><img src=<%=rs("pic2")%> border=0  alt="点击预览效果" width=180 height=110></a>
	</td>	</tr>	
	<tr><td width=20% align=right height="18">链接</td><td> <input type=text value="<%=rs("pic2_lnk")%>" name=pic2_lnk size=40 maxlength=100></td></tr>			
	<tr><td width=20% align=right height="18">说明</td><td> <input type=text value="<%=rs("tit2")%>" name=tit2 size=40 maxlength=100></td></tr>	
	<tr><td width=20% align=right height="18">图片三的路径和文件名</td><td><input type=text value="<%=rs("pic3")%>" name=pic3 size=40 maxlength=100>
	<INPUT TYPE="button" value="浏览上传图片"  onClick="window.open('../ads/DNJY_upload_ad.asp?ttup=slideb&fupname=g3&frmname=pic3','blank_','scrollbars=yes,resizable=no,top=40,left=150,width=500,height=150')"> <%=rs("width")%>*<%=rs("height")%>像素
	</td>
	<td  align="center" rowspan="3"><a href=<%=rs("pic3")%> target=_blank><img src=<%=rs("pic3")%> border=0  alt="点击预览效果" width=180 height=110></a>
	</td>	</tr>
	<tr><td width=20% align=right height="18">链接</td><td> <input type=text value="<%=rs("pic3_lnk")%>" name=pic3_lnk size=40 maxlength=100></td></tr>
	<tr><td width=20% align=right height="18">说明</td><td> <input type=text value="<%=rs("tit3")%>" name=tit3 size=40 maxlength=100></td></tr>		                
	<tr><td width=20% align=right height="18">图片四的路径和文件名</td><td> <input type=text value="<%=rs("pic4")%>" name=pic4 size=40 maxlength=100>
	<INPUT TYPE="button" value="浏览上传图片"  onClick="window.open('../ads/DNJY_upload_ad.asp?ttup=slideb&fupname=g4&frmname=pic4','blank_','scrollbars=yes,resizable=no,top=40,left=150,width=500,height=150')"> <%=rs("width")%>*<%=rs("height")%>像素
	</td>
	<td  align="center" rowspan="3"><a href=<%=rs("pic4")%> target=_blank><img src=<%=rs("pic4")%> border=0  alt="点击预览效果" width=180 height=110></a>
	</td>	</tr>
	<tr><td width=20% align=right height="18">链接</td><td> <input type=text value="<%=rs("pic4_lnk")%>" name=pic4_lnk size=40 maxlength=100></td></tr>		                
	<tr><td width=20% align=right height="18">说明</td><td> <input type=text value="<%=rs("tit4")%>" name=tit4 size=40 maxlength=100></td></tr>

	<tr><td width=20% align=right height="18">图片五的路径和文件名</td><td> <input type=text value="<%=rs("pic5")%>" name=pic5 size=40 maxlength=100>
	<INPUT TYPE="button" value="浏览上传图片"  onClick="window.open('../ads/DNJY_upload_ad.asp?ttup=slideb&fupname=g5&frmname=pic5','blank_','scrollbars=yes,resizable=no,top=40,left=150,width=500,height=150')"> <%=rs("width")%>*<%=rs("height")%>像素
	</td>
	<td  align="center" rowspan="3"><a href=<%=rs("pic5")%> target=_blank><img src=<%=rs("pic5")%> border=0  alt="点击预览效果" width=180 height=110></a>
	</td>	</tr>
	<tr><td width=20% align=right height="18">链接</td><td> <input type=text value="<%=rs("pic5_lnk")%>" name=pic5_lnk size=40 maxlength=100></td></tr>		                
	<tr><td width=20% align=right height="18">说明</td><td> <input type=text value="<%=rs("tit5")%>" name=tit5 size=40 maxlength=100></td></tr>
	<tr><td width=20% align=right height="18"></td></tr>	
	<tr><td width=20% align=right height="18">图片宽度</td><td colspan=3> <input type=text value="<%=rs("width")%>" name=width size=10 maxlength=4> 像素 要为数字整数</td></tr>
	<tr><td width=20% align=right height="18">图片高度</td><td colspan=3> <input type=text value="<%=rs("height")%>" name=height size=10 maxlength=4>  像素 要为数字整数</td></tr>
	<tr><td width=20% align=right height="18">是否显示文字标签</td><td colspan=3> <input type=text value="<%=rs("show_text")%>" name=show_text size=10 maxlength=1 <%=onKeyUp%>>  要为数字整数（1显示 0不显示）</td></tr>
	<tr><td width=20% align=right height="18">文字标签颜色</td><td colspan=3> <input type=text value="<%=rs("txtcolor")%>" name=txtcolor size=10 maxlength=7 style="background:<%=rs("txtcolor")%>" onClick="Getcolor(this)">  要为颜色RGB代码</td></tr>
	<tr><td width=20% align=right height="18">文字标签底色</td><td colspan=3> <input type=text value="<%=rs("bgcolor")%>" name=bgcolor size=10 maxlength=7 style="background:<%=rs("bgcolor")%>" onClick="Getcolor(this)">  要为颜色RGB代码</td></tr>
	<tr><td width=20% align=right height="18">图片停留时间</td><td colspan=3> <input type=text value="<%=rs("stop_time")%>" name=stop_time size=10 maxlength=4 <%=onKeyUp%>>  要为数字整数（每1000为1秒钟）</td></tr>
	<tr><td width=20% align=right height="18">按钮位置</td><td colspan=3> <input type=text value="<%=rs("button_pos")%>" name=button_pos size=10 maxlength=1 <%=onKeyUp%>> 要为数字整数（1左 2右 3上 4下）</td></tr>
	<tr><td colspan=3>
<INPUT name="ok" TYPE="hidden" value="ok"><INPUT name=action TYPE="submit" value="保存设置并生成js代码"></td></tr>
</form></table>
	<%end if%>

<%              
if action="ok" then

	if trim(request.form("pic1"))="" or trim(request.form("pic2"))="" or trim(request.form("pic3"))="" or trim(request.form("pic4"))="" or trim(request.form("pic5"))=""  then
	response.write "<script language='javascript'>"
	response.write "alert('五张滚动图片的路径和文件名必须填写，单击“确定”返回上一页！');"
	response.write "location.href='javascript:history.go(-1)';"
	response.write "</script>"
	response.end
	end if

	pic=request.form("pic1")
	pic=split(pic,".")
	n=UBound(pic)
	if LCase(pic(n))<>"gif" and LCase(pic(n))<>"jpg" then
	response.write "<script language='javascript'>"
	response.write "alert('图片一的格式不正确，仅支持JPG、GIF格式的图片，单击“确定”返回上一页！');"
	response.write "location.href='javascript:history.go(-1)';"
	response.write "</script>"
	response.end
	end if

	pic=request.form("pic2")
	pic=split(pic,".")
	n=UBound(pic)
	if LCase(pic(n))<>"gif" and LCase(pic(n))<>"jpg" then
	response.write "<script language='javascript'>"
	response.write "alert('图片二的格式不正确，仅支持JPG、GIF格式的图片，单击“确定”返回上一页！');"
	response.write "location.href='javascript:history.go(-1)';"
	response.write "</script>"
	response.end
	end if

	pic=request.form("pic3")
	pic=split(pic,".")
	n=UBound(pic)
	if LCase(pic(n))<>"gif" and LCase(pic(n))<>"jpg" then
	response.write "<script language='javascript'>"
	response.write "alert('图片三的格式不正确，仅支持JPG、GIF格式的图片，单击“确定”返回上一页！');"
	response.write "location.href='javascript:history.go(-1)';"
	response.write "</script>"
	response.end
	end if

	pic=request.form("pic4")
	pic=split(pic,".")
	n=UBound(pic)
	if LCase(pic(n))<>"gif" and LCase(pic(n))<>"jpg" then
	response.write "<script language='javascript'>"
	response.write "alert('图片四的格式不正确，仅支持JPG、GIF格式的图片，单击“确定”返回上一页！');"
	response.write "location.href='javascript:history.go(-1)';"
	response.write "</script>"
	response.end
	end if

	pic=request.form("pic5")
	pic=split(pic,".")
	n=UBound(pic)
	if LCase(pic(n))<>"gif" and LCase(pic(n))<>"jpg" then
	response.write "<script language='javascript'>"
	response.write "alert('图片五的格式不正确，仅支持JPG、GIF格式的图片，单击“确定”返回上一页！');"
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

'生成js代码开始
Dim JsCode
JsCode="<script type=text/javascript>" & vbCrLf &_ 
"var pic_width="&rs("width")&"; //图片宽度" & vbCrLf &_ 
"var pic_height="&rs("height")&"; //图片高度" & vbCrLf &_ 
"var button_pos="&rs("button_pos")&"; //按扭位置 1左 2右 3上 4下" & vbCrLf &_ 
"var stop_time="&rs("stop_time")&"; //图片停留时间(1000为1秒钟)" & vbCrLf &_ 
"var show_text="&rs("show_text")&"; //是否显示文字标签 1显示 0不显示" & vbCrLf &_ 
"var txtcolor="""&Mid(rs("txtcolor"),2)&"""; //文字色" & vbCrLf &_ 
"var bgcolor="""&Mid(rs("bgcolor"),2)&"""; //背景色" & vbCrLf &_ 
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
"//可编辑内容结束" & vbCrLf &_ 
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
'生成js代码结束
Call CloseRs(rs)
Call CloseDB(conn)
	response.write "<script language='javascript'>"
	response.write "alert('幻灯图片设置成功并生成了JS代码！');"
	response.write "location.href='adv.asp';"
	response.write "</script>"
	response.end
end if%>

</body>
<TABLE align="center">
<TR>
	<TD><button onClick="document.location.reload()">刷新看效果</button>效果预览：<br>
<script src="../UploadFiles/js/hd.js"></script>	
	
<%Call CloseRs(rs)
Call CloseDB(conn)%></TD>
	<TD>调用方式：&lt;script src="UploadFiles/js/hd.js">&lt;/script></TD>
</TR>
</TABLE>
<p></html>

