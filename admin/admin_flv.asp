<!--#include file="../conn.asp"-->
<!--#include file="../SqlIn.Asp"-->
<!--#include file="../config.asp"-->
<!--#include file=cookies.asp-->
<!--#include file=../ads/Html2Js.asp-->
<%
'=====================================
'天天供求信息网站管理系统
'天天网络科技工作室出品
'客服中心http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>
<%call checkmanage("13")%>
<html>
<head>
<title>FLV视频广告设置</title>
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
	
	<tr class=backs><td colspan=3 class=td height=18 align="center">FLV视频广告设置</td></tr>

	<tr><td width=20% align=right height="18">视频一标题</td><td> <input type=text value="<%=flvconfig0%>" name=flvconfig0 size=40 maxlength=100></td></tr>
	<tr><td width=20% align=right height="18">视频一文件地址</td><td> <input type=text value="<%=flvconfig1%>" name=flvconfig1 size=40 maxlength=100>
	 <INPUT TYPE="button" value="上传FLV视频文件一"  onClick="window.open('../ads/DNJY_upload_ad.asp?ttup=flv&fupname=flv1&frmname=flvconfig1','blank_','scrollbars=yes,resizable=no,top=40,left=150,width=500,height=150')"> <br><FONT color=#AB3801>样式：绝对路径如http://www.ip126.com/UploadFiles/flv/flv1.flv。或相对路径（适合根目录下安装程序）/UploadFiles/flv/flv1.flv</font></td></tr>		
	
	<tr><td width=20% align=right height="18">视频二标题</td><td> <input type=text value="<%=flvconfig2%>" name=flvconfig2 size=40 maxlength=100></td></tr>
	<tr><td width=20% align=right height="18">视频二文件地址</td><td> <input type=text value="<%=flvconfig3%>" name=flvconfig3 size=40 maxlength=100>
	  <INPUT TYPE="button" value="上传FLV视频文件二"  onClick="window.open('../ads/DNJY_upload_ad.asp?ttup=flv&fupname=flv2&frmname=flvconfig3','blank_','scrollbars=yes,resizable=no,top=40,left=150,width=500,height=150')"> <br><FONT color=#AB3801>样式：绝对路径如http://www.ip126.com/UploadFiles/flv/flv2.flv。或相对路径（适合根目录下安装程序）/UploadFiles/flv/flv2.flv</font></td></tr>
	
	<tr><td width=20% align=right height="18">视频三标题</td><td> <input type=text value="<%=flvconfig4%>" name=flvconfig4 size=40 maxlength=100></td></tr>
	<tr><td width=20% align=right height="18">视频三文件地址</td><td> <input type=text value="<%=flvconfig5%>" name=flvconfig5 size=40 maxlength=100>
	  <INPUT TYPE="button" value="上传FLV视频文件三"  onClick="window.open('../ads/DNJY_upload_ad.asp?ttup=flv&fupname=flv3&frmname=flvconfig5','blank_','scrollbars=yes,resizable=no,top=40,left=150,width=500,height=150')"> <br><FONT color=#AB3801>样式：绝对路径如http://www.ip126.com/UploadFiles/flv/flv3.flv。或相对路径（适合根目录下安装程序）/UploadFiles/flv/flv3.flv</font></td></tr>

	<tr><td width=20% align=right height="18">视频四标题</td><td> <input type=text value="<%=flvconfig6%>" name=flvconfig6 size=40 maxlength=100></td></tr>
	<tr><td width=20% align=right height="18">视频四文件地址</td><td> <input type=text value="<%=flvconfig7%>" name=flvconfig7 size=40 maxlength=100>
	  <INPUT TYPE="button" value="上传FLV视频文件四"  onClick="window.open('../ads/DNJY_upload_ad.asp?ttup=flv&fupname=flv4&frmname=flvconfig7','blank_','scrollbars=yes,resizable=no,top=40,left=150,width=500,height=150')"> <br><FONT color=#AB3801>样式：绝对路径如http://www.ip126.com/UploadFiles/flv/flv4.flv。或相对路径（适合根目录下安装程序）/UploadFiles/flv/flv4.flv</font></td></tr>	 

	<tr><td width=20% align=right height="18">视频五标题</td><td> <input type=text value="<%=flvconfig8%>" name=flvconfig8 size=40 maxlength=100></td></tr>
	<tr><td width=20% align=right height="18">视频五文件地址</td><td> <input type=text value="<%=flvconfig9%>" name=flvconfig9 size=40 maxlength=100>
	 <INPUT TYPE="button" value="上传FLV视频文件五"  onClick="window.open('../ads/DNJY_upload_ad.asp?ttup=flv&fupname=flv5&frmname=flvconfig9','blank_','scrollbars=yes,resizable=no,top=40,left=150,width=500,height=150')"> <br><FONT color=#AB3801>样式：绝对路径如http://www.ip126.com/UploadFiles/flv/flv5.flv。或相对路径（适合根目录下安装程序）/UploadFiles/flv/flv5.flv</font></td></tr>	

	<tr><td width=20% align=right height="18"></td></tr>	
	<tr><td width=20% align=right height="18">参数设置</td><td colspan=3>
播放器宽度 <input type=text value="<%=flvconfig10%>" name=flvconfig10 size=4 maxlength=3>像素<br>
播放器高度 <input type=text value="<%=flvconfig11%>" name=flvconfig11 size=4 maxlength=3>像素<br>
是否自动播放<input type="radio" value="1" <%if flvconfig12="1" then%>checked<%end if%> name="flvconfig12">自动
<input type="radio" value="0" <%if flvconfig12="0" then%>checked<%end if%> name="flvconfig12">不自动<br>
是否连续播放<input type="radio" value="1" <%if flvconfig13="1" then%>checked<%end if%> name="flvconfig13">连续
<input type="radio" value="0" <%if flvconfig13="0" then%>checked<%end if%> name="flvconfig13">不连续 <br>
是否随机播放<input type="radio" value="1" <%if flvconfig14="1" then%>checked<%end if%> name="flvconfig14">随机
<input type="radio" value="0" <%if flvconfig14="0" then%>checked<%end if%> name="flvconfig14">顺序<br>
是否显示时间<input type="radio" value="1" <%if flvconfig15="1" then%>checked<%end if%> name="flvconfig15">显示
<input type="radio" value="0" <%if flvconfig15="0" then%>checked<%end if%> name="flvconfig15">不显示<br>
控制栏显示&nbsp;&nbsp;<input type="radio" value="1" <%if flvconfig16="1" then%>checked<%end if%> name="flvconfig16">一直显示 <input type="radio" value="2" <%if flvconfig16="2" then%>checked<%end if%> name="flvconfig16">鼠标悬停时显示 <input type="radio" value="3" <%if flvconfig16="3" then%>checked<%end if%> name="flvconfig16">鼠标悬停后显示 <input type="radio" value="0" <%if flvconfig16="0" then%>checked<%end if%> name="flvconfig16">不显示<br>
播放控制栏颜色 <input type=text value="<%=flvconfig17%>" name=flvconfig17 size=8 maxlength=6 style="background:<%=flvconfig17%>" onClick="Getcolor(this)">(16进制数字表示 例：000033)<br>
播放器文字颜色 <input type=text value="<%=flvconfig18%>" name=flvconfig18 size=8 maxlength=6 style="background:<%=flvconfig18%>" onClick="Getcolor(this)">(16进制数字表示 例：ffffff)<br>
按键图标颜色&nbsp;&nbsp; <input type=text value="<%=flvconfig19%>" name=flvconfig19 size=8 maxlength=6 style="background:<%=flvconfig19%>" onClick="Getcolor(this)">(16进制数字表示 例：66ff00)<br>
鼠标悬停光晕色 <input type=text value="<%=flvconfig20%>" name=flvconfig20 size=8 maxlength=6 style="background:<%=flvconfig20%>" onClick="Getcolor(this)">(16进制数字表示 例：FFFFFF)<br>
影片缓冲时间&nbsp;&nbsp; <input type=text value="<%=flvconfig21%>" name=flvconfig21 size=8 maxlength=6>(单位:秒)<br>
默认音量参数&nbsp;&nbsp; <input type=text value="<%=flvconfig22%>" name=flvconfig22 size=8 maxlength=6>(0-100 的数值)<br>

文字LOGO地址 &nbsp;&nbsp;<input type=text value="<%=flvconfig23%>" name=flvconfig23 size=40 maxlength=100>(英文)<br>
图片LOGO地址 &nbsp;&nbsp;<input type=text value="<%=flvconfig24%>" name=flvconfig24 size=40 maxlength=100>(填写LOGO地址,支持图片格式和swf格式)<br>
<!--播放前swf文件地址<input type=text value="<%=flvconfig25%>" name=flvconfig25 size=40 maxlength=100>(影片开始播放前从外部读取swf文件作广告等)<br>
播放后swf文件地址<input type=text value="<%=flvconfig26%>" name=flvconfig26 size=40 maxlength=100>(影片播放结束后从外部读取swf文件作影片信息等)-->
</td></tr>
	
	<tr><td colspan=3>
<INPUT name="ok" TYPE="hidden" value="ok"><INPUT name=action TYPE="submit" value="保存设置并生成JS文件"></td></tr>
</form></table>
	<%end if%>

<%              
if action="ok" then

	if trim(request.form("flvconfig0"))="" or trim(request.form("flvconfig1"))=""  then
	response.write "<script language='javascript'>"
	response.write "alert('第一个视频广告一定要填写，单击“确定”返回上一页！');"
	response.write "location.href='javascript:history.go(-1)';"
	response.write "</script>"
	response.end
	end if

	if trim(request.form("flvconfig2"))<>"" and trim(request.form("flvconfig3"))=""  then
	response.write "<script language='javascript'>"
	response.write "alert('视频二标题必须与视频地址对应，单击“确定”返回上一页！');"
	response.write "location.href='javascript:history.go(-1)';"
	response.write "</script>"
	response.end
	end if

	if trim(request.form("flvconfig4"))<>"" and trim(request.form("flvconfig5"))=""  then
	response.write "<script language='javascript'>"
	response.write "alert('视频三标题必须与视频地址对应，单击“确定”返回上一页！');"
	response.write "location.href='javascript:history.go(-1)';"
	response.write "</script>"
	response.end
	end if

	if trim(request.form("flvconfig6"))<>"" and trim(request.form("flvconfig7"))=""  then
	response.write "<script language='javascript'>"
	response.write "alert('视频四标题必须与视频地址对应，单击“确定”返回上一页！');"
	response.write "location.href='javascript:history.go(-1)';"
	response.write "</script>"
	response.end
	end if

	if trim(request.form("flvconfig8"))<>"" and trim(request.form("flvconfig9"))=""  then
	response.write "<script language='javascript'>"
	response.write "alert('视频五标题必须与视频地址对应，单击“确定”返回上一页！');"
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

'获取参数
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

'生成js代码
JsCode="<SCRIPT type=text/javascript>" & vbCrLf & _
"var swf_width="""&flvconfig10 &"""//播放器宽度" & vbCrLf & _
"var swf_height="&flvconfig11 &"//播放器高度" & vbCrLf & _
"var texts='"&texts&"'//广告标题参数，多个使用|分开,个数与filesurl配合一致" & vbCrLf & _
"var filesurl='"&filesurl&"'//flv文件绝对地址参数，多个使用|分开,个数与texts配合一致" & vbCrLf & _
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
	response.write "alert('FLV广告设置成功并生成js代码！');"
	response.write "location.href='admin_flv.asp';"
	response.write "</script>"
	response.end
end if%>

</body>
<TABLE align="center">
<TR>
	<TD><button onClick="document.location.reload()">刷新看效果</button>效果预览 :<br>注：1、服务器必须具备播放.flv视频条件才能播放，若不能播放要做设置，具体见<a href="http://www.ip126.com/Article/fwq/940.html" target="_blank"><FONT color=#0000ff>Flv流式视频的解决方案</font></a>。<br>2、视频地址和标题要与实际相符，如果没有就留空，否则出错。<br>
<span class="f14b" >
<script src="../UploadFiles/js/flv.js"></script>


<%Call CloseRs(rs)
Call CloseDB(conn)%></TD>
	<TD>调用方式：&lt;script src="UploadFiles/js/flv.js">&lt;/script></TD>
</TR>
</TABLE><br>&nbsp;<p>
</html>

