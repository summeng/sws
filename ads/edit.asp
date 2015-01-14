<!--#include file="../conn.asp"-->
<!--#include file="../config.asp"-->
<!--#include file=cookies.asp-->
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
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>COAD System</title>
<link href="../css/style.css" rel="stylesheet" type="text/css">
<script language=javascript src=../Include/mouse_on_title.js></script>
<SCRIPT LANGUAGE="JavaScript" src="images/js.js"></SCRIPT>
<SCRIPT LANGUAGE="JavaScript" src="images/Time_selector.js"></SCRIPT>
<script language="JavaScript" type="text/JavaScript">
//网站logo上传
function JM_wu(ob){
 ob.style.display="none";
}
function JM_you(ob){
 ob.style.display="";
}
function uppic(model,frmname) {
//确定上传文件名，可选用下面与表单结合判断编号并以编号命名
id=document.form.ADID.value;
id1=document.form.city_one.value;
id2=document.form.city_two.value;
id3=document.form.city_three.value;
if (id1=='') id1=0
if (id2=='') id2=0
if (id3=='') id3=0
if (id=='') {
 alert("请先填写广告ID号！");
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
	Response.Write ("<script>alert(' 操作错误! \n\n 本地区广告ID重复,请使用其他ID !');history.back();</script>")
	Call CloseRs(rs)
	Call CloseDB(conn)
	Response.end
End If
Call CloseRs(rs)
'打开/更新广告数据
Set rs = Server.CreateObject("ADODB.Recordset")
sql="select ADID,ADType,ADSrc,ADHeight,ADWidth,adsjs,ADLink,ADAlt,ADStopViews,ADStopHits,ADStopDate,ADNote,ADViews,ADHits,ADjs,ADkg,city_oneid,city_twoid,city_threeid,Advertisement_host from [DNJY_ad] where id=" & id
rs.Open sql,conn,1,3
'是否更新
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
	
	Response.Write "<script language=javascript>setTimeout(""alert('修改广告成功！');" &"location.replace('createjs.asp?page="&request("page")&"&ADID="&rs("ADID")&"&city_oneid="&rs("city_oneid")&"&city_twoid="&rs("city_twoid")&"&city_threeid="&rs("city_threeid")&"')"",0)</script>"

End If
If Err <> 0 Then
			Response.Write "<font color=red size=2>错误:"&Err.Description
			Response.End
End If
%>
<table width="100%" border="0" align="center" cellpadding="3" cellspacing="1" bgcolor="#6699cc">
  <form name=form action=edit.asp?id=<%=id%>  method=post onsubmit="return chkinput()">
    <tr align="center" bgcolor="#6699cc">
      <td height="28" colspan="2"><div align="center" ><font color="#FFFFFF"><b>修改广告</b></font></div></td>
    </tr>
    <tr bgcolor="#FFFFFF">
      <td width="15%" align="right" bgcolor="#FAFAFA"><div align="center">广告&nbsp;I&nbsp;D：</div></td>
      <td width="400">
      <INPUT   maxLength=14 size=20 name="dddd" value="<%=rs("ADID")%>" DISABLED> <font color="#FF0000">*</font></td>
    </tr><input type="hidden" name="ADID" value="<%=rs("ADID")%>">
    <tr bgcolor="#FFFFFF">
      <td width="15%" align="right" bgcolor="#FAFAFA"><div align="center">所属地区：</div></td>
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
	document.form.city_two.options[0] = new Option('选择城市','');
	document.form.city_three.length = 0; 
	document.form.city_three.options[0] = new Option('选择城市','');
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
    document.form.city_three.options[0] = new Option('选择城市','');
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
                                     <OPTION value="" selected>选择城市</OPTION>
                                     <%set rsi=conn.execute("select * from DNJY_city where id>0 and twoid=0 order by indexid asc")
if rsi.eof and rsi.bof then
response.write "<option value=''>没有分类</option>"
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
                                     <OPTION value="" selected>选择城市</OPTION>
                                     <%set rsi=conn.execute("select * from DNJY_city where id="&rs("city_oneid")&" and twoid<>0 and threeid=0")
if rsi.eof and rsi.bof then
response.write "<option value=''>没有分类</option>"
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
                                     <OPTION value="" selected>选择城市</OPTION>
                                     <%set rsi=conn.execute("select * from DNJY_city where id="&rs("city_oneid")&" and twoid="&rs("city_twoid")&" and threeid<>0")
if rsi.eof and rsi.bof then
response.write "<option value=''>没有分类</option>"
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
      <td width="15%" height="25" align="right" bgcolor="#FaFaFa"><div align="center">广告类型：</div></td>
      <td><select name="ADType" size="1" class="input1" onChange="ChangeType(this.options[this.selectedIndex].value)">
          <option <%If rs("ADType")=1 Then Response.Write"selected"%> value="1">普通图片</option>
          <option <%If rs("ADType")=2 Then Response.Write"selected"%> value="2">满屏浮动显示</option>
          <option <%If rs("ADType")=3 Then Response.Write"selected"%> value="3">上下浮动显示 - 右</option>
          <option <%If rs("ADType")=4 Then Response.Write"selected"%> value="4">上下浮动显示 - 左</option>
          <option <%If rs("ADType")=5 Then Response.Write"selected"%> value="5">全屏幕渐隐消失</option>
          <option <%If rs("ADType")=6 Then Response.Write"selected"%> value="6">普通网页对话框 </option>
          <option <%If rs("ADType")=7 Then Response.Write"selected"%> value="7">可移动透明对话框 </option>
          <option <%If rs("ADType")=8 Then Response.Write"selected"%> value="8">打开新窗口</option>
          <option <%If rs("ADType")=9 Then Response.Write"selected"%> value="9">弹出新窗口</option>
          <option <%If rs("ADType")=10 Then Response.Write"selected"%> value="10">对联式广告（固定）</option>
          <option <%If rs("ADType")=11 Then Response.Write"selected"%> value="11">对联式广告（浮动）</option>
		  <option <%If rs("ADType")=12 Then Response.Write"selected"%> value="12">走马灯文字</option>  
		  <option <%If rs("ADType")=13 Then Response.Write"selected"%> value="13">普通文字广告</option>
          <option style="background-color:#F8F4F5 ;color: #FF0000" <%If rs("ADType")=14 Then Response.Write"selected"%> value="14">自写代码广告</option>
		  <option <%If rs("ADType")=15 Then Response.Write"selected"%> value="15">Flash图片广告</option>
		  <option <%If rs("ADType")=16 Then Response.Write"selected"%> value="16">视频广告</option>
      </select> 若选择“满屏浮动显示”会造成切换城市列表无法弹出！</td>
    </tr>
    <tr bgcolor="#FFFFFF">
      <td width="15%" height="25" align="right" bgcolor="#FaFaFa"><div align="center">图片地址：</div></td>
      <td><%If rs("ADType")= 6 Then
    	Response.Write ("<textarea rows=3 name=ADSrc cols=27 class=input2>"&server.HTMLencode(rs("ADSrc"))&"</textarea>")
	else%>
          <INPUT name="ADSrc" type="text" class="input1" value="<%=rs("ADSrc")%>" size=30>
          <%end if%> <INPUT alt="请单击“浏览”上传图片<br>或填写图片的网址，如<font color=blue>http://www.ip126.com/adv/mp3.jpg</font><br>或填写站内的图片路径、文件名，如<font color=blue>adv/001.jpg</font>" TYPE="button" value="浏览图片上传" onclick="javascript:uppic('','ADSrc');">
      *<a href="#" onclick=openhelp("ext") title="点击查看帮助"><font color="#0000FF">上传普通图片、FLASH图片</font></a>，或填写视频文件地址 <a href="javascript:void(0)" onclick = "document.getElementById('light').style.display='block';document.getElementById('fade').style.display='block'" title="点此查看视频格式要求"><font color="#FF0000">[视频广告添加说明]</font></a></td>
    </tr>


<!--弹出页面内容开始-->
<div id="light" class="white_content">
<div class="white_content_top"><span class="white_content_top_span1">视频广告添加说明</span><span class="white_content_top_span2"><a href="javascript:void(0)" onclick = "document.getElementById('light').style.display='none';document.getElementById('fade').style.display='none'">关闭</a></span></div><span class="fRed">&nbsp;&nbsp;&nbsp;&nbsp;1、视频文件地址添加有两种方法：一是相对路径，视频文件放在本服务器，地址形式如/ads/pic/abc.wmv；二是绝对路径，视频文件可放在本服务器，也可远程，地址形式为http://www.abc.com/video/abc.wmv。<br>&nbsp;&nbsp;&nbsp;&nbsp;2、视频一般较大，如果放在本服务器不宜直接上传，可用FTP上传到/ads/pic目录下。<br>&nbsp;&nbsp;&nbsp;&nbsp;3、视频广告要求电脑必须安装Windows Media Player 9.0以上版本播放器才能正常播放（一般xp及以上版本操作系统都具备）。<br>&nbsp;&nbsp;&nbsp;&nbsp;4、视频格式兼容：wmv，mpg，avi,rmvb等，建议wmv格式文件体积小些。</div>
<div id="fade" class="black_overlay"></div>
<!--弹出页面内容结束-->

    <tr bgcolor="#FFFFFF">
      <td width="15%" height="25" align="right" bgcolor="#FaFaFa"><div align="center">图片规格：</div></td>
      <td><INPUT name=ADWidth type="text" class="input1" onkeypress="return Num();" value="<%=rs("ADWidth")%>" size=8 maxlength=4>
        ×
      <INPUT name=ADHeight type="text" class="input1" onkeypress="return Num();" value="<%=rs("ADHeight")%>" size=8 maxlength=4> 图片(Flash)宽高|弹窗宽高|视频广告宽高|前一框填走马灯类型广告文字滚动速度，一般填5</td>
    </tr>
    <tr bgcolor="#FFFFFF" style="visibility:hide;">
      <td width="15%" height="25" align="right" bgcolor="#FaFaFa"><div align="center">链接地址：</div></td>
      <td><INPUT name=ADLink type="text" class="input1" value="<%=rs("ADLink")%>" size=30 maxlength=150> 链接地址|弹窗链接地址|走马灯文字链接地址</td>
    </tr>
    <tr bgcolor="#FFFFFF">
      <td width="15%" height="25" align="right" bgcolor="#FaFaFa"><div align="center">提示文字：</div></td>
      <td><INPUT name=ADAlt type="text" class="input1" value="<%=rs("ADAlt")%>" size=30 maxlength=50> 鼠标悬停提示|走马灯类型广告文字</td>
    </tr>
    <tr bgcolor="#FFFFFF">
      <td width="15%" height="25" align="right" bgcolor="#FaFaFa"><div align="center">自写代码：</div></td>
      <td><textarea name="ADjs" cols="70" rows="5" id="ADjs"><%=rs("ADjs")%></textarea></td>
    </tr>
    <tr bgcolor="#FFFFFF">
      <td width="15%" height="25" align="right" bgcolor="#FaFaFa"><div align="center">投放限制：</div></td>
      <td><INPUT name=ADStopViews type="text" class="input1" onkeypress="return Num();" value="<%=rs("ADStopViews")%>" size=8 maxlength=10>
        ·
          <INPUT name=ADStopHits type="text" class="input1" onkeypress="return Num();" value="<%=rs("ADStopHits")%>" size=8 maxlength=10>
        ·
        <INPUT name=ADStopDate type="text" class="input1" value="<%=rs("ADStopDate")%>" size=12 maxlength=19 onclick="setday(this)">
      <a href="#" onclick=openhelp("stop") title="点击查看帮助">显示·点击·日期？</a></td>
    </tr>
    <tr bgcolor="#FFFFFF">
      <td width="15%" height="25" align="right" bgcolor="#FaFaFa"><div align="center">重新统计：</div></td>
      <td><input type="checkbox" name="ADRESET" value="YES">
      重置显示和点击次数</td>
    </tr>
    <tr bgcolor="#FFFFFF">
      <td width="15%" height="25" align="right" bgcolor="#FaFaFa"><div align="center">简单注释：</div></td>
      <td><INPUT name=ADNote type="text" class="input1" value="<%=rs("ADNote")%>" size=30 maxlength=100>
      备注不显示在广告中</td>
    <tr bgcolor="#FFFFFF">
      <td width="15%" height="25" align="right" bgcolor="#FaFaFa"><div align="center">是否活动：</div></td>
      <td><%if rs("adkg")=1 then%>               
                <input type="radio" name="adkg" value="1" id="adkg" checked>
                 活动&nbsp;&nbsp;&nbsp;&nbsp;
                <input type="radio" name="adkg" value="0" id="adkg">
                暂停</td>
                <%else%>                         
                <input type="radio" name="adkg" value="1" id="adkg">
                 活动&nbsp;&nbsp;&nbsp;&nbsp;
                <input type="radio" name="adkg" value="0" id="adkg" checked>
                暂停<%end if%></td>
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
      <td width="15%" align="right" bgcolor="#FAFAFA"><div align="center">广告主信息：</div></td>
      <td height="17">姓名<INPUT name=ADhost1 value="<%=ADhost1%>" type="text" class="input1" size=10 maxlength=30>&nbsp;&nbsp;电话<INPUT name=ADhost2  value="<%=ADhost2%>" type="text" class="input1" size=20 maxlength=30>&nbsp;&nbsp;手机<INPUT name=ADhost3 value="<%=ADhost3%>"  type="text" class="input1" size=20 maxlength=30>&nbsp;&nbsp;腾讯QQ<INPUT name=ADhost4 value="<%=ADhost4%>"  type="text" class="input1" size=20 maxlength=30>&nbsp;&nbsp;邮箱<INPUT name=ADhost5 value="<%=ADhost5%>"  type="text" class="input1" size=30 maxlength=30><br>地址<INPUT name=ADhost6 value="<%=ADhost6%>"  type="text" class="input1" size=38 maxlength=30>&nbsp;&nbsp;邮编<INPUT name=ADhost7 value="<%=ADhost7%>"  type="text" class="input1" size=20 maxlength=30>&nbsp;&nbsp;广告费<INPUT name=ADhost8 value="<%=ADhost8%>"  type="text" class="input1" size=17 maxlength=30> 元&nbsp;&nbsp;备注<INPUT name=ADhost9 value="<%=ADhost9%>"  type="text" class="input1" size=30 maxlength=30></td>
    </tr>	
    <tr bgcolor="#eeeeee">
      <td height="35" colspan="2" align="center">        <table width="100%"  border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td><div align="center">
			<input type="hidden" name="page" value="<%=request("page")%>">
              <input name="ChangeAD" type="submit" class="input1" value="确定修改">
            </div></td>
            <td><div align="center">
              <input name="ChangeAD" type="reset" class="input1" value="还原设置">
            </div></td>
          </tr>
        </table>
		
      </td>
    </tr>
  </form>
</table>
<CENTER><button onClick="document.location.reload()">刷新看效果</button> 广告效果显示，如果不显示广告请查看是否为过期或暂停！<br>
<%if rs("ADType")=2 Or rs("ADType")=3 Or rs("ADType")=4 Or rs("ADType")=5 Or rs("ADType")=6 Or rs("ADType")=7 Or rs("ADType")=8 Or rs("ADType")=9 Or rs("ADType")=10 Or rs("ADType")=11 Or rs("ADType")=12 Or rs("ADType")=13 Or rs("ADType")=14 Or rs("ADType")=15 then%>
<script src="/<%=strInstallDir%>ads/JS/<%=rs("ADID")%>_<%=rs("city_oneid")%>_<%=rs("city_twoid")%>_<%=rs("city_threeid")%>.js"></script>
<%ElseIf rs("ADType")=16 then
Response.Write "视频广告不预览"
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
	Response.Write "没有指明要编辑的ID。"
End If

Rem 过滤可能出错误的符号
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
	alert("请输入用户名！");
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