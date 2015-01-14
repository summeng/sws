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
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>
<SCRIPT src="../Include/sinaflash.js" type=text/javascript></SCRIPT>
<SCRIPT src="../Include/sinaflashfp.js" type=text/javascript></SCRIPT>

<%call checkmanage("13")%>
<%dim action,id,dname,dimgurl,adlink,i
id=request.QueryString("id")
action=request.querystring("action")
dimgurl=trim(request("imgurl"&request("i")&""))
select case action
case "add" 
set rs=server.CreateObject("adodb.recordset")
rs.open "select * from DNJY_adfp",conn,1,3
rs.AddNew
rs("imgurl")=dimgurl
rs("adlink")=trim(request("adlink"))
rs.Update
Call CloseRs(rs)

case "edit"
set rs=server.CreateObject("adodb.recordset")
rs.open "select * from DNJY_adfp where id="&id,conn,1,3
rs("imgurl")=dimgurl
rs("adlink")=trim(request("adlink"))
rs.Update
Call Alert("修改成功，并已更新js代码！","?action=js")
Call CloseRs(rs)

case "del"
conn.execute ("delete from DNJY_adfp where id="&id)
Call CloseDB(conn) 
response.Redirect "adfp.asp"

case "js"
Dim fp_pics,fp_links,JsCode,JsCodefp
set rs=server.CreateObject("adodb.recordset")
rs.Open "select * from DNJY_adfp order by id",conn,1,1
if rs.EOF and rs.BOF then
response.write"<tr bgcolor=#FFFFFF><td colspan='4'><p align='center'><font color='red'>无广告！</font></td></tr></table><br>"
follows=0
else
do while not rs.EOF
fp_pics=fp_pics+"/"+strInstallDir+rs("imgurl")+"§"
fp_links=fp_links+rs("ADLink")+"§"
dimgurl=dimgurl+rs("imgurl")+"|"
adlink=adlink+rs("adlink")+"|"
rs.MoveNext
Loop
follows=rs.RecordCount
end If
fp_pics=StrReverse(Mid(StrReverse(fp_pics),2))
fp_links=StrReverse(Mid(StrReverse(fp_links),2))
dimgurl=StrReverse(Mid(StrReverse(dimgurl),2))
adlink=StrReverse(Mid(StrReverse(adlink),2))
dimgurl=split(dimgurl,"|")
adlink=split(adlink,"|")

'生成方式一代码
JsCode="<div id=FP_L style=""WIDTH: 180; TEXT-ALIGN: center""></div>" & vbCrLf & _
"<script type=text/javascript>" & vbCrLf & _
"<!--" & vbCrLf & _
"var fp_pics="""",fp_links="""";" & vbCrLf & _
"fp_pics=""§"&fp_pics&"""" & vbCrLf & _
"fp_links=""§"&fp_links&"""" & vbCrLf & _
"fp_pics=fp_pics.substring(1);" & vbCrLf & _
"fp_links=fp_links.substring(1);" & vbCrLf & _
"var def_pic = ""/"&strInstallDir&"UploadFiles/adfp/01.JPG""; //默认图片地址" & vbCrLf & _
"var def_link = escape(""http://www.ip126.com/""); //默认图片链接" & vbCrLf & _
"var FP = new sinaFlash(""/"&strInstallDir&"img/flipper.swf"", ""FP_L_swf"", 160, 86, ""7"", ""#FFFFFF"", false, ""High"");"&vbCrLf & _
"FP.addParam(""menu"", ""false"");" & vbCrLf & _
"FP.addParam(""wmode"", ""transparent"");" & vbCrLf & _
"FP.addVariable(""ad_num"", ""7""); //翻牌数量" & vbCrLf & _
"FP.addVariable(""pic_width"", ""150""); //图片宽" & vbCrLf & _
"FP.addVariable(""pic_height"", ""80""); //图片高" & vbCrLf & _
"FP.addVariable(""flip_time"", ""300""); //翻牌速度" & vbCrLf & _
"FP.addVariable(""pause_time"", ""2000""); //间隔时间" & vbCrLf & _
"FP.addVariable(""wait_time"", ""1000""); //等待时间" & vbCrLf & _
"FP.addVariable(""pics"", fp_pics); //广告图片地址" & vbCrLf & _
"FP.addVariable(""urls"", fp_links); //广告链接地址" & vbCrLf & _
"FP.addVariable(""def_pic"", def_pic); //默认图片地址" & vbCrLf & _
"FP.addVariable(""def_link"", def_link); //默认链接地址" & vbCrLf & _
"FP.write(""FP_L"");" & vbCrLf & _
"//-->" & vbCrLf & _
"</script>" & vbCrLf & _
""

JsCode=Html2Js(JsCode)
JsCode="document.write('"&JsCode&"')"
JsCode = "<!--" & vbCrLf & JsCode & vbCrLf &"//-->" 
dim fs,f
Set fs = CreateObject("Scripting.FileSystemObject")
Set f = fs.CreateTextFile(Server.MapPath("/")&"/"&""&strInstallDir&"UploadFiles/adfp/sinaflash.js", True)
f.write(JsCode)
f.close
Set f = nothing
Set fs = Nothing

'生成方式二代码＝＝＝＝＝＝
JsCodefp="<DIV id=flipper_div align=center></DIV>" & vbCrLf & _
"<SCRIPT type=text/javascript>" & vbCrLf & _
"<!--" & vbCrLf & _
"var fp_data = new Array();" & vbCrLf & _
"fp_data.push(["""&dimgurl(0)&""","""&adlink(0)&"""]);" & vbCrLf & _
"fp_data.push(["""&dimgurl(1)&""","""&adlink(1)&"""]);" & vbCrLf & _
"fp_data.push(["""&dimgurl(2)&""","""&adlink(2)&"""]);" & vbCrLf & _
"fp_data.push(["""&dimgurl(3)&""","""&adlink(3)&"""]);" & vbCrLf & _
"fp_data.push(["""&dimgurl(4)&""","""&adlink(4)&"""]);" & vbCrLf & _
"fp_data.push(["""&dimgurl(5)&""","""&adlink(5)&"""]);" & vbCrLf & _
"fp_data.push(["""&dimgurl(6)&""","""&adlink(6)&"""]);" & vbCrLf & _
"fp_data.push(["""&dimgurl(7)&""","""&adlink(7)&"""]);" & vbCrLf & _
"fp_data.push(["""&dimgurl(8)&""","""&adlink(8)&"""]);" & vbCrLf & _
"fp_data.push(["""&dimgurl(9)&""","""&adlink(9)&"""]);" & vbCrLf & _
"fp_data.push(["""&dimgurl(10)&""","""&adlink(10)&"""]);" & vbCrLf & _
"fp_data.push(["""&dimgurl(11)&""","""&adlink(11)&"""]);" & vbCrLf & _
"fp_data.push(["""&dimgurl(12)&""","""&adlink(12)&"""]);" & vbCrLf & _
"fp_data.push(["""&dimgurl(13)&""","""&adlink(13)&"""]);" & vbCrLf & _

"var fp_pics="""",fp_links="""";" & vbCrLf & _
"var fp_lens = fp_data.length;" & vbCrLf & _
"for(var i=0; i<fp_lens; i++){" & vbCrLf & _
	"fp_pics += fp_data[i][0];" & vbCrLf & _
	"fp_links += escape(fp_data[i][1]);" & vbCrLf & _
	"if(i%2==0 && i!=fp_lens-1){" & vbCrLf & _
		"fp_pics += ""§"";" & vbCrLf & _
		"fp_links += ""§"";" & vbCrLf & _
	"}else if(i%2==1 && i!=fp_lens-1){" & vbCrLf & _
		"fp_pics += ""§_§"";" & vbCrLf & _
		"fp_links += ""§_§"";" & vbCrLf & _
	"}" & vbCrLf & _
"}" & vbCrLf & _
"var oswf = new sinaFlash(""img/flipper_v2.swf"", ""flipper"", 970, 90, ""7"", ""#FFFFFF"", false, ""High"");" & vbCrLf & _
"oswf.addParam(""allowScriptAccess"", ""always"");" & vbCrLf & _
"oswf.addParam(""menu"", ""false"");" & vbCrLf & _
"oswf.addParam(""wmode"", ""opaque"");" & vbCrLf & _
"oswf.addParam(""scale"", ""noscale"");" & vbCrLf & _
"oswf.addVariable(""pic_width"", ""120"");" & vbCrLf & _
"oswf.addVariable(""pic_height"", ""80"");" & vbCrLf & _
"oswf.addVariable(""colnum"", ""7"");" & vbCrLf & _
"oswf.addVariable(""hspace"", ""15"");" & vbCrLf & _
"oswf.addVariable(""vspace"", ""20"");" & vbCrLf & _
"oswf.addVariable(""flip_time"", ""200"");" & vbCrLf & _
"oswf.addVariable(""pause_time"", ""2000"");" & vbCrLf & _
"oswf.addVariable(""pics"", fp_pics);" & vbCrLf & _
"oswf.addVariable(""urls"", fp_links);" & vbCrLf & _
"oswf.addVariable(""rand"", ""0"");" & vbCrLf & _
"oswf.write(""flipper_div"");" & vbCrLf & _
"//-->" & vbCrLf & _
"</SCRIPT>" & vbCrLf & _
""
JsCodefp=Html2Js(JsCodefp)
JsCodefp="document.write('"&JsCodefp&"')"
JsCodefp = "<!--" & vbCrLf & JsCodefp & vbCrLf &"//-->" 
Set fs = CreateObject("Scripting.FileSystemObject")
Set f = fs.CreateTextFile(Server.MapPath("/")&"/"&""&strInstallDir&"UploadFiles/adfp/sinaflashfp.js", True)
f.write(JsCodefp)
f.close
Set f = nothing
Set fs = Nothing


Call CloseRs(rs)
end select
%>
<html>
<head>
<title>翻牌广告管理</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="style.css" type="text/css">
<style type="text/css">
<!--
body {
	background-color: #E7EEF5;
	margin-left: 0px;
	margin-top: 0px;
	margin-right: 0px;
	margin-bottom: 0px;
}
-->
</style></head>
<body>
<table width="90%" border="1" align="center" cellpadding="3" cellspacing="0" bordercolorlight="#335078" bordercolordark="#FFFFFF" bgcolor="#C5D5E4"  style="font-size:12px; letter-spacing:1px;">
  <tr> 
    <td height="20" bgcolor="#C5D5E4" align="center"><span style="color: #FF0000"><strong><font style="font-size:14px">翻牌广告管理（注意图片类型必须为jpg，120*80像素）</font></strong></span></td>
  </tr>
  <tr> 
    <td bgcolor="#EAEEF4"> <br>
	<table width="98%" border="1" align="center" cellpadding="3" cellspacing="0" bordercolorlight="#335078" bordercolordark="#FFFFFF" bgcolor="#C5D5E4"  style="font-size:12px; letter-spacing:1px;">
        <tr align="center" bgcolor="#FFFFFF" height="20"> 
          <td width="30" bgcolor="#D9E4EE">编号</td>
          <td bgcolor="#D9E4EE">广告图片地址</td>
		  <td bgcolor="#D9E4EE">图片预览</td>
          <td bgcolor="#D9E4EE">广告链接地址（不能带有字符“|”)</td>
          <td bgcolor="#D9E4EE">管理操作</td>
        </tr>
        <%set rs=server.CreateObject("adodb.recordset")
		  rs.Open "select * from DNJY_adfp order by id",conn,1,1
		  dim follows
		  if rs.EOF and rs.BOF then
          response.write"<tr bgcolor=#FFFFFF><td colspan='4'><p align='center'><font color='red'>无此广告！</font></td></tr></table><br>"
		  follows=0          
		  Else
		  i=0
		  do while not rs.EOF
		  i=i+1
		  %>
        <form name="form<%=i%>" method="post" action="?action=edit&id=<%=int(rs("id"))%>">
          <tr bgcolor="#FFFFFF" align="center">
		  <td bgcolor="#F4F7F9"><%=i%></td>		  
            <td bgcolor="#F4F7F9"><input type=text value="<%=trim(rs("imgurl"))%>" name=imgurl<%=i%> size=40 maxlength=100> <INPUT TYPE="button" value="上传图片<%=i%>"  onClick="window.open('../ads/DNJY_upload_ad.asp?ttup=adfp&dform=form<%=i%>&fupname=<%=i%>&frmname=imgurl<%=i%>','blank_','scrollbars=yes,resizable=no,top=40,left=150,width=500,height=150')"></td>
			<td bgcolor="#F4F7F9"><img src=../<%=trim(rs("imgurl"))%>></td>
            <td bgcolor="#F4F7F9"><input name="adlink" type="text" id="adlink" value="<%=trim(rs("adlink"))%>" size="35"></td>
            <td bgcolor="#F4F7F9"><input type="submit" name="Submit" value="修 改">
           <!-- &nbsp; <input type="button" name="DEL" onClick="{if(confirm('确定要删除这个广告吗？\n此操作不可以恢复！')){location.href='?id=<%=rs("id")%>&action=del';}return false;}" value="删除" >  --->          </td>
          </tr>
<INPUT name="i" TYPE="hidden" value="<%=i%>">
        </form>
        <%rs.MoveNext
          loop
          follows=rs.RecordCount
          end If
          %>
      </table>
	<br></td>
  </tr>
</table>
<br>
<!--<table width="90%" border="1" align="center" cellpadding="3" cellspacing="0" bordercolorlight="#335078" bordercolordark="#FFFFFF" bgcolor="#C5D5E4"  style="font-size:12px; letter-spacing:1px;">
  <tr> 
    <td height="20" bgcolor="#C5D5E4" align="center"><span style="color: #FF0000"><strong><font style="font-size:14px">添 加 翻 牌 广 告</font></strong></span></td>
  </tr>
  <tr> 
    <td bgcolor="#EAEEF4">
	<br>
	<table width="98%" border="1" align="center" cellpadding="3" cellspacing="0" bordercolorlight="#335078" bordercolordark="#FFFFFF" bgcolor="#C5D5E4"  style="font-size:12px; letter-spacing:1px;">
        <tr align="center" bgcolor="#FFFFFF" height="20"> 
          <td width="30" bgcolor="#D9E4EE">编号 </td>
          <td bgcolor="#D9E4EE">广告图片地址</td>
          <td bgcolor="#D9E4EE">广告链接地址</td>
          <td bgcolor="#D9E4EE">确定操作</td>
        </tr>
        <form name="form1" method="post" action="?action=add">
          <tr align="center" bgcolor="#FFFFFF"> 
            <td bgcolor="#F4F7F9"><%=rs.RecordCount+1%></td>
	        <td bgcolor="#F4F7F9"><input name="imgurl" type="text" id="imgurl" size="32"></td>
	        <td bgcolor="#F4F7F9"><input name="adlink" type="text" id="adlink" size="20"></td>
	        <td bgcolor="#F4F7F9"><input type="submit" name="Submit3" value="添 加"></td>
          </tr>
        </form>
      </table>
	<br></td>
  </tr>
</table>--->
<%CloseRs(rs)%>
<TABLE width="90%" border="1" align="center" cellpadding="3" cellspacing="0" bordercolorlight="#335078" bordercolordark="#FFFFFF" bgcolor="#C5D5E4"  style="font-size:12px; letter-spacing:1px;">
<TR>
	<TD>方式二预览</TD>
	<TD><!--翻牌方式二开始-->
<%set rs=server.CreateObject("adodb.recordset")
rs.Open "select * from DNJY_adfp order by id",conn,1,1
if rs.EOF and rs.BOF then
response.write"<tr bgcolor=#FFFFFF><td colspan='4'><p align='center'><font color='red'>无广告！</font></td></tr></table><br>"
follows=0
Else
dimgurl=""
adlink=""
do while not rs.eof
dimgurl=dimgurl+rs("imgurl")+"|"
adlink=adlink+rs("adlink")+"|"
rs.MoveNext
Loop
follows=rs.RecordCount
end If
dimgurl=StrReverse(Mid(StrReverse(dimgurl),2))
adlink=StrReverse(Mid(StrReverse(adlink),2))
dimgurl=split(dimgurl,"|")
adlink=split(adlink,"|")
Closers(rs)%>
<DIV id=flipper_div align=center></DIV>
<SCRIPT type=text/javascript>
<!--
var fp_data = new Array();
fp_data.push(["../<%=dimgurl(0)%>","<%=adlink(0)%>"]);
fp_data.push(["../<%=dimgurl(1)%>","<%=adlink(1)%>"]);
fp_data.push(["../<%=dimgurl(2)%>","<%=adlink(2)%>"]);
fp_data.push(["../<%=dimgurl(3)%>","<%=adlink(3)%>"]);
fp_data.push(["../<%=dimgurl(4)%>","<%=adlink(4)%>"]);
fp_data.push(["../<%=dimgurl(5)%>","<%=adlink(5)%>"]);
fp_data.push(["../<%=dimgurl(6)%>","<%=adlink(6)%>"]);

fp_data.push(["../<%=dimgurl(7)%>","<%=adlink(7)%>"]);
fp_data.push(["../<%=dimgurl(8)%>","<%=adlink(8)%>"]);
fp_data.push(["../<%=dimgurl(9)%>","<%=adlink(9)%>"]);
fp_data.push(["../<%=dimgurl(10)%>","<%=adlink(10)%>"]);
fp_data.push(["../<%=dimgurl(11)%>","<%=adlink(11)%>"]);
fp_data.push(["../<%=dimgurl(12)%>","<%=adlink(12)%>"]);
fp_data.push(["../<%=dimgurl(13)%>","<%=adlink(13)%>"]);

var fp_pics="",fp_links="";
var fp_lens = fp_data.length;
for(var i=0; i<fp_lens; i++){
	fp_pics += fp_data[i][0];
	fp_links += escape(fp_data[i][1]);
	if(i%2==0 && i!=fp_lens-1){
		fp_pics += "§";
		fp_links += "§";
	}else if(i%2==1 && i!=fp_lens-1){
		fp_pics += "§_§";
		fp_links += "§_§";
	}
}
var oswf = new sinaFlash("../img/flipper_v2.swf", "flipper", 980, 90, "7", "#FFFFFF", false, "High");
oswf.addParam("allowScriptAccess", "always");
oswf.addParam("menu", "false");
oswf.addParam("wmode", "opaque");
oswf.addParam("scale", "noscale");
oswf.addVariable("pic_width", "120");
oswf.addVariable("pic_height", "75");
oswf.addVariable("colnum", "7");
oswf.addVariable("hspace", "15");
oswf.addVariable("vspace", "20");
oswf.addVariable("flip_time", "200");
oswf.addVariable("pause_time", "2000");
oswf.addVariable("pics", fp_pics);
oswf.addVariable("urls", fp_links);
oswf.addVariable("rand", "0");
oswf.write("flipper_div");
//-->
</SCRIPT></TD>
</TR>
</TABLE>
<table width="90%" border="1" align="center" cellpadding="3" cellspacing="0" bordercolorlight="#335078" bordercolordark="#FFFFFF" bgcolor="#C5D5E4"  style="font-size:12px; letter-spacing:1px;">
  <tr> 
    <td height="20" bgcolor="#C5D5E4" align="center"><span style="color: #FF0000"><strong><font style="font-size:14px">翻 牌 广 告 代 码 生 成 及 调 用 方 式</font></strong></span></td>
  </tr>
  <tr> 
    <td bgcolor="#EAEEF4">
<TABLE>
<TR>
	<TD>   
	<TABLE>
	<textarea name="S1" cols="50" rows="6" class="input2">
<!--翻牌方式一广告代码开始-->
<script src="UploadFiles/adfp/sinaflash.js"></script>
<!--广告代码结束-->
      </textarea>
	<textarea name="S1" cols="50" rows="4" class="input2">注意：此方式广告同一页面只能放置一处，且前面要包含文件<SCRIPT src="Include/sinaflash.js" type=text/javascript></SCRIPT>
      </textarea>
          <tr bgcolor="#FFFFFF" align="center">            
            <td bgcolor="#F4F7F9">
			<form name="forma" method="post" action="?action=js">
			<input type="submit" name="Submit" value="生成js代码">
			</form>
            </td>
          </tr>
		  </TABLE>
        </TD>
<TD>方式一预览<br><!--翻牌方式一开始-->
<div id=FP_L style="WIDTH: 100%; TEXT-ALIGN: center"></div>
<script type=text/javascript>
<!--
var fp_pics="",fp_links="";
<%dim rsnx,sqlnx
set rsnx=server.createobject("adodb.recordset")
sqlnx= "select  * from DNJY_adfp order by id "&DN_OrderType&""
rsnx.open sqlnx,conn,1,1
do while not rsnx.eof%>
fp_pics += "§../"+"<%=rsnx("imgurl")%>"; //广告图片地址
fp_links += "§"+escape("<%=rsnx("adlink")%>"); //广告链接地址
<%rsnx.movenext
loop
rsnx.close
set rsnx=nothing%>
fp_pics=fp_pics.substring(1);
fp_links=fp_links.substring(1);
var def_pic = "../UploadFiles/adfp/01.JPG"; //默认图片地址
var def_link = escape("http://www.ip126.com/"); //默认图片链接
var FP = new sinaFlash("../img/flipper.swf", "FP_L_swf", 160, 86, "7", "#FFFFFF", false, "High");
FP.addParam("menu", "false");
FP.addParam("wmode", "transparent");
FP.addVariable("ad_num", "7"); //翻牌数量
FP.addVariable("pic_width", "140"); //图片宽
FP.addVariable("pic_height", "73"); //图片高
FP.addVariable("flip_time", "300"); //翻牌速度
FP.addVariable("pause_time", "2000"); //间隔时间
FP.addVariable("wait_time", "1000"); //等待时间
FP.addVariable("pics", fp_pics); //广告图片地址
FP.addVariable("urls", fp_links); //广告链接地址
FP.addVariable("def_pic", def_pic); //默认图片地址
FP.addVariable("def_link", def_link); //默认链接地址
FP.write("FP_L");
//-->
  </script>
              <!--翻牌结束--></TD>
	<TD> <TABLE> 
	<textarea name="S1" cols="50" rows="6" class="input2">
<!--翻牌方式二广告代码开始-->
<script src="UploadFiles/adfp/sinaflashfp.js"></script>
<!--广告代码结束-->
      </textarea>
	<textarea name="S1" cols="50" rows="4" class="input2">注意：此方式广告同一页面只能放置一处，且前面要包含文件<SCRIPT src="Include/sinaflashfp.js" type=text/javascript></SCRIPT>
      </textarea>
          <tr bgcolor="#FFFFFF" align="center">            
            <td bgcolor="#F4F7F9">
			<form name="formb" method="post" action="?action=js">
			<input type="submit" name="Submit" value="生成js代码">
			</form>
            </td>
          </tr>
		</TABLE></TD>
</TR>
</TABLE>

		</td>
  </tr>
</table>
</body> 
</html>
 <%Call CloseDB(conn)%>