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
Call Alert("�޸ĳɹ������Ѹ���js���룡","?action=js")
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
response.write"<tr bgcolor=#FFFFFF><td colspan='4'><p align='center'><font color='red'>�޹�棡</font></td></tr></table><br>"
follows=0
else
do while not rs.EOF
fp_pics=fp_pics+"/"+strInstallDir+rs("imgurl")+"��"
fp_links=fp_links+rs("ADLink")+"��"
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

'���ɷ�ʽһ����
JsCode="<div id=FP_L style=""WIDTH: 180; TEXT-ALIGN: center""></div>" & vbCrLf & _
"<script type=text/javascript>" & vbCrLf & _
"<!--" & vbCrLf & _
"var fp_pics="""",fp_links="""";" & vbCrLf & _
"fp_pics=""��"&fp_pics&"""" & vbCrLf & _
"fp_links=""��"&fp_links&"""" & vbCrLf & _
"fp_pics=fp_pics.substring(1);" & vbCrLf & _
"fp_links=fp_links.substring(1);" & vbCrLf & _
"var def_pic = ""/"&strInstallDir&"UploadFiles/adfp/01.JPG""; //Ĭ��ͼƬ��ַ" & vbCrLf & _
"var def_link = escape(""http://www.ip126.com/""); //Ĭ��ͼƬ����" & vbCrLf & _
"var FP = new sinaFlash(""/"&strInstallDir&"img/flipper.swf"", ""FP_L_swf"", 160, 86, ""7"", ""#FFFFFF"", false, ""High"");"&vbCrLf & _
"FP.addParam(""menu"", ""false"");" & vbCrLf & _
"FP.addParam(""wmode"", ""transparent"");" & vbCrLf & _
"FP.addVariable(""ad_num"", ""7""); //��������" & vbCrLf & _
"FP.addVariable(""pic_width"", ""150""); //ͼƬ��" & vbCrLf & _
"FP.addVariable(""pic_height"", ""80""); //ͼƬ��" & vbCrLf & _
"FP.addVariable(""flip_time"", ""300""); //�����ٶ�" & vbCrLf & _
"FP.addVariable(""pause_time"", ""2000""); //���ʱ��" & vbCrLf & _
"FP.addVariable(""wait_time"", ""1000""); //�ȴ�ʱ��" & vbCrLf & _
"FP.addVariable(""pics"", fp_pics); //���ͼƬ��ַ" & vbCrLf & _
"FP.addVariable(""urls"", fp_links); //������ӵ�ַ" & vbCrLf & _
"FP.addVariable(""def_pic"", def_pic); //Ĭ��ͼƬ��ַ" & vbCrLf & _
"FP.addVariable(""def_link"", def_link); //Ĭ�����ӵ�ַ" & vbCrLf & _
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

'���ɷ�ʽ�����룽����������
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
		"fp_pics += ""��"";" & vbCrLf & _
		"fp_links += ""��"";" & vbCrLf & _
	"}else if(i%2==1 && i!=fp_lens-1){" & vbCrLf & _
		"fp_pics += ""��_��"";" & vbCrLf & _
		"fp_links += ""��_��"";" & vbCrLf & _
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
<title>���ƹ�����</title>
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
    <td height="20" bgcolor="#C5D5E4" align="center"><span style="color: #FF0000"><strong><font style="font-size:14px">���ƹ�����ע��ͼƬ���ͱ���Ϊjpg��120*80���أ�</font></strong></span></td>
  </tr>
  <tr> 
    <td bgcolor="#EAEEF4"> <br>
	<table width="98%" border="1" align="center" cellpadding="3" cellspacing="0" bordercolorlight="#335078" bordercolordark="#FFFFFF" bgcolor="#C5D5E4"  style="font-size:12px; letter-spacing:1px;">
        <tr align="center" bgcolor="#FFFFFF" height="20"> 
          <td width="30" bgcolor="#D9E4EE">���</td>
          <td bgcolor="#D9E4EE">���ͼƬ��ַ</td>
		  <td bgcolor="#D9E4EE">ͼƬԤ��</td>
          <td bgcolor="#D9E4EE">������ӵ�ַ�����ܴ����ַ���|��)</td>
          <td bgcolor="#D9E4EE">�������</td>
        </tr>
        <%set rs=server.CreateObject("adodb.recordset")
		  rs.Open "select * from DNJY_adfp order by id",conn,1,1
		  dim follows
		  if rs.EOF and rs.BOF then
          response.write"<tr bgcolor=#FFFFFF><td colspan='4'><p align='center'><font color='red'>�޴˹�棡</font></td></tr></table><br>"
		  follows=0          
		  Else
		  i=0
		  do while not rs.EOF
		  i=i+1
		  %>
        <form name="form<%=i%>" method="post" action="?action=edit&id=<%=int(rs("id"))%>">
          <tr bgcolor="#FFFFFF" align="center">
		  <td bgcolor="#F4F7F9"><%=i%></td>		  
            <td bgcolor="#F4F7F9"><input type=text value="<%=trim(rs("imgurl"))%>" name=imgurl<%=i%> size=40 maxlength=100> <INPUT TYPE="button" value="�ϴ�ͼƬ<%=i%>"  onClick="window.open('../ads/DNJY_upload_ad.asp?ttup=adfp&dform=form<%=i%>&fupname=<%=i%>&frmname=imgurl<%=i%>','blank_','scrollbars=yes,resizable=no,top=40,left=150,width=500,height=150')"></td>
			<td bgcolor="#F4F7F9"><img src=../<%=trim(rs("imgurl"))%>></td>
            <td bgcolor="#F4F7F9"><input name="adlink" type="text" id="adlink" value="<%=trim(rs("adlink"))%>" size="35"></td>
            <td bgcolor="#F4F7F9"><input type="submit" name="Submit" value="�� ��">
           <!-- &nbsp; <input type="button" name="DEL" onClick="{if(confirm('ȷ��Ҫɾ����������\n�˲��������Իָ���')){location.href='?id=<%=rs("id")%>&action=del';}return false;}" value="ɾ��" >  --->          </td>
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
    <td height="20" bgcolor="#C5D5E4" align="center"><span style="color: #FF0000"><strong><font style="font-size:14px">�� �� �� �� �� ��</font></strong></span></td>
  </tr>
  <tr> 
    <td bgcolor="#EAEEF4">
	<br>
	<table width="98%" border="1" align="center" cellpadding="3" cellspacing="0" bordercolorlight="#335078" bordercolordark="#FFFFFF" bgcolor="#C5D5E4"  style="font-size:12px; letter-spacing:1px;">
        <tr align="center" bgcolor="#FFFFFF" height="20"> 
          <td width="30" bgcolor="#D9E4EE">��� </td>
          <td bgcolor="#D9E4EE">���ͼƬ��ַ</td>
          <td bgcolor="#D9E4EE">������ӵ�ַ</td>
          <td bgcolor="#D9E4EE">ȷ������</td>
        </tr>
        <form name="form1" method="post" action="?action=add">
          <tr align="center" bgcolor="#FFFFFF"> 
            <td bgcolor="#F4F7F9"><%=rs.RecordCount+1%></td>
	        <td bgcolor="#F4F7F9"><input name="imgurl" type="text" id="imgurl" size="32"></td>
	        <td bgcolor="#F4F7F9"><input name="adlink" type="text" id="adlink" size="20"></td>
	        <td bgcolor="#F4F7F9"><input type="submit" name="Submit3" value="�� ��"></td>
          </tr>
        </form>
      </table>
	<br></td>
  </tr>
</table>--->
<%CloseRs(rs)%>
<TABLE width="90%" border="1" align="center" cellpadding="3" cellspacing="0" bordercolorlight="#335078" bordercolordark="#FFFFFF" bgcolor="#C5D5E4"  style="font-size:12px; letter-spacing:1px;">
<TR>
	<TD>��ʽ��Ԥ��</TD>
	<TD><!--���Ʒ�ʽ����ʼ-->
<%set rs=server.CreateObject("adodb.recordset")
rs.Open "select * from DNJY_adfp order by id",conn,1,1
if rs.EOF and rs.BOF then
response.write"<tr bgcolor=#FFFFFF><td colspan='4'><p align='center'><font color='red'>�޹�棡</font></td></tr></table><br>"
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
		fp_pics += "��";
		fp_links += "��";
	}else if(i%2==1 && i!=fp_lens-1){
		fp_pics += "��_��";
		fp_links += "��_��";
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
    <td height="20" bgcolor="#C5D5E4" align="center"><span style="color: #FF0000"><strong><font style="font-size:14px">�� �� �� �� �� �� �� �� �� �� �� �� ʽ</font></strong></span></td>
  </tr>
  <tr> 
    <td bgcolor="#EAEEF4">
<TABLE>
<TR>
	<TD>   
	<TABLE>
	<textarea name="S1" cols="50" rows="6" class="input2">
<!--���Ʒ�ʽһ�����뿪ʼ-->
<script src="UploadFiles/adfp/sinaflash.js"></script>
<!--���������-->
      </textarea>
	<textarea name="S1" cols="50" rows="4" class="input2">ע�⣺�˷�ʽ���ͬһҳ��ֻ�ܷ���һ������ǰ��Ҫ�����ļ�<SCRIPT src="Include/sinaflash.js" type=text/javascript></SCRIPT>
      </textarea>
          <tr bgcolor="#FFFFFF" align="center">            
            <td bgcolor="#F4F7F9">
			<form name="forma" method="post" action="?action=js">
			<input type="submit" name="Submit" value="����js����">
			</form>
            </td>
          </tr>
		  </TABLE>
        </TD>
<TD>��ʽһԤ��<br><!--���Ʒ�ʽһ��ʼ-->
<div id=FP_L style="WIDTH: 100%; TEXT-ALIGN: center"></div>
<script type=text/javascript>
<!--
var fp_pics="",fp_links="";
<%dim rsnx,sqlnx
set rsnx=server.createobject("adodb.recordset")
sqlnx= "select  * from DNJY_adfp order by id "&DN_OrderType&""
rsnx.open sqlnx,conn,1,1
do while not rsnx.eof%>
fp_pics += "��../"+"<%=rsnx("imgurl")%>"; //���ͼƬ��ַ
fp_links += "��"+escape("<%=rsnx("adlink")%>"); //������ӵ�ַ
<%rsnx.movenext
loop
rsnx.close
set rsnx=nothing%>
fp_pics=fp_pics.substring(1);
fp_links=fp_links.substring(1);
var def_pic = "../UploadFiles/adfp/01.JPG"; //Ĭ��ͼƬ��ַ
var def_link = escape("http://www.ip126.com/"); //Ĭ��ͼƬ����
var FP = new sinaFlash("../img/flipper.swf", "FP_L_swf", 160, 86, "7", "#FFFFFF", false, "High");
FP.addParam("menu", "false");
FP.addParam("wmode", "transparent");
FP.addVariable("ad_num", "7"); //��������
FP.addVariable("pic_width", "140"); //ͼƬ��
FP.addVariable("pic_height", "73"); //ͼƬ��
FP.addVariable("flip_time", "300"); //�����ٶ�
FP.addVariable("pause_time", "2000"); //���ʱ��
FP.addVariable("wait_time", "1000"); //�ȴ�ʱ��
FP.addVariable("pics", fp_pics); //���ͼƬ��ַ
FP.addVariable("urls", fp_links); //������ӵ�ַ
FP.addVariable("def_pic", def_pic); //Ĭ��ͼƬ��ַ
FP.addVariable("def_link", def_link); //Ĭ�����ӵ�ַ
FP.write("FP_L");
//-->
  </script>
              <!--���ƽ���--></TD>
	<TD> <TABLE> 
	<textarea name="S1" cols="50" rows="6" class="input2">
<!--���Ʒ�ʽ�������뿪ʼ-->
<script src="UploadFiles/adfp/sinaflashfp.js"></script>
<!--���������-->
      </textarea>
	<textarea name="S1" cols="50" rows="4" class="input2">ע�⣺�˷�ʽ���ͬһҳ��ֻ�ܷ���һ������ǰ��Ҫ�����ļ�<SCRIPT src="Include/sinaflashfp.js" type=text/javascript></SCRIPT>
      </textarea>
          <tr bgcolor="#FFFFFF" align="center">            
            <td bgcolor="#F4F7F9">
			<form name="formb" method="post" action="?action=js">
			<input type="submit" name="Submit" value="����js����">
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