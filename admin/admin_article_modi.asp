<!--#include file="../conn.asp"-->
<!--#include file="../config.asp"-->
<!--#include file="../Include/head.asp"-->
<!--#include file="../Include/Function.asp"-->
<!--#include file="../Include/label.asp"-->
<!--#include file="../Include/RemoteImg_save.asp"-->
<!--#include file=cookies.asp-->
<%
'=====================================
'���칩����Ϣ��վ����ϵͳ
'��������Ƽ������ҳ�Ʒ
'�ͷ�����http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
call checkmanage("06")
dim rso,Content,newsuser,come,T_SaveImg,AspJpeg,T_UploadDir,c_oneid,c_twoid,c_threeid,c_one,c_two,c_three

If request("del")="yes" Then'ɾ�����º����ͼƬ
Dim str2,contentimg,Strsimg,photoimg,rsdel,sqldel
Server.ScriptTimeout=50000
id=trim(request("DelID"))
if trim(id)="" then
response.write "<script language=JavaScript>" & chr(13) & "alert('û��ѡ���¼��');" &"window.location='admin_article_list.asp'" & "</script>"
response.end
end if
str2=split(id,",")
for i=0 to ubound(str2)
set rsdel=server.createobject("adodb.recordset")
set rs=server.createobject("adodb.recordset")
sql="select content,s_photo from DNJY_news where id="&trim(str2(i))
rs.open sql,conn,1,1
contentimg=content_pic(rs("content"))
If contentimg<>"" Then'ɾ��������ͼƬ
contentimg=split(contentimg,"|")
For Each Strsimg In contentimg
DoDelFile(Strsimg)
Next
End If
photoimg=rs("s_photo")
If photoimg<>"" Then'ɾ�������ϴ����ͼƬ
photoimg=split(photoimg,"|")
For Each Strsimg In photoimg
DoDelFile("/"&strInstallDir&Replace(Strsimg,"_s",""))'ɾ��ԭͼ
DoDelFile("/"&strInstallDir&Strsimg)'ɾ������ͼ
Next
End If
sqldel="delete from [DNJY_news] where id="&cstr(str2(i))
rsdel.open sqldel,conn,1,3
Next'ѭ������
Call CloseRs(rs)
response.write "<script language=JavaScript>" & chr(13) & "alert('ɾ���ɹ���');" &"window.location='admin_article_list.asp'" & "</script>"
response.end
End If'ɾ��END

if request("no")="modi" Then'�޸�
title=trim(request.form("title"))
keywords=Replace_Text(trim(request("keywords")))
id=request("id")
city_oneid=HtmlEncode(request.form("city_one"))
city_twoid=HtmlEncode(request.form("city_two"))
city_threeid=HtmlEncode(request.form("city_three"))
if city_oneid<>"" then
set rs=server.createobject("adodb.recordset")
rs.open "select city from DNJY_city where twoid=0 and threeid=0 and id="&city_oneid,conn,1,1
city_one=rs("city")
rs.close
if city_twoid<>"" then
rs.open "select city from DNJY_city where id="&city_oneid&" and threeid=0 and twoid="&city_twoid,conn,1,1
city_two=rs("city")
rs.close
end if
if city_twoid<>"" and city_threeid<>"" then
rs.open "select city from DNJY_city where id="&city_oneid&" and twoid="&city_twoid&" and threeid="&city_threeid,conn,1,1
city_three=rs("city")
rs.close
end If
set rs=nothing
end If

c_oneid=HtmlEncode(request.form("c_oneid"))
c_twoid=HtmlEncode(request.form("c_twoid"))
c_threeid=HtmlEncode(request.form("c_threeid"))
if c_oneid<>"" then
set rs=server.createobject("adodb.recordset")
rs.open "select name from DNJY_news_class where twoid=0 and threeid=0 and oneid="&c_oneid,conn,1,1
c_one=rs("name")
rs.close
if c_twoid<>"" then
rs.open "select name from DNJY_news_class where oneid="&c_oneid&" and threeid=0 and twoid="&c_twoid,conn,1,1
c_two=rs("name")
rs.close
end if
if c_twoid<>"" and c_threeid<>"" then
rs.open "select name from DNJY_news_class where oneid="&c_oneid&" and twoid="&c_twoid&" and threeid="&c_threeid,conn,1,1
c_three=rs("name")
rs.close
end If
set rs=nothing
end If

If trim(request("content"))="" Then
Call Alert ("����д����","-1")
else
content=trim(request("content"))
End If
newsuser=request("newsuser")
if newsuser="" then
	newsuser="ip126"
end if
come=request("come")
if come="" then
	come=web
end If



set rs=server.createobject("adodb.recordset")
sql="select * from DNJY_news where id="&id
rs.open sql,conn,1,3
if city_oneid="" then city_oneid=0
rs("city_oneid")=city_oneid
rs("city_one")=city_one
if city_twoid="" then city_twoid=0
rs("city_twoid")=city_twoid
rs("city_two")=city_two
if city_threeid="" then city_threeid=0
rs("city_threeid")=city_threeid
rs("city_three")=city_three
if c_oneid="" then c_oneid=0
rs("c_oneid")=c_oneid
rs("c_one")=c_one
if c_twoid="" then c_twoid=0
rs("c_twoid")=c_twoid
rs("c_two")=c_two
if c_threeid="" then c_threeid=0
rs("c_threeid")=c_threeid
rs("c_three")=c_three
rs("title")=title
rs("keywords")=keywords
rs("content")=content
rs("oStyle")=request.form("oStyle")
rs("oColor")=request.form("oColor")
rs("newstop")=request("newstop")
rs("tuijian")=request("tuijian")
rs("newsuser")=newsuser
rs("come")=come
rs("ifshow")=request("ifshow")
rs.update
Call CloseRs(rs)
Call Alert ("�����޸ĳɹ�","admin_article_list.asp")
Response.End
end if%>
<html>
<head>
<title>�޸�����</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="../css/manage.css" type="text/css">
<script language = "JavaScript">
function CheckForm()
{
	if (document.addNEWS.title.value.length == 0) {
		alert("���±���û����д.");
		document.addNEWS.title.focus();
		return false;
	}
	if (document.addNEWS.keywords.value.length == 0) {
		alert("���¹ؼ���û����д.");
		document.addNEWS.keywords.focus();
		return false;
	}
        var editor=KindEditor.create('#editor');
        if (editor.isEmpty())
	    {
	    alert("�������ݲ���Ϊ�գ�")	  
	     return false
	     }
		if (document.addNEWS.newsuser.value.length == 0) {
		alert("��������û����д");
		document.addNEWS.newsuser.focus();
		return false;
	}
		if (document.addNEWS.come.value.length == 0) {
		alert("������Դû����д");
		document.addNEWS.come.focus();
		return false;
	}
	return true;
}
</script>
<script language=javascript src="../Include/mouse_on_title.js"></script>
<!--kindeditor�༭��-->
	<link rel="stylesheet" href="../KindEditor/themes/default/default.css" />
	<link rel="stylesheet" href="../KindEditor/plugins/code/prettify.css" />
	<script charset="utf-8" src="../KindEditor/kindeditor.js"></script>
	<script charset="utf-8" src="../KindEditor/lang/zh_CN.js"></script>
	<script charset="utf-8" src="../KindEditor/plugins/code/prettify.js"></script>
	<script>
		KindEditor.ready(function(K) {
			var editor = K.create('textarea[name="content"]', {
				cssPath : '../KindEditor/plugins/code/prettify.css',
				uploadJson : '../KindEditor/asp/upload_json.asp?action=photo',
				afterBlur: function(){this.sync();},//ʧȥ����ִ��this���ֵ
				allowFileManager : true,				
				afterCreate : function() {
					var self = this;
					K.ctrl(document, 13, function() {
						self.sync();
						K('form[name=example]')[0].submit();
					});
					K.ctrl(self.edit.doc, 13, function() {
						self.sync();
						K('form[name=example]')[0].submit();
					});
				}
			});
			prettyPrint();
		});
	</script>
<!--kindeditor�༭��END-->
</head>
<body leftmargin="0" topmargin="0" bgcolor="#E3EBF9">
<%id=request("id")
Set rso=Server.CreateObject("ADODB.RecordSet") 
sql="select * from DNJY_news where  id="&id
rso.Open sql,conn,1,1
if rso.eof and rso.bof then
response.Write("û�м�¼")
else%>
<form name="addNEWS" method="post" action="admin_article_modi.asp?no=modi" onSubmit="return CheckForm();">
<table width="98%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="#6699cc">  
    <tr align="center" bgcolor="#E3EBF9"> 
      <td height="30" colspan="3"><font color="#0000FF"><strong>�޸�����</strong></font></td>
    </tr>
    <tr> 
      <td width="100" height="30" align="right" bgcolor="#E3EBF9">&nbsp;&nbsp;����������</td>
      <td colspan="2" valign="top" bgcolor="#E3EBF9">�� 
<%Dim rsi
set rsi=conn.execute("select * from DNJY_city where id>0 and twoid>0 and threeid=0")%>
      <SCRIPT language = "JavaScript">
var onecount;
onecount=0;
subcat = new Array();
        <%dim count:count = 0
        do while not rsi.eof%>
subcat[<%=count%>] = new Array("<%=rsi("city")%>","<%=rsi("id")%>","<%=rsi("twoid")%>");
        <%count = count + 1
        rsi.movenext
        loop
        rsi.close
		set rsi=nothing%>
onecount=<%=count%>;
</SCRIPT>
<%set rsi=conn.execute("select * from DNJY_city where id>0 and twoid>0 and threeid>0")%>
 <SCRIPT language = "JavaScript">
var onecount4;
onecount4=0;
subcat4 = new Array();
        <%dim count4:count4 = 0
        do while not rsi.eof%>
subcat4[<%=count4%>] = new Array("<%=rsi("city")%>","<%=rsi("id")%>","<%=rsi("twoid")%>","<%=rsi("threeid")%>");
        <%count4 = count4 + 1
        rsi.movenext
        loop
        rsi.close
		set rsi = nothing%>
onecount4=<%=count4%>;

function changelocation(locationid)
    {
    document.addNEWS.city_two.length = 0; 
	document.addNEWS.city_two.options[0] = new Option('ѡ�����','');
	document.addNEWS.city_three.length = 0; 
	document.addNEWS.city_three.options[0] = new Option('ѡ�����','');
    var locationid=locationid;
    var i;
    for (i=0;i < onecount; i++)
        {
            if (subcat[i][1] == locationid)
            { 
                document.addNEWS.city_two.options[document.addNEWS.city_two.length] = new Option(subcat[i][0], subcat[i][2]);
            }        
        }
        
    }    
	
	function changelocation4(locationid,locationid1)
    {
    document.addNEWS.city_three.length = 0; 
    document.addNEWS.city_three.options[0] = new Option('ѡ�����','');
    var locationid=locationid;
	 var locationid1=locationid1;
    var i;
    for (i=0;i < onecount4; i++)
        {
            if (subcat4[i][2] == locationid)
            { 
			if (subcat4[i][1] == locationid1)
			{
                document.addNEWS.city_three.options[document.addNEWS.city_three.length] = new Option(subcat4[i][0], subcat4[i][3]);
            }        
        }
       } 
    }            
              </SCRIPT>
      <SELECT name="city_one" size="1" id="select2" onChange="changelocation(document.addNEWS.city_one.options[document.addNEWS.city_one.selectedIndex].value)">
            <OPTION value="" selected>ѡ�����</OPTION>
  <%set rsi=conn.execute("select * from DNJY_city where id>0 and twoid=0 and threeid=0 order by indexid asc")
if rsi.eof or rsi.bof then
response.write "<option value=''>û�з���</option>"
else
do until rsi.eof%>
  <OPTION value="<%=rsi("id")%>" <%if rsi("id")=rso("city_oneid") then%>selected<%end if%>><%=rsi("city")%></OPTION>
 <%rsi.movenext
    loop
	end if
rsi.close
set rsi = nothing%>
      </SELECT> 
	  <SELECT name="city_two" id="city_two" onChange="changelocation4(document.addNEWS.city_two.options[document.addNEWS.city_two.selectedIndex].value,document.addNEWS.city_one.options[document.addNEWS.city_one.selectedIndex].value)">
    <OPTION value="" selected>ѡ�����</OPTION>
<%set rsi=conn.execute("select * from DNJY_city where id="&rso("city_oneid")&" and twoid>0 and threeid=0")
if rsi.eof and rsi.bof then
response.write "<option value=''>û�з���</option>"
else
do until rsi.eof%>
  <OPTION value="<%=rsi("twoid")%>" <%if rsi("twoid")=rso("city_twoid") then%>selected<%end if%>><%=rsi("city")%></OPTION>
 <%rsi.movenext
    loop
	end if
rsi.close
set rsi = nothing%>
	</SELECT>
	     <SELECT name="city_three" id="city_three">
         <OPTION value="" selected>ѡ�����</OPTION>
<%set rsi=conn.execute("select * from DNJY_city where id="&rso("city_oneid")&" and twoid="&rso("city_twoid")&" and threeid<>0")
if rsi.eof and rsi.bof then
response.write "<option value=''>û�з���</option>"
else
do until rsi.eof%>
<OPTION value="<%=rsi("threeid")%>" <%if rsi("threeid")=rso("city_threeid") then%>selected<%end if%>><%=rsi("city")%></OPTION>
   <%rsi.movenext
    loop
 end if
rsi.close
set rsi = nothing%>
    </SELECT></td>
    </tr>
    <tr> 
      <td height="30" align="right" bgcolor="#E3EBF9">������Ŀ��</td>
      <td colspan="2" valign="top" bgcolor="#E3EBF9">&nbsp;&nbsp;&nbsp;&nbsp;
<%set rsi=conn.execute("select * from DNJY_news_class where oneid>0 and twoid>0 and threeid=0")%>
 <SCRIPT language = "JavaScript">
var onecount2;
onecount2=0;
subcat2 = new Array();
        <%dim count2:count2 = 0
        do while not rsi.eof%>
subcat2[<%=count2%>] = new Array("<%=rsi("name")%>","<%=rsi("oneid")%>","<%=rsi("twoid")%>");
        <%count2 = count2 + 1
        rsi.movenext
        loop
        rsi.close
	set rsi = nothing%>
onecount2=<%=count2%>;
</SCRIPT>
<%set rsi=conn.execute("select * from DNJY_news_class where oneid>0 and twoid>0 and threeid>0")%>
<SCRIPT language = "JavaScript">
var onecount3;
onecount3=0;
subcat3 = new Array();
        <%dim count3:count3 = 0
        do while not rsi.eof%>
subcat3[<%=count3%>] = new Array("<%=rsi("name")%>","<%=rsi("oneid")%>","<%=rsi("twoid")%>","<%=rsi("threeid")%>");
        <%count3 = count3 + 1
        rsi.movenext
        loop
        rsi.close
	set rsi = nothing%>
        onecount3=<%=count3%>;
function changelocation2(locationid)
    {
    document.addNEWS.c_twoid.length = 0; 
    document.addNEWS.c_twoid.options[0] = new Option('ѡ����Ŀ','');
	document.addNEWS.c_threeid.length = 0; 
    document.addNEWS.c_threeid.options[0] = new Option('ѡ����Ŀ','');
    var locationid=locationid;
	var i;
    for (i=0;i < onecount2; i++)
        {
            if (subcat2[i][1] == locationid)
            { 
                document.addNEWS.c_twoid.options[document.addNEWS.c_twoid.length] = new Option(subcat2[i][0], subcat2[i][2]);
            }        
        }
       
    }    
     function changelocation3(locationid,locationid1)
    {
    document.addNEWS.c_threeid.length = 0; 
    document.addNEWS.c_threeid.options[0] = new Option('ѡ����Ŀ','');
    var locationid=locationid;
	 var locationid1=locationid1;
    var i;
    for (i=0;i < onecount3; i++)
        {
            if (subcat3[i][2] == locationid)
            { 
			if (subcat3[i][1] == locationid1)
			{
                document.addNEWS.c_threeid.options[document.addNEWS.c_threeid.length] = new Option(subcat3[i][0], subcat3[i][3]);
            }        
        }
       } 
    }            
	                       </SCRIPT>
                                 <SELECT name="c_oneid" size="1" id="select3" onChange="changelocation2(document.addNEWS.c_oneid.options[document.addNEWS.c_oneid.selectedIndex].value)">
                                   <OPTION value="" selected>ѡ����Ŀ</OPTION>
<%set rsi=conn.execute("select * from DNJY_news_class where oneid>0 and twoid=0 order by sorting asc")
if rsi.eof and rsi.bof then
response.write "<option value=''>û�з���</option>"
else
do until rsi.eof%>
                                   <OPTION value="<%=rsi("oneid")%>" <%if rsi("oneid")=rso("c_oneid") then%>selected<%end if%>><%=rsi("name")%></OPTION>
<% rsi.movenext
    loop
	end if
rsi.close
set rsi= nothing%>
                                 </SELECT>
                                 <SELECT name="c_twoid" id="select4" onChange="changelocation3(document.addNEWS.c_twoid.options[document.addNEWS.c_twoid.selectedIndex].value,document.addNEWS.c_oneid.options[document.addNEWS.c_oneid.selectedIndex].value)">
                                   <OPTION value="" selected>ѡ����Ŀ</OPTION>
<%set rsi=conn.execute("select * from DNJY_news_class where oneid="&rso("c_oneid")&" and twoid<>0 and threeid=0")
if rsi.eof and rsi.bof then
response.write "<option value=''>û�з���</option>"
else
do until rsi.eof%>
                                   <OPTION value="<%=rsi("twoid")%>" <%if rsi("twoid")=rso("c_twoid") then%>selected<%end if%>><%=rsi("name")%></OPTION>
<% rsi.movenext
    loop
	end if
rsi.close
set rsi= nothing%>
                                 </SELECT>
                                 <SELECT name="c_threeid" id="select5">
                                   <OPTION value="" selected>ѡ����Ŀ</OPTION>
                                   <%set rsi=conn.execute("select * from DNJY_news_class where oneid="&rso("c_oneid")&" and twoid="&rso("c_twoid")&" and threeid<>0")
if rsi.eof and rsi.bof then
response.write "<option value=''>û�з���</option>"
else
do until rsi.eof%>
                                   <OPTION value="<%=rsi("threeid")%>" <%if rsi("threeid")=rso("c_threeid") then%>selected<%end if%>><%=rsi("name")%></OPTION>
<% rsi.movenext
    loop
	end if
rsi.close
set rsi= Nothing%>
</SELECT>
</td>
    </tr>
    <tr> 
      <td width="100" height="30" align="right" bgcolor="#E3EBF9">���±��⣺</td>
      <td colspan="2" valign="top" bgcolor="#E3EBF9">�� 
        <input name="title" type="text" class="input" value="<%=rso("title")%>" size="50" maxlength="100"><font color="#FF0000">*</font></td>
    </tr>

    <tr> 
      <td width="100" height="30" align="right" bgcolor="#E3EBF9">�� �� �ʣ�</td>
      <td colspan="2" valign="top" bgcolor="#E3EBF9">�� 
        <input name="keywords" type="text" class="input" value="<%=rso("keywords")%>" size="50" maxlength="40"> <font color="#FF0000">*</font> <input type="button" name="btn" value="�ӱ��⸴��" title="���Ʊ��⵽�ؼ���" onClick="CopyWebTitle(document.addNEWS.title.value);"> <br>��Ӣ�ġ�,���ŷָ�����40���ַ�</td>
    </tr>
	<tr>
  <td height="30" align="right" bgcolor="#E3EBF9">�������ԣ�</td>
  <td colspan="2" bgcolor="#E3EBF9">
         <select class='css_select' name="oStyle" type="text" id="oStyle">
          <option value="">������ʽ</option>
          <option value="s01" <%if rso("oStyle")="s01" then%> selected<%end if%>>����</option>
          <option value="s02" <%if rso("oStyle")="s02" then%> selected<%end if%>>б��</option>
        </select>
        <select class='css_select' name="oColor" type="text" id="oColor">
          <option value="links">������ɫ</option>
          <option value="a01" style="background-color:#FF0000;" <%if rso("oColor")="a01" then%> selected<%end if%>></option>
          <option value="a02" style="background-color:#0000FF;" <%if rso("oColor")="a02" then%> selected<%end if%>></option>
          <option value="a03" style="background-color:#00FFFF;" <%if rso("oColor")="a03" then%> selected<%end if%>></option>
          <option value="a04" style="background-color:#FF9900;" <%if rso("oColor")="a04" then%> selected<%end if%>></option>
          <option value="a05" style="background-color:#339966;" <%if rso("oColor")="a05" then%> selected<%end if%>></option>
        </select>
		<select class='css_select' name="newstop" type="text" id="newstop">
          <option value="0">�ö�</option>
          <option value="1" <%if rso("newstop")=1 then%> selected<%end if%>>��</option>
          <option value="0" <%if rso("newstop")=0 then%> selected<%end if%>>��</option>
        </select>
		<select class='css_select' name="tuijian" type="text" id="tuijian">
          <option value="0">�Ƽ�</option>
          <option value="1" <%if rso("tuijian")=1 then%> selected<%end if%>>��</option>
          <option value="0" <%if rso("tuijian")=0 then%> selected<%end if%>>��</option>
        </select></td>
</tr>

    <tr> 
      <td height="30" align="right" valign="top" bgcolor="#E3EBF9">�������ݣ�</td>
      <td colspan="2" valign="top" bgcolor="#E3EBF9"><textarea id="editor" name="content" style="width:670px;height:400px;"><%=rso("content")%></textarea><input type="checkbox" name="T_SaveImg" value="1" /> ����������Զ��ͼƬ������&nbsp;&nbsp;<input type="checkbox" name="AspJpeg" value="1" checked/>���ϴ�ͼƬ��ˮӡ&nbsp;&nbsp;<font color="#FF0000">*</font></td>
    </tr>
	<tr> 
      <td height="24"  align="right" bgcolor="#E3EBF9">�������ߣ�</td>
      <td colspan="2" valign="top" bgcolor="#E3EBF9">�� <input name="newsuser" type="text" class="input" size="30" value="<%=rso("newsuser")%>"> <font color="#ff0000">*</font> &lt;=��<FONT color=blue><FONT 
      style="CURSOR: hand" onclick="newsuser.value='վ��'" 
      color=#993300>վ��</FONT></FONT>����<FONT 
      color=blue><FONT style="CURSOR: hand" onclick='newsuser.value="δ֪"' 
      color=#993300>δ֪</FONT></FONT>����<FONT color=blue><FONT 
      style="CURSOR: hand" onclick="newsuser.value='����'" 
      color=#993300>����</FONT></FONT>��</td>
    </tr>
    <tr> 
      <td height="24"  align="right" bgcolor="#E3EBF9">������Դ��</td>
      <td colspan="2" valign="top" bgcolor="#E3EBF9">�� <input name="come" type="text" class="input" size="30" value="<%=rso("come")%>"> <font color="#ff0000">*</font> &lt;=��<FONT color=blue><FONT 
      style="CURSOR: hand" onclick="come.value='����'" 
      color=#993300>����</FONT></FONT>��</td>
    </tr>
     <tr> 
      <td height="25" align="right">������ʾ��</td>
      <td colspan="2"> 
        <input type="radio" name="ifshow" value="1" <%if rso("ifshow")=1 then Response.Write "checked"%>>
        ��ʾ 
        <input name="ifshow" type="radio" value="0" <%if rso("ifshow")=0 then Response.Write "checked"%>>
        ����</td>
    </tr>  
    <tr align="center"> 
      <td height="35" colspan="3" bgcolor="#E3EBF9"> 
        <input type="submit" name="Submit" value="�ύ" class="input">
        <input type="hidden" name="Id" value="<%=Id%>">�� 
        <input type="reset" name="Submit2" value="����" class="input"> 
  </td>
    </tr>
  </form>
</table>
<%End If
rso.close
set rso=nothing %>
</body>
</html>
