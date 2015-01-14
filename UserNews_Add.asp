<!--#include file="conn.asp"-->
<!--#include file="SqlIn.Asp"-->
<!--#include file=usercookies.asp-->
<!--#include file=top.asp-->
<!--#include file="Include/err.asp"-->
<%
'=====================================
'天天供求信息网站管理系统
'天天网络科技工作室出品
'客服中心http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>
<%if instr(request.servervariables("http_referer"),"http://"&request.servervariables("http_host") )<1 then 
	response.write "<br><br><div align='center'>禁止从外部发布文章</div>"
	response.End()
end If
if lmkg2="1" then
call errdick(50)
response.end
end if
Call OpenConn
vip=conn.Execute("Select vip From DNJY_user Where username='"&request.cookies("dick")("username")&"'")(0)
sdname=conn.Execute("Select sdname From DNJY_user Where username='"&request.cookies("dick")("username")&"'")(0)
Dim rsnews
Set rsnews=Server.CreateObject("ADODB.RecordSet") 
sql="select id from DNJY_userNews where username='"&request.cookies("dick")("username")&"'"
rsnews.Open sql,conn,1,1
If vip="0" Then
response.write "<br><br><div align='center'>您不是VIP会员，不能发布文章！<a href='user.asp'>返回</a></div>"
rsnews.close
set rsnews=Nothing
Call CloseDB(conn)
response.end
End If
If usernews>0 And rsnews.recordcount>=usernews Then
response.write "<br><br><div align='center'>您发布文章已达到限定数量，不能再发布！<a href='user.asp'>返回</a></div>"
rsnews.close
set rsnews=Nothing
Call CloseDB(conn)
response.end
End If
rsnews.close
set rsnews=nothing%>
<html>
<head>
<link rel="stylesheet" type="text/css" href="css/css.css">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>添加信息</title>
<script language="JavaScript">
<!--


	function checkForm(){
	
		if (document.postart.ClassName.value=="") {
			alert("请选择类别！");
			document.postart.ClassName.focus();
			return false;
		}
		
		if (document.postart.txttitle.value.length == 0) {
			alert("请输入标题！");
			document.postart.txttitle.focus();
			return false;
		}
        var editor=KindEditor.create('#editor');
        if (editor.isEmpty())
	    {
	    alert("文章内容不能为空！")	  
	     return false
	     }
	}
	function doSubmit(){
		document.postart.submit();
	}
	//-->
</script>
<style type="text/css">
<!--
.STYLE2 {color: #FF0000}
-->
</style>
<!--kindeditor编辑器-->
	<link rel="stylesheet" href="../KindEditor/themes/default/default.css" />
	<link rel="stylesheet" href="../KindEditor/plugins/code/prettify.css" />
	<script charset="utf-8" src="../KindEditor/kindeditor.js"></script>
	<script charset="utf-8" src="../KindEditor/lang/zh_CN.js"></script>
	<script charset="utf-8" src="../KindEditor/plugins/code/prettify.js"></script>
	<script>
		KindEditor.ready(function(K) {
			var editor = K.create('textarea[name="d_content"]', {
				cssPath : '../KindEditor/plugins/code/prettify.css',
				uploadJson : '../KindEditor/asp/upload_json.asp?action=usernews',
				afterBlur: function(){this.sync();},//失去焦点执行this获得值
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
<!--kindeditor编辑器END-->
</head>
<body>
<div align="center">
<%Dim ClassName
ClassName=ReplaceUnsafe(Request.QueryString("ClassName"))
%>
<form method="POST" action="UserNews_AddSave.asp" name="postart" onSubmit="return checkForm()">
  <table width="980" border="0" align="center" cellpadding="0" cellspacing="0" bordercolor="#111111" style="border-collapse: collapse">
    <tr>
      <td width="214" height="356" valign="top" bgcolor="#FFFFFF"><div align="center"><!--#include file=userleft.asp--></div></td>
	  <td width="5">&nbsp;</td>
      <td width="860" align="center" valign="top" bgcolor="#FFFFFF">
	<!---上部通栏广告开始---->
  <table align="center" bgcolor="#ffffff">
  <tr>  
   <td ><script language=javascript src="<%=adjs_path("ads/js","fygg",c1&"_"&c2&"_"&c3)%>"></script></td>
  </tr>
  </table>
  <!---上部通栏广告结束----> 
           
             <table border="0" align="center" cellspacing="1" class="list">
                <tr> 
                  <th colspan="3">添 加 信 息</th>
                </tr>
                <tr>
                  <td>&nbsp;信息类别：&nbsp; </td>
                  <td colspan="2">		<%

set rs=conn.execute("select * from DNJY_city where id>0 and twoid>0 and threeid=0")
%>
          <SCRIPT language = "JavaScript">
var onecount;
onecount=0;
subcat = new Array();
        <%
	dim count5:count5 = 0
        do while not rs.eof 
        %>
subcat[<%=count5%>] = new Array("<%=rs("city")%>","<%=rs("id")%>","<%=rs("twoid")%>");
        <%
        count5 = count5 + 1
        rs.movenext
        loop
Call CloseRs(rs)
        %>
onecount=<%=count5%>;
          </SCRIPT>
          <%
set rs=conn.execute("select * from DNJY_city where id>0 and twoid>0 and threeid>0")
%>
          <SCRIPT language = "JavaScript">
var onecount4;
onecount4=0;
subcat4 = new Array();
        <%
		dim count4:count4 = 0
        do while not rs.eof 
        %>
subcat4[<%=count4%>] = new Array("<%=rs("city")%>","<%=rs("id")%>","<%=rs("twoid")%>","<%=rs("threeid")%>");
        <%
        count4 = count4 + 1
        rs.movenext
        loop
Call CloseRs(rs)
        %>
onecount4=<%=count4%>;

function changelocation(locationid)
    {
    document.postart.city_twoid.length = 0; 
	document.postart.city_twoid.options[0] = new Option('选择城市','');
	document.postart.city_threeid.length = 0; 
	document.postart.city_threeid.options[0] = new Option('选择城市','');
    var locationid=locationid;
    var i;
    for (i=0;i < onecount; i++)
        {
            if (subcat[i][1] == locationid)
            { 
                document.postart.city_twoid.options[document.postart.city_twoid.length] = new Option(subcat[i][0], subcat[i][2]);
            }        
        }
        
    }    
	
	function changelocation4(locationid,locationid1)
    {
    document.postart.city_threeid.length = 0; 
    document.postart.city_threeid.options[0] = new Option('选择城市','');
    var locationid=locationid;
	 var locationid1=locationid1;
    var i;
    for (i=0;i < onecount4; i++)
        {
            if (subcat4[i][2] == locationid)
            { 
			if (subcat4[i][1] == locationid1)
			{
                document.postart.city_threeid.options[document.postart.city_threeid.length] = new Option(subcat4[i][0], subcat4[i][3]);
            }        
        }
       } 
    }            
              </SCRIPT>
          <SELECT name="city_oneid" size="1" id="select" onChange="changelocation(document.postart.city_oneid.options[document.postart.city_oneid.selectedIndex].value)">
            <OPTION value="" selected>选择地区</OPTION>
            <%set rs=conn.execute("select * from DNJY_city where id>0 and twoid=0 order by indexid asc")
if rs.eof or rs.bof then
response.write "<option value=''>没有分类</option>"
else
do until rs.eof
response.write "<option value='"&rs("id")&"'>"&rs("city")&"</option>"
    rs.movenext
    loop
end if
Call CloseRs(rs)
%>
          </SELECT>
          <SELECT name="city_twoid" id="select6" onChange="changelocation4(document.postart.city_twoid.options[document.postart.city_twoid.selectedIndex].value,document.postart.city_oneid.options[document.postart.city_oneid.selectedIndex].value)">
            <OPTION value="" selected>选择城市</OPTION>
          </SELECT>
          <SELECT name="city_threeid" id="select7">
            <OPTION value="" selected>选择城市</OPTION>
          </SELECT>
<font color="#ff0000"> *</font>
		  </td>
                </tr>			
                <tr>
                  <td>&nbsp;信息类别：&nbsp; </td>
                  <td colspan="2">
				  <select name="ClassName" size="1" id="ClassName" >
            <option value="" selected>选择类别</option>
			<%set rs=conn.execute("select * From DNJY_userNewsClass")
			if rs.eof or rs.bof then
				response.write "<option value=''>没有分类</option>"
			else
				do until rs.eof
			%>
			<option value="<%=rs("classname")%>" <%if rs("ClassName")=ClassName then%>selected<%end if%>><%=rs("ClassName")%></option> 
			<%  rs.movenext
				loop
			end if
Call CloseRs(rs)
Call CloseDB(conn)
			%>
          </select>
				  <span class="STYLE2">*</span> </td>
                </tr>
                <tr> 
                  <td width="10%" align="right"><div align="center">标&nbsp;&nbsp;&nbsp; 题：</div></td>
                  <td  colspan="2"><input name="txttitle" id=me type="TEXT" size=60>
                    <span class="STYLE2">*</span></td>
                </tr>	
				
                <tr> 
                  <td width="10%" align="right" valign="top"  ><div align="center">关 键 字：</div></td>
                  <td colspan=2 ><input name="keywords" id="keywords" type="TEXT" size=60>
                    (如留空则关键字为:标题_网站名)</td>
                </tr>
                <tr>
                  <td align="right" valign="middle"  ><div align="center">描述信息：</div></td>
                  <td colspan=2 ><textarea name="Descrip" cols="75" rows="3" id="Descrip"></textarea>
                  (如留空则描述信息为文章前150字符)</td>
                </tr>
				<tr> 
                  <td width="10%"><div align="center">作&nbsp;&nbsp;&nbsp; 者：</div></td>
                  <td  colspan="2"><input name="author"  type="TEXT" size=30 value="<% =username%>"> 
                  &lt;=<%If session("author")<>"" Then response.write "【<FONT color=blue><FONT style=""CURSOR: hand"" onclick=""author.value='"&session("author")&"'"" color=#993300>"&session("author")&"</FONT></FONT>】"%>【<FONT 
				  color=blue><FONT style="CURSOR: hand" onclick='author.value="未知"' 
				  color=#993300>未知</FONT></FONT>】【<FONT color=blue><FONT 
				  style="CURSOR: hand" onclick="author.value='佚名'" 
				  color=#993300>佚名</FONT></FONT>】【<FONT color=blue><FONT 
				  style="CURSOR: hand" onclick="author.value='站长'" 
				  color=#993300>站长</FONT></FONT>】 </td>
                </tr>
				<tr> 
                  <td width="10%"><div align="center">来&nbsp;&nbsp;&nbsp; 源：</div></td>
                  <td  colspan="2"><input name="origin"  type="TEXT" size=30 value="<%=sdname%>">
                  &lt;=<%If session("origin")<>"" Then response.write "【<FONT color=blue><FONT style=""CURSOR: hand"" onclick=""origin.value="""&session("origin")&""" color=#993300>"&session("origin")&"</FONT></FONT>】"%>【<FONT 
				  color=blue><FONT style="CURSOR: hand" onclick="origin.value='不详'" 
				  color=#993300>不详</FONT></FONT>】【<FONT color=blue><FONT 
				  style="CURSOR: hand" onclick="origin.value='本站'" 
				  color=#993300>本站</FONT></FONT>】【<FONT color=blue><FONT 
				  style="CURSOR: hand" onclick="origin.value='互联网'" 
				  color=#993300>互联网</FONT></FONT>】</td>
                </tr>
 				<tr> 
                  <td width="10%"><div align="center">是否共享：</div></td>
                  <td  colspan="2"> <input type="radio" name="newsShared" value="0" id="newsShared" checked>
                 共享&nbsp;&nbsp;&nbsp;&nbsp;
                <input type="radio" name="newsShared" value="1" id="newsShared">
                不共享 (设置共享后可在总站相关栏目显示)
                </td>
                </tr>               
                <tr> 
                  <td height="25" colspan=3 align="left"><textarea id="editor" name="d_content" style="width:670px;height:400px;"></textarea><input type="checkbox" name="T_SaveImg" value="1" /> 保存内容中远程图片到本地&nbsp;&nbsp;<input type="checkbox" name="AspJpeg" value="1" checked/>给上传图片加水印<br></td>
                </tr>
                <tr> 
                  <td height="20" colspan=3 align=center > 
                    <input style="height:20px;" name="cmdok" type="submit" class="button" value=" 添 加 " >
                    &nbsp; 
                    <input style="height:20px;" name="cmdcancel" type="reset" class="button" value=" 重 写 " >
&nbsp;&nbsp;                  </td>
                </tr>
	</table>
	</td>
	  </tr>
    </table>
</form>
</div>
