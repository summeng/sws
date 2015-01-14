<!--#include file="../conn.asp"-->
<!--#include file="../SqlIn.Asp"-->
<!--#include file="../config.asp"-->
<!--#include file="../Include/head.asp"-->
<!--#include file=../Include/Function.asp-->
<!--#include file=cookies.asp-->
<%
'=====================================
'天天供求信息网站管理系统
'天天网络科技工作室出品
'客服中心http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>
<%call checkmanage("04")%>
<%
Function inHTML(str)
	Dim sTemp
	sTemp = str
	inHTML = ""
	If IsNull(sTemp) = True Then
		Exit Function
	End If
	sTemp = Replace(sTemp, "&", "&amp;")
	sTemp = Replace(sTemp, "<", "&lt;")
	sTemp = Replace(sTemp, ">", "&gt;")
	sTemp = Replace(sTemp, Chr(34), "&quot;")
	inHTML = sTemp
End Function	
Dim isse,page,ClassName,se_ClassName,searchkeys,selecttype,Descripm,Descrip,Content,addtime,author,origin,SavePathFileName,newsShared,ifshow,newstop,tj,rsnew
isse=ReplaceUnsafe(request.QueryString("isse"))
page=ReplaceUnsafe(Request.QueryString("page"))
ClassName=ReplaceUnsafe(Request.QueryString("ClassName"))
se_ClassName=ReplaceUnsafe(Request.QueryString("se_ClassName"))
searchkeys=ReplaceUnsafe(Request.QueryString("searchkeys"))
ID=ReplaceUnsafe(Request.QueryString("id"))
selecttype=ReplaceUnsafe(Request.QueryString("selecttype"))

set rsnew = server.CreateObject ("Adodb.recordset")
sql="select * from DNJY_userNews where ID="& id
rsnew.Open sql,conn,1,1
city_oneid=rsnew("city_oneid")
city_twoid=rsnew("city_twoid")
city_threeid=rsnew("city_threeid")
ClassName=rsnew("ClassName")
keywords = rsnew("keywords")
Descrip=rsnew("Descrip")
Title=rsnew("title")
Content=trim(rsnew("Content"))
If IsNull(rsnew("Content")) Then Content="暂无内容..."
addtime=rsnew("addtime")
author=rsnew("author")
origin=rsnew("origin")
SavePathFileName = rsnew("SavePathFileName")
newsShared=rsnew("newsShared")
tj=rsnew("tj")
newstop=rsnew("newstop")
ifshow=rsnew("ifshow")
rsnew.close
set rsnew=nothing

Dim aSavePathFileName
if SavePathFileName<>"" then 
' 把带"|"的字符串转为数组，用于初始下拉框表单
	aSavePathFileName = Split(SavePathFileName, "|")
' 函数InitSelect，根据数组值及初始值返回下拉框输出字串，具体请见startup.asp文件中函数的说明部分
else
	SavePathFileName=""
	aSavePathFileName = Split(SavePathFileName, "|")
	
end if
	
%>
<html>
<head>
<link rel="stylesheet" type="text/css" href="../css/css.css">
<style type="text/css">
<!--
body {
	background-color: #E3EBF9;
}
.style1 {color: #FFFFFF;
	font-weight: bold;
}
.style2 {color: #FFFFFF}
-->
</style>

<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>修改信息</title>
<script language="JavaScript">
	<!--
	
	 function   checkDate(s)   
    {   
            var   isOk   =   false;   
  			tempArray   =   s.split('-');   
      		if   (tempArray.length   ==   3)   
    	    if   (   parseInt(tempArray[0]).toString().length   ==   4)   
            if   (   parseInt(tempArray[1])   >=1   &&   parseInt(tempArray[1])   <=12)   
     		if   (   parseInt(tempArray[2])   >=1   &&   parseInt(tempArray[2])   <=   31)   
            isOk   =   true;   
    
           return isOk;  
	 }   
   
	function checkForm(){
		
		//if (checkDate(document.postart.addtime.value)==false)
		//{alert("请正确填写日期（例如:2008-1-23）！");
		//return false;
		//}
		 
		if (document.postart.ClassName.value=="") {
			alert("请选择栏目！");
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
<form method="POST" action="UserNews_ModifySave.asp?isse=<%=isse%>&id=<%=id%>&page=<%=page%>&searchkeys=<%=searchkeys%>&selecttype=<%=selecttype%>&se_ClassName=<%=se_ClassName%>" name="postart" onSubmit="return checkForm()">
	
	<input type=hidden name=d_savepathfilename  value="<%=SavePathFileName%>">
            
     <table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#ADAED6" width="70%">
	<TR> 
    <TD height="5" colspan="8" bgcolor="#799AE1" align="center"><FONT color="#FFFFFF" style="font-size:14px">修改信息--该信息栏目ID为：<b><%=ID%></FONT></TD>
    </TR>

      <tr> 
                  <td width="10%" align="right" valign="middle"><div align="center">所属地区：</td>
      <td colspan="2" valign="top" >　 
<%Dim rsi,count
set rsi=conn.execute("select * from DNJY_city where id>0 and twoid>0 and threeid=0")
%>
      <SCRIPT language = "JavaScript">
var onecount;
onecount=0;
subcat = new Array();
        <%
		//dim count:
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
    document.addNEWS.city_two.length = 0; 
	document.addNEWS.city_two.options[0] = new Option('选择城市','');
	document.addNEWS.city_three.length = 0; 
	document.addNEWS.city_three.options[0] = new Option('选择城市','');
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
    document.addNEWS.city_three.options[0] = new Option('选择城市','');
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
            <OPTION value="" selected>选择城市</OPTION>
  <%set rsi=conn.execute("select * from DNJY_city where id>0 and twoid=0 and threeid=0 order by indexid asc")
if rsi.eof or rsi.bof then
response.write "<option value=''>没有分类</option>"
else
do until rsi.eof%>
  <OPTION value="<%=rsi("id")%>" <%if rsi("id")=city_oneid then%>selected<%end if%>><%=rsi("city")%></OPTION>
 <%rsi.movenext
    loop
	%>
	<%end if
rsi.close
set rsi = nothing
%>
      </SELECT> 
	  <SELECT name="city_two" id="city_two" onChange="changelocation4(document.addNEWS.city_two.options[document.addNEWS.city_two.selectedIndex].value,document.addNEWS.city_one.options[document.addNEWS.city_one.selectedIndex].value)">
    <OPTION value="" selected>选择城市</OPTION>
   <%
set rsi=conn.execute("select * from DNJY_city where id="&city_oneid&" and twoid>0 and threeid=0")
if rsi.eof and rsi.bof then
response.write "<option value=''>没有分类</option>"
else
do until rsi.eof%>
  <OPTION value="<%=rsi("twoid")%>" <%if rsi("twoid")=city_twoid then%>selected<%end if%>><%=rsi("city")%></OPTION>
 <%rsi.movenext
    loop
	%>
	<%end if
rsi.close
set rsi = nothing
%>
	</SELECT>
	     <SELECT name="city_three" id="city_three">
         <OPTION value="" selected>选择城市</OPTION>
		<%
set rsi=conn.execute("select * from DNJY_city where id="&city_oneid&" and twoid="&city_twoid&" and threeid<>0")
if rsi.eof and rsi.bof then
response.write "<option value=''>没有分类</option>"
else
do until rsi.eof%>
<OPTION value="<%=rsi("threeid")%>" <%if rsi("threeid")=city_threeid then%>selected<%end if%>><%=rsi("city")%></OPTION>
   <% rsi.movenext
    loop
	%>
<%end if
rsi.close
set rsi = nothing
%>
    </SELECT></td>
    </tr>
                <tr> 
                  <td width="10%" align="right" valign="middle"><div align="center">信息栏目：</div></td>
                 
                  <td >

          <select name="ClassName" size="1" id="ClassName">
            <option value="" selected>选择栏目</option>
<%set rs=conn.execute("select * From DNJY_userNewsClass")
if rs.eof or rs.bof then
	response.write "<option value=''>没有分类</option>"
else
	do until rs.eof
%>
<option value="<%=rs("classname")%>" <%if rs("ClassName")=ClassName then%>selected<%end if%>><%=rs("ClassName")%></option> 
<%   rs.movenext
     loop
	
end if
rs.close
set rs = nothing
%>
          </select>
          <span class="STYLE2">*</span></td>
                </tr>
                
                
                <tr> 
                  <td width="10%" align="right"  valign="middle" ><div align="center">标&nbsp;&nbsp;&nbsp; 题：</div></td>
                  <td  colspan="2"><input name="txttitle" type="TEXT" id="txttitle"  value="<%=inHTML(Title)%>" size=60>
                  <span class="STYLE2">*</span></td>
                </tr>
                <tr>
                  <td align="right"  valign="middle" ><div align="center">关 键 字：</div></td>
                  <td  colspan="2"><input name="keywords" type="TEXT" id="keywords"  value="<%=keywords%>" size=60>
                  (如留空则关键字为:标题_网站名)</td>
                </tr>
				<tr>
                  <td align="right" valign="middle"  ><div align="center">描述信息：</div></td>
                  <td colspan=2 ><textarea name="Descrip" cols="75" rows="3" id="Descrip"><%=Descrip%></textarea>
                  (如留空则描述信息为文章前150字符)</td>
                </tr>
   				<tr> 
                  <td width="10%"><div align="center">作&nbsp;&nbsp;&nbsp; 者：</div></td>
                  <td  colspan="2"><input name="author"  type="TEXT" value="<%=author%>" size=20>
                  &lt;=【<FONT 
      color=blue><FONT style="CURSOR: hand" onclick='author.value="未知"' 
      color=#993300>未知</FONT></FONT>】【<FONT color=blue><FONT 
      style="CURSOR: hand" onclick="author.value='佚名'" 
      color=#993300>佚名</FONT></FONT>】【<FONT color=blue><FONT 
      style="CURSOR: hand" onclick="author.value='站长'" 
      color=#993300>站长</FONT></FONT>】 </td>
                </tr>
				<tr> 
                  <td width="10%"><div align="center">来&nbsp;&nbsp;&nbsp; 源：</div></td>
                  <td  colspan="2"><input name="origin" type="TEXT" value="<%=origin%>" size=20>
                  &lt;=【<FONT 
      color=blue><FONT style="CURSOR: hand" onclick="origin.value='不详'" 
      color=#993300>不详</FONT></FONT>】【<FONT color=blue><FONT 
      style="CURSOR: hand" onclick="origin.value='本站'" 
      color=#993300>本站</FONT></FONT>】【<FONT color=blue><FONT 
      style="CURSOR: hand" onclick="origin.value='互联网'" 
      color=#993300>互联网</FONT></FONT>】</td>
                </tr >
      			<tr> 
                  <td width="10%"><div align="center">固&nbsp;&nbsp;&nbsp;&nbsp;顶：</div></td>
                  <td  colspan="2"> <%if newstop=1 then%>               
                <input type="radio" name="newstop" value="1" id="newstop" checked>
                 固顶&nbsp;&nbsp;&nbsp;&nbsp;
                <input type="radio" name="newstop" value="0" id="newstop">
                不固顶
                <%else%>                         
                <input type="radio" name="newstop" value="1" id="newstop">
                 固顶&nbsp;&nbsp;&nbsp;&nbsp;
                <input type="radio" name="newstop" value="0" id="newstop" checked>
                不固顶<%end if%>  
                </td>
                </tr>  
      			<tr> 
                  <td width="10%"><div align="center">推&nbsp;&nbsp;&nbsp;&nbsp;荐：</div></td>
                  <td  colspan="2"> <%if tj=True then%>               
                <input type="radio" name="tj" value="True" id="tj" checked>
                 推荐&nbsp;&nbsp;&nbsp;&nbsp;
                <input type="radio" name="tj" value="false" id="tj">
                不推荐
                <%else%>                         
                <input type="radio" name="tj" value="True" id="tj">
                 推荐&nbsp;&nbsp;&nbsp;&nbsp;
                <input type="radio" name="tj" value="false" id="tj" checked>
                不推荐<%end if%>
                </td>
                </tr>  
      			<tr> 
                  <td width="10%"><div align="center">共享审核：</div></td>
                  <td  colspan="2"> <%if ifshow=1 then%>               
                <input type="radio" name="ifshow" value="1" id="ifshow" checked>
                 同意&nbsp;&nbsp;&nbsp;&nbsp;
                <input type="radio" name="ifshow" value="0" id="ifshow">
                不同意
                <%else%>                         
                <input type="radio" name="ifshow" value="1" id="ifshow">
                 同意&nbsp;&nbsp;&nbsp;&nbsp;
                <input type="radio" name="ifshow" value="0" id="ifshow" checked>
                不同意<%end if%>  (通过审核后才可在总站相关栏目显示)
                </td>
                </tr>                     
                
                <tr><td width="10%"><div align="center">资讯内容：</div></td> 
                  <td height="25" colspan=3 align="left" ><textarea id="editor" name="d_content" style="width:670px;height:400px;"><%=Content%></textarea><input type="checkbox" name="T_SaveImg" value="1" /> 保存内容中远程图片到本地&nbsp;&nbsp;<input type="checkbox" name="AspJpeg" value="1" checked/>给上传图片加水印<br></td>
                </tr>
                <tr>
                  <td height="20" align=center ><div align="center">发布日期：</div></td>
                  <td height="20" colspan="2"  ><input name="addtime" type="TEXT" size=25 maxlength=20  value="<%=addtime%>">
&lt;=用当前日期：<FONT color=blue><FONT 
      style="CURSOR: hand" onclick="addtime.value='<%=now()%>'" 
      color=#993300><%=now()%></FONT></FONT> </td>
                </tr>

                <tr> 
                  <td height="20" colspan=3 align=center > 
                    <input style="height:20px;" name="cmdok" type="submit" class="button" value=" 修 改 " >
                    &nbsp; 
                    <input style="height:20px;" name="cmdcancel" type="reset" class="button" value=" 重 写 " >
&nbsp;&nbsp;                  </td>
                </tr>
    </table>


</form>
</div>
<%
Call CloseRs(rs)
Call CloseDB(conn)
%>