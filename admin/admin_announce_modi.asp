<!--#include file="../conn.asp"-->
<!--#include file="../config.asp"-->
<!--#include file=cookies.asp-->
<!--#include file="../Include/Function.asp"-->
<%
'=====================================
'天天供求信息网站管理系统
'天天网络科技工作室出品
'客服中心http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
call checkmanage("06")%>`

<HEAD>
<META HTTP-EQUIV="ContentCH-Type" ContentCH="text/html; charset=gb2312" />
<TITLE>公告管理</TITLE>
<link rel="stylesheet" href="../css/manage.css" type="text/css">
<script language = "JavaScript">
function CheckForm()
{
	if (document.editForm.title.value.length == 0) {
		alert("标题没有填写.");
		document.editForm.title.focus();
		return false;}
	if (document.editForm.content.value.length == 0) {
		alert("内容没有填写.");
		document.editForm.content.focus();
		return false;}	 
 	return true;
}
</script>
</HEAD>

<BODY>
<%Dim city_oneid,city_twoid,city_threeid,city_one,city_two,city_three,Content,rso,Solidtop,activity
Call OpenConn

If request("del")="yes" Then'删除文章
Dim str2,rsdel,sqldel,i
Server.ScriptTimeout=50000
id=trim(request("DelID"))
if trim(id)="" then
response.write "<script language=JavaScript>" & chr(13) & "alert('没有选择记录！');" &"window.location='admin_announce.asp'" & "</script>"
response.end
end if
str2=split(id,",")
for i=0 to ubound(str2)
set rsdel=server.createobject("adodb.recordset")
sqldel="delete from [DNJY_announce] where id="&cstr(str2(i))
rsdel.open sqldel,conn,1,3
Next'循环结束
response.write "<script language=JavaScript>" & chr(13) & "alert('删除成功！');" &"window.location='admin_announce.asp'" & "</script>"
response.end
End If'删除END

'得出地区名
city_oneid=strint(request.form("city_one"))
city_twoid=strint(request.form("city_two"))
city_threeid=strint(request.form("city_three"))
if city_oneid>0 then
set rs=server.createobject("adodb.recordset")
rs.open "select city from DNJY_city where twoid=0 and threeid=0 and id="&city_oneid,conn,1,1
city_one=rs("city")
rs.close
if city_twoid>0 then
rs.open "select city from DNJY_city where id="&city_oneid&" and threeid=0 and twoid="&city_twoid,conn,1,1
city_two=rs("city")
rs.close
end if
if city_twoid>0 and city_threeid>0 then
rs.open "select city from DNJY_city where id="&city_oneid&" and twoid="&city_twoid&" and threeid="&city_threeid,conn,1,1
city_three=rs("city")
rs.close
end if
set rs=nothing
end If

dim Result
Result=request.QueryString("Result")
dim ID
ID=request.QueryString("ID")
call NewsEdit()%>
<br>
<form name="editForm" method="post" action="admin_announce_modi.asp?Action=SaveEdit&Result=<%=Result%>&ID=<%=ID%>" onSubmit="return CheckForm()">
<table class="tableborder" width="95%" border="0" align="center" cellpadding="3" cellspacing="1">  
    <tr>
      <th height="22" colspan="2">【<%If Result = "Add" then%>添加<%ElseIf Result = "Modify" then%>修改<%End If%>公告】</th>
    </tr>
	    <tr> 
      <td align="right">所属地区：</td>
      <td colspan="2" valign="top" bgcolor="#E3EBF9">　 
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
    document.editForm.city_two.length = 0; 
	document.editForm.city_two.options[0] = new Option('选择城市','');
	document.editForm.city_three.length = 0; 
	document.editForm.city_three.options[0] = new Option('选择城市','');
    var locationid=locationid;
    var i;
    for (i=0;i < onecount; i++)
        {
            if (subcat[i][1] == locationid)
            { 
                document.editForm.city_two.options[document.editForm.city_two.length] = new Option(subcat[i][0], subcat[i][2]);
            }        
        }
        
    }    
	
	function changelocation4(locationid,locationid1)
    {
    document.editForm.city_three.length = 0; 
    document.editForm.city_three.options[0] = new Option('选择城市','');
    var locationid=locationid;
	 var locationid1=locationid1;
    var i;
    for (i=0;i < onecount4; i++)
        {
            if (subcat4[i][2] == locationid)
            { 
			if (subcat4[i][1] == locationid1)
			{
                document.editForm.city_three.options[document.editForm.city_three.length] = new Option(subcat4[i][0], subcat4[i][3]);
            }        
        }
       } 
    }            
              </SCRIPT>
      <SELECT name="city_one" size="1" id="select2" onChange="changelocation(document.editForm.city_one.options[document.editForm.city_one.selectedIndex].value)">
            <OPTION value="" selected>选择城市</OPTION>
  <%set rsi=conn.execute("select * from DNJY_city where id>0 and twoid=0 and threeid=0 order by indexid asc")
if rsi.eof or rsi.bof then
response.write "<option value=''>没有分类</option>"
else
do until rsi.eof%>
  <OPTION value="<%=rsi("id")%>" <%if rsi("id")=city_oneid then%>selected<%end if%>><%=rsi("city")%></OPTION>
 <%rsi.movenext
    loop
	end if
rsi.close
set rsi = nothing%>
      </SELECT> 
	  <SELECT name="city_two" id="city_two" onChange="changelocation4(document.editForm.city_two.options[document.editForm.city_two.selectedIndex].value,document.editForm.city_one.options[document.editForm.city_one.selectedIndex].value)">
    <OPTION value="" selected>选择城市</OPTION>
<%set rsi=conn.execute("select * from DNJY_city where id="&city_oneid&" and twoid>0 and threeid=0")
if rsi.eof and rsi.bof then
response.write "<option value=''>没有分类</option>"
else
do until rsi.eof%>
  <OPTION value="<%=rsi("twoid")%>" <%if rsi("twoid")=city_twoid then%>selected<%end if%>><%=rsi("city")%></OPTION>
 <%rsi.movenext
    loop
	end if
rsi.close
set rsi = nothing%>
	</SELECT>
	     <SELECT name="city_three" id="city_three">
         <OPTION value="" selected>选择城市</OPTION>
<%set rsi=conn.execute("select * from DNJY_city where id="&city_oneid&" and twoid="&city_twoid&" and threeid<>0")
if rsi.eof and rsi.bof then
response.write "<option value=''>没有分类</option>"
else
do until rsi.eof%>
<OPTION value="<%=rsi("threeid")%>" <%if rsi("threeid")=city_threeid then%>selected<%end if%>><%=rsi("city")%></OPTION>
   <%rsi.movenext
    loop
 end if
rsi.close
set rsi = nothing%>
    </SELECT></td>
    </tr>
    <tr>
      <td width="20%" align="right">公告标题：</td>
      <td width="80%"><input name="title" type="text" id="title" style="width: 280" value="<%=title%>" maxlength="100"> <font color="red">*</font> </td>
    </tr>
 
   <tr>
      <td align="right">公告内容：</td>
      <td><textarea id="content" name="content" style="width:670px;height:200px;"><%=content%></textarea></td>
    </tr>
      <tr>
        <td align="right">固&nbsp;&nbsp;&nbsp;&nbsp;顶：</td>
        <td><input name="Solidtop" type="checkbox" style="HEIGHT: 13px;WIDTH: 13px;" value="1" <%if Solidtop=1 then response.write ("checked")%>>
          &nbsp;固顶&nbsp;<input name="activity" type="checkbox" value="1" style="HEIGHT: 13px;WIDTH: 13px;" <%if activity=1 then response.write ("checked")%>>&nbsp;显示</td>
      </tr>
	<tr>
      <td align="right"></td>
      <td><input name="submitSaveEdit" type="submit" id="submitSaveEdit" value="保存"> <input type="button" value="返回上一页" onClick="history.back(-1)"></td>
    </tr>
  </form>
</table>
</BODY>
</HTML>
<%
sub NewsEdit()
  dim Action,rsRepeat,rs,sql
  Action=request.QueryString("Action")
  if Action="SaveEdit" then'保存

    set rs = server.createobject("adodb.recordset")

     if Result="Add" then'----------------添加
	  sql="select * from DNJY_announce"
      rs.open sql,conn,1,3
      rs.addnew
      if city_twoid="" then city_twoid=0
      if city_threeid="" then city_threeid=0
      rs("city_oneid")=city_oneid
      rs("city_one")=city_one
      rs("city_two")=city_two
      rs("city_twoid")=city_twoid
      rs("city_three")=city_three
      rs("city_threeid")=city_threeid
      rs("title")=trim(Request.Form("title"))
	  rs("content")=trim(Request.Form("content"))
	  If Request.Form("Solidtop")=1 then
	  rs("Solidtop")=Request.Form("Solidtop")
	  else
      rs("Solidtop")=0
	  end If
      If Request.Form("activity")=1 then
	  rs("activity")=Request.Form("activity")
	  else
      rs("activity")=0
	  end If
	  rs("addtime")=now()
	  rs.update
	  Call CloseRs(rs)
	  set rs=server.createobject("adodb.recordset")
	  sql="select top 1 ID from DNJY_announce order by ID desc"
	  rs.open sql,conn,1,1
	  ID=rs("ID")
      Call CloseRs(rs)
      Call Alert ("添加成功！","admin_announce.asp")
	end if'----------------添加END

	if Result="Modify" then'----------------修改
      sql="select * from DNJY_announce where ID="&ID
      rs.open sql,conn,1,3
      if city_twoid="" then city_twoid=0
      if city_threeid="" then city_threeid=0
      rs("city_oneid")=city_oneid
      rs("city_one")=city_one
      rs("city_two")=city_two
      rs("city_twoid")=city_twoid
      rs("city_three")=city_three
      rs("city_threeid")=city_threeid
      rs("title")=trim(Request.Form("title"))
	  rs("content")=trim(Request.Form("content"))
	  If Request.Form("Solidtop")=1 then
	  rs("Solidtop")=Request.Form("Solidtop")
	  else
      rs("Solidtop")=0
	  end If
      If Request.Form("activity")=1 then
	  rs("activity")=Request.Form("activity")
	  else
      rs("activity")=0
	  end If
	  rs.update
      Call CloseRs(rs)
      Call Alert ("修改成功！","admin_announce.asp")
	end If'----------------修改END
  else'编辑
	if Result="Modify" then
      set rs = server.createobject("adodb.recordset")
      sql="select * from DNJY_announce where ID="& ID
      rs.open sql,conn,1,1
      if rs.bof and rs.eof then
        response.write ("<center>数据库记录读取错误！</center>")
        response.end
      end if
	  city_oneid=rs("city_oneid")	  
	  city_one=rs("city_one")
	  city_twoid=rs("city_twoid")
	  city_two=rs("city_two")
	  city_threeid=rs("city_threeid")
	  city_three=rs("city_three")
	  title=rs("title")
	  content=rs("content")
	  Solidtop=rs("Solidtop")
	  activity=rs("activity")
      Call CloseRs(rs)
    end if
  end If  
end Sub

%>