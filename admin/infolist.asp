<!--#include file="../conn.asp"-->
<!--#include file="../SqlIn.Asp"-->
<!--#include file=cookies.asp-->
<!--#include file="../config.asp"-->
<!--#include file="../Include/head.asp"-->
<!--#include file="../Include/Function.asp"-->
<!--#include file="../Include/label.asp"-->
<%
'=====================================
'天天供求信息网站管理系统
'天天网络科技工作室出品
'客服中心http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>
<%call checkmanage("05")%>
<link rel="stylesheet" type="text/css" href="../css/css.css">
<BODY>
<style type="text/css">
<!--
body {
	background-color: #E3EBF9;
	margin-top: 16px;
}
-->
</style>
<body leftmargin="0">
<script language=javascript>
//弹出和管理回复开始
<!--
function left_menu(meval)
{
  var left_n=eval(meval);
  if (left_n.style.display=="none")
  { eval(meval+".style.display='';"); }
  else
  { eval(meval+".style.display='none';"); }
}
function DoEmpty(params)
{
if (confirm("真的要删除这条记录吗？删除后将无法恢复！"))
window.location = params ;
}
function test()
{
  if(!confirm('是否确定进行批量操作？操作后不能恢复！')) return false;
}
-->
</script>

<!---弹出面板开始--->
<script>
function loadThreadFollow(t_id,b_id){
	var targetImg =eval("document.all.followImg" + t_id);
	var targetDiv =eval("document.all.follow" + t_id);
	
	if ("object"==typeof(targetImg)){
		if (targetDiv.style.display!=''){
			targetDiv.style.display="";
			
		}else{
			targetDiv.style.display="none";
		}
	}
}
</script>
<!---弹出面板结束--->
<!---全选验证删除开始--->
<SCRIPT language=JavaScript>
function showoperatealert(id)
{
		if (id==1)
        {

		{
		  	thisForm.target='_self';
			thisForm.action="info_yz.asp?qxyz=yes";
			thisForm.submit();
		}
		}
		if (id==2)
        {

		{
		  	thisForm.target='_self';
			thisForm.action="info_yz.asp?qxyz=no";
			thisForm.submit();
		}
		}
		if (id==3)
        {

		{
		  	thisForm.target='_self';
			thisForm.action="info_xxtj.asp?tj=1";
			thisForm.submit();
		}
		}
		if (id==4)
        {

		{
		  	thisForm.target='_self';
			thisForm.action="info_xxtj.asp?tj=2";
			thisForm.submit();
		}
		}
		if (id==5)
        {

		{
		  	thisForm.target='_self';
			thisForm.action="info_tjdel.asp";
			thisForm.submit();
		}
		}		
		if (id==6)
        {
		{
			
		  	thisForm.target='_self';
			thisForm.action="info_del.asp?urlpage=infolist.asp";
			thisForm.submit();
		}
		}						
}

function CheckAll(form)
  {
  for (var i=0;i<form.elements.length;i++)
    {
    var e = form.elements[i];
    if (e.name != 'chkall' )
       e.checked = form.chkall.checked;
    }
  }
//-->
    </SCRIPT>

<!---全选验证删除结束--->
<div align="center">
  <center>
<table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="99%" height="1" bordercolor="#ADAED6">
  <tr>
    <td width="100%" height="24" colspan="14">
    <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%" height="25">
<TR>
    <TD height="20" bgcolor="#799AE1" align="center" colspan="14"><FONT color="#FFFFFF" style="font-size:14px">发 布 信 息 管 理</FONT></TD>
 </TR>
      <tr>
        <td width="200" height="24">
    <FORM name=sou1 action="?dick=1" method=POST>
    <p align="center">按帐号：<input type="text" name="key1" size="15">
    <input type="submit" value="查找" name="B1">
    </form>
        </td>
        <td width="200" height="24">
    <FORM name=sou2 action="?dick=2" method=POST>
    <p align="center">按姓名：<input type="text" name="key2" size="15">
    <input type="submit" value="查找" name="B1">
    </form>
        </td>
        <td width="360">
    <FORM name=sou3 action="?dick=3" method=POST>
    <p align="center">按关键词：<input type="text" name="key3" size="28">
	<select name="skey" style="BACKGROUND-COLOR: #FFCCFF; BORDER-BOTTOM: #FFFFFF 1px solid; BORDER-LEFT: #a2a2a2 1px solid; BORDER-RIGHT: #ffffff 1px solid; BORDER-TOP: #a2a2a2 1px solid; COLOR: #333333; FONT-SIZE: 9pt; HEIGHT: 19px;">
                            <option value="0">按标题</option>
                            <option value="1">按简介</option>                            
                          </select>
    <input type="submit" value="查找" name="B1">
    </form>
        </td>
      </tr>
	  <tr>
</table>
    <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%" height="25">
<%
dim k,dick,tj
dim ThisPage,Pagesize,Allrecord,Allpage

   dick=trim(request("dick"))
	city_oneid=Request("city_one")
	city_twoid=Request("city_two")
	city_threeid=Request("city_three")
   If city_oneid="" Then city_oneid=0
   If city_twoid="" Then city_twoid=0
   If city_threeid="" Then city_threeid=0

	IF city_oneid=0 then
	ttcc=""
	elseIF city_threeid>0 then
	ttcc=" and city_oneid="&city_oneid&" and city_twoid="&city_twoid&" and city_threeid="&city_threeid
	elseIF city_twoid>0 then
	ttcc=" and city_oneid="&city_oneid&" and city_twoid="&city_twoid
	Else
	ttcc=" and city_oneid="&city_oneid
	End If
	
	type_oneid=Request("type_one")
	type_twoid=Request("type_two")
	type_threeid=Request("type_three")
   If type_oneid="" Then type_oneid=0
   If type_twoid="" Then type_twoid=0
   If type_threeid="" Then type_threeid=0

	IF type_oneid=0 then
	tttt=""
	elseIF type_threeid>0 then
	tttt=" and type_oneid="&type_oneid&" and type_twoid="&type_twoid&" and type_threeid="&type_threeid
	elseIF type_twoid>0 then
	tttt=" and type_oneid="&type_oneid&" and type_twoid="&type_twoid
	Else
	tttt=" and type_oneid="&type_oneid
	End If%>
	<td height="25" width="20%">按地区类别查询：</td>
        <td width="50%" colspan="4" height="24">
 <form method="POST" name="form" id="form" language="javascript"  action="?t1=<%=trim(request("t1"))%>&t2=<%=trim(request("t2"))%>&t3=<%=trim(request("t3"))%>&dick=<%=dick%>">
<% 
set rs=server.createobject("adodb.recordset")
set rs=conn.execute("select * from DNJY_city where id>0 and twoid>0 and threeid=0")%>
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
          <SELECT name="city_one" size="1" id="select" onChange="changelocation(document.form.city_one.options[document.form.city_one.selectedIndex].value)">
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
          <SELECT name="city_two" id="select6" onChange="changelocation4(document.form.city_two.options[document.form.city_two.selectedIndex].value,document.form.city_one.options[document.form.city_one.selectedIndex].value)">
            <OPTION value="" selected>选择城市</OPTION>
          </SELECT>
          <SELECT name="city_three" id="select7">
            <OPTION value="" selected>选择城市</OPTION>
          </SELECT>
 <input class="inputb" name="I2" type="submit" style="BACKGROUND-COLOR: #eeeeee; BORDER-BOTTOM: #a2a2a2 1px solid; BORDER-LEFT: #ffffff 1px solid; BORDER-RIGHT: #a2a2a2 1px solid; BORDER-TOP: #ffffff 1px solid; COLOR: #333333; FONT-SIZE: 9pt; HEIGHT: 19px;"  value="查询" border="0" /></font>
</form> </td><%If city_oneid>0 then%><td width="30%" colspan="2" align="left"><%city_one=Conn.Execute("select city from DNJY_city where twoid=0 and id="&city_oneid)(0)
	If city_twoid>0 Then city_two=Conn.Execute("select city from DNJY_city where id="&city_oneid&" and threeid=0 and twoid="&city_twoid)(0)
	If city_threeid<>0 Then city_three=Conn.Execute("select city from DNJY_city where id="&city_oneid&" and twoid="&city_twoid&" and threeid="&city_threeid)(0)
	response.write ""&city_one&"/"&city_two&"/"&city_three&""%></td><%End if%>
      </tr>

<tr>
<form method="POST" name="myform" id="myform" language="javascript"  action="?t1=<%=trim(request("t1"))%>&t2=<%=trim(request("t2"))%>&t3=<%=trim(request("t3"))%>&dick=<%=dick%>">
<td height="25" width="20%">按信息类别查询：</td>
<td height="25" width="50%"><%
set rs=conn.execute("select * from DNJY_type where id>0 and twoid>0 and threeid=0 order by indexid")
%>
          <SCRIPT language = "JavaScript">
var onecount2;
onecount2=0;
subcat2 = new Array();
        <%
		dim count2:count2 = 0
        do while not rs.eof 
        %>
subcat2[<%=count2%>] = new Array("<%=rs("name")%>","<%=rs("id")%>","<%=rs("twoid")%>");
        <%
        count2 = count2 + 1
        rs.movenext
        loop
Call CloseRs(rs)
        %>
onecount2=<%=count2%>;
          </SCRIPT>
          <%
set rs=conn.execute("select * from DNJY_type where id>0 and twoid>0 and threeid>0 order by indexid")
%>
          <SCRIPT language = "JavaScript">
var onecount3;
onecount3=0;
subcat3 = new Array();
        <%dim count3:count3 = 0
        do while not rs.eof%>
subcat3[<%=count3%>] = new Array("<%=rs("name")%>","<%=rs("id")%>","<%=rs("twoid")%>","<%=rs("threeid")%>");
        <%
        count3 = count3 + 1
        rs.movenext
        loop
Call CloseRs(rs)
        %>
onecount3=<%=count3%>;

function changelocation2(locationid)
    {
    document.myform.type_two.length = 0; 
    document.myform.type_two.options[0] = new Option('二级分类','');
	document.myform.type_three.length = 0; 
    document.myform.type_three.options[0] = new Option('三级分类','');
    var locationid=locationid;
	var i;
    for (i=0;i < onecount2; i++)
        {
            if (subcat2[i][1] == locationid)
            { 
                document.myform.type_two.options[document.myform.type_two.length] = new Option(subcat2[i][0], subcat2[i][2]);
            }        
        }
       
    }    
     function changelocation3(locationid,locationid1)
    {
    document.myform.type_three.length = 0; 
    document.myform.type_three.options[0] = new Option('三级分类','');
    var locationid=locationid;
	 var locationid1=locationid1;
    var i;
    for (i=0;i < onecount3; i++)
        {
            if (subcat3[i][2] == locationid)
            { 
			if (subcat3[i][1] == locationid1)
			{
                document.myform.type_three.options[document.myform.type_three.length] = new Option(subcat3[i][0], subcat3[i][3]);
            }        
        }
       } 
    }            
	      </SCRIPT>
          <SELECT name="type_one" size="1" id="select8" onChange="changelocation2(document.myform.type_one.options[document.myform.type_one.selectedIndex].value)">
            <OPTION value="" selected>一级分类</OPTION>
            <%set rs=conn.execute("select * from DNJY_type where id>0 and twoid=0 order by indexid asc")
if rs.eof or rs.bof then
response.write "<option value=''>没有分类</option>"
else
do until rs.eof
response.write "<option value='"&rs("id")&"'>"&rs("name")&"</option>"
    rs.movenext
    loop
end if
Call CloseRs(rs)%>
          </SELECT>
          <SELECT name="type_two" id="select9" onChange="changelocation3(document.myform.type_two.options[document.myform.type_two.selectedIndex].value,document.myform.type_one.options[document.myform.type_one.selectedIndex].value)">
            <OPTION value="" selected>二级分类</OPTION>
          </SELECT>
          <SELECT name="type_three" id="select10">
            <OPTION value="" selected>三级分类</OPTION>
          </SELECT>  <input class="inputb" name="t2" type="submit" style="BACKGROUND-COLOR: #eeeeee; BORDER-BOTTOM: #a2a2a2 1px solid; BORDER-LEFT: #ffffff 1px solid; BORDER-RIGHT: #a2a2a2 1px solid; BORDER-TOP: #ffffff 1px solid; COLOR: #333333; FONT-SIZE: 9pt; HEIGHT: 19px;"  value="查询" border="0" /></font>
</form></td><%If type_oneid>0 then%><td width="30%" colspan="2" align="left"><%type_one=Conn.Execute("select name from DNJY_type where twoid=0 and id="&type_oneid)(0)
	If type_twoid>0 Then type_two=Conn.Execute("select name from DNJY_type where id="&type_oneid&" and threeid=0 and twoid="&type_twoid)(0)
	If type_threeid<>0 Then type_three=Conn.Execute("select name from DNJY_type where id="&type_oneid&" and twoid="&type_twoid&" and threeid="&type_threeid)(0)
	response.write ""&type_one&"/"&type_two&"/"&type_three&""%></td><%End if%>
                        </tr>
						

      <tr>
        <td width="100%" colspan="14" height="24">
        <p align="center"><b>显示：</b><a href="?dick=4&city_one=<%=city_oneid%>&city_two=<%=city_twoid%>&city_three=<%=city_threeid%>&type_one=<%=type_oneid%>&type_two=<%=type_twoid%>&type_three=<%=type_threeid%>" title='指尚未验证的信息'>未(<font color="#FF0000">ⅹ</font>)验证信息</a>---<a href="?dick=5&city_one=<%=city_oneid%>&city_two=<%=city_twoid%>&city_three=<%=city_threeid%>&type_one=<%=type_oneid%>&type_two=<%=type_twoid%>&type_three=<%=type_threeid%>" title='指已经验证的信息'>已(√)验证信息</a>---<a href="?dick=14&city_one=<%=city_oneid%>&city_two=<%=city_twoid%>&city_three=<%=city_threeid%>&type_one=<%=type_oneid%>&type_two=<%=type_twoid%>&type_three=<%=type_threeid%>" title='指已经置顶的信息'>已置顶信息</a>---<a href="?dick=6&city_one=<%=city_oneid%>&city_two=<%=city_twoid%>&city_three=<%=city_threeid%>&type_one=<%=type_oneid%>&type_two=<%=type_twoid%>&type_three=<%=type_threeid%>" title='指已经推荐的信息'>已推荐信息</a>---<a href="?dick=7&city_one=<%=city_oneid%>&city_two=<%=city_twoid%>&city_three=<%=city_threeid%>&type_one=<%=type_oneid%>&type_two=<%=type_twoid%>&type_three=<%=type_threeid%>" title='指有效期已到的信息'>已到期信息</a>---<a href="?dick=15&city_one=<%=city_oneid%>&city_two=<%=city_twoid%>&city_three=<%=city_threeid%>&type_one=<%=type_oneid%>&type_two=<%=type_twoid%>&type_three=<%=type_threeid%>"  title='按点击数排序'>热门信息</a>---<a href="?dick=16&city_one=<%=city_oneid%>&city_two=<%=city_twoid%>&city_three=<%=city_threeid%>&type_one=<%=type_oneid%>&type_two=<%=type_twoid%>&type_three=<%=type_threeid%>" title='指点击数为0的信息'>冷门信息</a>---<a href="?dick=8&city_one=<%=city_oneid%>&city_two=<%=city_twoid%>&city_three=<%=city_threeid%>&type_one=<%=type_oneid%>&type_two=<%=type_twoid%>&type_three=<%=type_threeid%>" title='指发布3个月以上且点击数为0的信息'>3月以上冷门信息</a>---<a href="?dick=17&city_one=<%=city_oneid%>&city_two=<%=city_twoid%>&city_three=<%=city_threeid%>&type_one=<%=type_oneid%>&type_two=<%=type_twoid%>&type_three=<%=type_threeid%>" title='指信息有效时间已过且点击数为0的信息'>过期冷门信息</a><br><a href="?dick=10&city_one=<%=city_oneid%>&city_two=<%=city_twoid%>&city_three=<%=city_threeid%>&type_one=<%=type_oneid%>&type_two=<%=type_twoid%>&type_three=<%=type_threeid%>" title='指已审核的竞价信息'>竞价信息</a>---<a href="?dick=9&city_one=<%=city_oneid%>&city_two=<%=city_twoid%>&city_three=<%=city_threeid%>&type_one=<%=type_oneid%>&type_two=<%=type_twoid%>&type_three=<%=type_threeid%>" title='指待审核的竞价信息'>申请竞价信息</a>---<a href="?dick=11&city_one=<%=city_oneid%>&city_two=<%=city_twoid%>&city_three=<%=city_threeid%>&type_one=<%=type_oneid%>&type_two=<%=type_twoid%>&type_three=<%=type_threeid%>" title='指已到期的竞价信息'>到期竞价信息</a>---<a href="?dick=12&city_one=<%=city_oneid%>&city_two=<%=city_twoid%>&city_three=<%=city_threeid%>&type_one=<%=type_oneid%>&type_two=<%=type_twoid%>&type_three=<%=type_threeid%>" title='指已经生成html文件的信息'>已生成html信息</a>---<a href="?dick=13&city_one=<%=city_oneid%>&city_two=<%=city_twoid%>&city_three=<%=city_threeid%>&type_one=<%=type_oneid%>&type_two=<%=type_twoid%>&type_three=<%=type_threeid%>" title='指尚未生成html文件的信息'>未生成html信息</a>---<a href="?dick=19&city_one=<%=city_oneid%>&city_two=<%=city_twoid%>&city_three=<%=city_threeid%>&type_one=<%=type_oneid%>&type_two=<%=type_twoid%>&type_three=<%=type_threeid%>" title='指尚未生成html文件的信息'>普通会员阅读</a>---<a href="?dick=20&city_one=<%=city_oneid%>&city_two=<%=city_twoid%>&city_three=<%=city_threeid%>&type_one=<%=type_oneid%>&type_two=<%=type_twoid%>&type_three=<%=type_threeid%>" title='指尚未生成html文件的信息'>VIP会员阅读</a>---<a href="?city_one=<%=city_oneid%>&city_two=<%=city_twoid%>&city_three=<%=city_threeid%>&type_one=<%=type_oneid%>&type_two=<%=type_twoid%>&type_three=<%=type_threeid%>" title='指按地区查看的全部信息'>本地区全部信息</a>---<a href="?dick=" title='指全部的信息'>全部信息</a><br>&nbsp;&nbsp;&nbsp;&nbsp;<a href="info_html.asp" title='跳转到html生成管理页面'><b>生成html管理</b></a>&nbsp;&nbsp;&nbsp;&nbsp;<a href="info_del_dqxx.asp" onClick="return confirm('确定要清空过期信息吗？\n\n将同时删除相关图片等文件数据\n\n该操作不可恢复！')" ><b>一键删除过期信息</b></a>&nbsp;&nbsp;&nbsp;&nbsp;<a href="#" ONCLICK="window.open('info_replyxx.asp','Sample','toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=yes,resizable=no,copyhistory=no,width=570,height=200,left=300,top=20')" title='进入下一步操作'><b>取消到期竞价信息资格</b></a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
<%set rs=server.createobject("adodb.recordset")
sql = "select * from [DNJY_info] where yz<>1"
rs.open sql,conn,1,1
if rs.eof then
response.write "<font color=#0000FF>目前没有待审核信息！</font>"
Else
response.write "<b><img src='../img/notice.gif' width=15 height=15 border=0><font color=#FF0000>铃声响有待审核信息！</font></b>"
response.write "<bgsound src='../img/msg.wav' loop='1'>"
end If
Call CloseRs(rs)%>
 </td>
      </tr>
    </table>
    </td>
  </tr>
  <FORM name=thisForm method=POST action="info_del.asp?urlpage=infolist.asp">
<%
function HTMLEncode(fString)
if not isnull(fString) then
    fString = replace(fString, ">", "&gt;")
    fString = replace(fString, "<", "&lt;")
    fString = Replace(fString, CHR(32), "&nbsp;")
    fString = Replace(fString, CHR(9), "&nbsp;")
    fString = Replace(fString, CHR(34), "&quot;")
    fString = Replace(fString, CHR(39), "&#39;")
    fString = Replace(fString, CHR(13), "")
    fString = Replace(fString, CHR(10) & CHR(10), "</P><P> ")
    fString = Replace(fString, CHR(10), "<BR> ")
    HTMLEncode = fString
end if
end Function


set rs = Server.CreateObject("ADODB.RecordSet")
Select Case dick
Case "1"
sql="select * from [DNJY_info] where username='"&trim(request("key1"))&"'"&ttcc&""&tttt&" order by fbsj "&DN_OrderType&"" 
Case "2"
sql="select * from [DNJY_info] where name like '%"&trim(request("key2"))&"%'"&ttcc&""&tttt&" order by fbsj "&DN_OrderType&"" 
Case "3"
If request("skey")=0 then
sql="select * from [DNJY_info] where (biaoti like '%"&trim(request("key3"))&"%')"&ttcc&""&tttt&" order by fbsj "&DN_OrderType&"" 
Else
sql="select * from [DNJY_info] where (memo like '%"&trim(request("key3"))&"%')"&ttcc&""&tttt&" order by fbsj "&DN_OrderType&""
End if
Case "4"
sql="select * from [DNJY_info] where yz<>1"&ttcc&""&tttt&" order by fbsj "&DN_OrderType&"" 
Case "5"
sql="select * from [DNJY_info] where yz=1"&ttcc&""&tttt&" order by fbsj "&DN_OrderType&""
Case "6"
sql="select * from [DNJY_info] where yz=1 and tj>0"&ttcc&""&tttt&" order by fbsj "&DN_OrderType&""
Case "7"
sql="select * from [DNJY_info] where "&DN_Now&">=dqsj"&ttcc&""&tttt&" order by fbsj "&DN_OrderType&""
Case "8"
sql="select * from [DNJY_info] where DateDiff("&DN_DatePart_D&",fbsj,"&DN_Now&")>=90"&ttcc&""&tttt&" and llcs=0 order by fbsj "&DN_OrderType&""
Case "9"
sql="select * from [DNJY_info] where hfxx=2"&ttcc&""&tttt&" order by fbsj "&DN_OrderType&""
Case "10"
sql="select * from [DNJY_info] where (hfxx=1 or hfxx=2)"&ttcc&""&tttt&" order by fbsj "&DN_OrderType&""
Case "11"
sql="select * from [DNJY_info] where hfxx=1"&ttcc&""&tttt&" and "&DN_Now&">=dqsj order by fbsj "&DN_OrderType&""
Case "12"
sql="select * from [DNJY_info] where fsohtm=1"&ttcc&""&tttt&" order by fbsj "&DN_OrderType&""
Case "13"
sql="select * from [DNJY_info] where fsohtm=0"&ttcc&""&tttt&" order by fbsj "&DN_OrderType&""
Case "14"
sql="select * from [DNJY_info] where b>0"&ttcc&""&tttt&" order by b "&DN_OrderType&""
Case "15"
sql="select * from [DNJY_info] where llcs>0"&ttcc&""&tttt&" order by llcs "&DN_OrderType&""
Case "16"
sql="select * from [DNJY_info] where llcs=0"&ttcc&""&tttt&" order by fbsj"
Case "17"
sql="select * from [DNJY_info] where "&DN_Now&">=dqsj"&ttcc&""&tttt&" and llcs=0 order by fbsj "&DN_OrderType&""
Case "18"
sql="select * from [DNJY_info] where username='"&trim(request("key1"))&"'"&ttcc&""&tttt&" order by fbsj "&DN_OrderType&"" 
Case "19"
sql="select * from [DNJY_info] where Readinfo=1"&ttcc&""&tttt&" order by fbsj "&DN_OrderType&""
Case "20"
sql="select * from [DNJY_info] where Readinfo=2"&ttcc&""&tttt&" order by fbsj "&DN_OrderType&""
Case Else
sql="select * from [DNJY_info] where fbts>0"&ttcc&""&tttt&" order by fbsj "&DN_OrderType&"" 
End Select



rs.open sql,conn,1,1
if rs.eof then
response.write "<table width=""96%""><tr><td width=""100%""  colspan=""12"" align=""center""><br><li>没有记录</table></td></tr>"
response.end
end if
if request("page")="" then
  ThisPage=1		
else
  ThisPage=request("page")
end if
rs.Pagesize=20'每页条数
Pagesize=rs.Pagesize
Allrecord=rs.Recordcount
Allpage=rs.Pagecount
if ThisPage<1 then                           
ThisPage=1
end if
On Error Resume Next
rs.move (ThisPage-1)*Pagesize
%>
  <tr bgcolor="#BDBEDE">
    <td width="5%" height="24" align="center" style="border-bottom-style: solid; border-bottom-width: 1" bordercolor="#666666">选择</td>
    <td width="20%" height="28" bordercolor="#666666" bgcolor="#BDBEDE" style="border-bottom-style: solid; border-bottom-width: 1">标题</td>
    <td width="10%" height="24" style="border-bottom-style: solid; border-bottom-width: 1" bordercolor="#666666">
    <p align="center">注册ID/姓名</td>
    <td width="16%" height="24" style="border-bottom-style: solid; border-bottom-width: 1" bordercolor="#666666">
    <p align="center">添加时间/发布时间</td>
    <td width="5%" height="24" align="center" style="border-bottom-style: solid; border-bottom-width: 1" bordercolor="#666666">阅读权限</td>
	<td width="5%" height="24" align="center" style="border-bottom-style: solid; border-bottom-width: 1" bordercolor="#666666">点击</td>
	<td width="5%" height="24" align="center" style="border-bottom-style: solid; border-bottom-width: 1" bordercolor="#666666">回复</td>
	<td width="5%" height="24" align="center" style="border-bottom-style: solid; border-bottom-width: 1" bordercolor="#666666">举报</td>
    <td width="4%" height="24" style="border-bottom-style: solid; border-bottom-width: 1" bordercolor="#666666">
    <p align="center">验证</td>
    <td width="4%" height="24" style="border-bottom-style: solid; border-bottom-width: 1" bordercolor="#666666">
    <p align="center">置顶</td>
    <td width="4%" height="24" style="border-bottom-style: solid; border-bottom-width: 1" bordercolor="#666666">
    <p align="center">推荐</td>
    <td width="5%" height="24" style="border-bottom-style: solid; border-bottom-width: 1" bordercolor="#666666">
    <p align="center">竞价|操作</td>
	<td width="4%" height="24" style="border-bottom-style: solid; border-bottom-width: 1" bordercolor="#666666">
    <p align="center">已生成</td>  
    <td width="4%" height="24" bordercolor="#666666" bgcolor="#BDBEDE" style="border-bottom-style: solid; border-bottom-width: 1"><div align="center">操作</div></td>
  </tr>
<%dim uid
k=0
do while not rs.eof
uid=rs("id")

	    dim rs1,sql1
		set rs1=server.createobject("adodb.recordset") 
		sql1="select * from DNJY_reply where xxid="&rs("id")&" order by id "&DN_OrderType&""
		rs1.open sql1,conn,1,1
		
		dim rs2,sql2
		set rs2=server.createobject("adodb.recordset") 
		sql2="select * from DNJY_Bad_info where infoid="&rs("id")&" order by id "&DN_OrderType&""
		rs2.open sql2,conn,1,1%>
  <tr>
    <td width="5%" height="24" align="center" bgcolor="#FFFFFF" style="border-top-style: solid; border-top-width: 1">
    <INPUT type=checkbox value="<%=uid%>" name=selectedid></td>

    <td width="20%" height="24" bgcolor="#FAFAFA" style="border-top-style: solid; border-top-width: 1">
   <% if rs("tupian")<>"0" then
response.write "<img src=""../images/num/pic.gif"">" 
end if%>
<%	belongtype=""&rs("type_one")&""
	IF rs("type_two")<>"" and not isnull(rs("type_two")) Then belongtype=belongtype&" / "&rs("type_two")&""
	IF rs("type_three")<>"" and not isnull(rs("type_three")) Then belongtype=belongtype&" / "&rs("type_three")&""
	%>
<A href="admin_info.asp?id=<%=uid%>" target="_blank" alt="<%=rs("biaoti")%><br>类别：<%=belongtype%>|有效期：<%=DateDiff("d",Now(),""&rs("dqsj")&"")%>天"><%=CutStr(rs("biaoti"),36,"...")%></a></td>
    <td width="10%" height="24" align="center" bgcolor="#FFFFFF" style="border-top-style: solid; border-top-width: 1"><a href="?dick=18&key1=<%= rs("username")%>&city_one=<%=city_oneid%>&city_two=<%=city_twoid%>&city_three=<%=city_threeid%>&type_one=<%=type_oneid%>&type_two=<%=type_twoid%>&type_three=<%=type_threeid%>" title='查<%=rs("username")%>发布的供求信息'><%=rs("username")%><font color="#FF0000">|</font><%=rs("name")%></a></td>
    <td width="15%" height="24" align="center" bgcolor="#FAFAFA" style="border-top-style: solid; border-top-width: 1"><%=rs("addsj")%><br><font color="#FF0000">/</font><%If Now()>=rs("dqsj") then%><font color="#FF0000"><%=rs("fbsj")%></font><%else%><%=rs("fbsj")%><%End if%></td>
<td width="4%" height="24" align="center" bgcolor="#FFFFFF" style="border-top-style: solid; border-top-width: 1"><%if rs("Readinfo")=0 then%>游客<%else%><%if rs("Readinfo")=1 then%><font color="#0000FF">普通会员</font><%else%><font color="#FF0000">VIP会员</font><%end if%><%end if%></td>
	<td width="4%" height="24" align="center" bgcolor="#FFFFFF" style="border-top-style: solid; border-top-width: 1"><%=rs("llcs")%></td>
	<td width="5%" align="center" bgcolor="#FFFFFF"><img onClick="javascript:left_menu('left_<%=rs("id")%>');" src="img/dedeexplode.gif" alt="展开回复" width="11" height="11"> <%if rs1.recordcount>0 then%><font color="#FF0000"><%=rs1.recordcount%></font><%else%><%=rs1.recordcount%><%End if%></td>
	<td width="5%" align="center" bgcolor="#FFFFFF"><img onClick="javascript:left_menu('left2_<%=rs("id")%>');" src="img/dedeexplode.gif" alt="展开举报" width="11" height="11"> <%if rs2.recordcount>0 then%><font color="#FF0000"><%=rs2.recordcount%></font><%else%><%=rs2.recordcount%><%End if%></td>
    <td width="4%" height="24" align="center" bgcolor="#FFFFFF" style="border-top-style: solid; border-top-width: 1"><%if rs("yz")=1 then%><a href="info_yz.asp?selectedid=<%=uid%>&qxyz=no&page=<%=(ThisPage)%>&t1=<%=trim(request("t1"))%>&t2=<%=trim(request("t2"))%>&t3=<%=trim(request("t3"))%>&dick=<%=dick%>&city_one=<%=city_oneid%>&city_two=<%=city_twoid%>&city_three=<%=city_threeid%>&type_one=<%=type_oneid%>&type_two=<%=type_twoid%>&type_three=<%=type_threeid%>">√</a><%else%><b><a href="info_yz.asp?selectedid=<%=uid%>&qxyz=yes&page=<%=(ThisPage)%>&t1=<%=trim(request("t1"))%>&t2=<%=trim(request("t2"))%>&t3=<%=trim(request("t3"))%>&dick=<%=dick%>&city_one=<%=city_oneid%>&city_two=<%=city_twoid%>&city_three=<%=city_threeid%>&type_one=<%=type_oneid%>&type_two=<%=type_twoid%>&type_three=<%=type_threeid%>"><font color="#FF0000">ⅹ</font></a></b><%end if%></td>
	<td width="4%" height="24" align="center" bgcolor="#FFFFFF" style="border-top-style: solid; border-top-width: 1"><%=rs("b")%></td>
    <td width="4%" height="24" align="center" bgcolor="#FFFFFF" style="border-top-style: solid; border-top-width: 1"><%if rs("tj")=0 then%>未推<%else%><%if rs("tj")=1 then%><font color="#FF0000">信息</font><%else%><font color="#FF0000">图片</font><%end if%><%end if%></td>
    <td width="5%" height="24" align="center" bgcolor="#FFFFFF" style="border-top-style: solid; border-top-width: 1"><%if rs("hfxx")=1 then%><font color="#0000FF"><span alt="竞价均价<%=FormatNumber(rs("hfje")/rs("fbts"),2,-1)%><br>竞价剩<b><%=Fix(rs("dqsj")-now())%></b>天" style="cursor:help">是</span>|<a href="info_gl.asp?selectedid=<%=uid%>&gl=jj&qxyz=no&page=<%=(ThisPage)%>&t1=<%=trim(request("t1"))%>&t2=<%=trim(request("t2"))%>&t3=<%=trim(request("t3"))%>&dick=<%=dick%>&city_one=<%=city_oneid%>&city_two=<%=city_twoid%>&city_three=<%=city_threeid%>&type_one=<%=type_oneid%>&type_two=<%=type_twoid%>&type_three=<%=type_threeid%>">√</a></font><%else%><%if rs("hfxx")=2 then%><font color="#FF0000">申请<%=FormatNumber(rs("hfje")/rs("fbts"),2,-1)%></font><%else%>否|<a href="info_gl.asp?selectedid=<%=uid%>&gl=jj&qxyz=yes&page=<%=(ThisPage)%>&t1=<%=trim(request("t1"))%>&t2=<%=trim(request("t2"))%>&t3=<%=trim(request("t3"))%>&dick=<%=dick%>&city_one=<%=city_oneid%>&city_two=<%=city_twoid%>&city_three=<%=city_threeid%>&type_one=<%=type_oneid%>&type_two=<%=type_twoid%>&type_three=<%=type_threeid%>"><font color="#FF0000">ⅹ</font></a><%end if%><%end if%></td> 
    <td width="4%" height="24" align="center" bgcolor="#FFFFFF" style="border-top-style: solid; border-top-width: 1"><%if rs("fsohtm")=1 then%><span title=文件位置：/<%=strInstallDir%>html/<%=rs("id")%>.html style="cursor:help">√</span><%else%><b><font color="#FF0000">ⅹ</font></a></b><%end if%></td>    
    <td width="4%" height="24" bgcolor="#FAFAFA" style="border-top-style: solid; border-top-width: 1">
    <p align="center"><font color="#FF0000"><span id="followImg<%=k%>" style="CURSOR: hand" onClick="loadThreadFollow(<%=k%>,5)">面板</span></font></td>
  </tr>
  <tr style="display:none" id="follow<%=k%>">
    <td width="100%" height="24" colspan="14" style="border-bottom-style: solid; border-bottom-width: 1" bgcolor="#FFF8EC">
    <div align="center">
      <center>
    <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%" height="100%">
      <tr>
        <td bgcolor="#CCCCFF"><div align="right">IP:<%= rs("ip") %> <a href='info_edit.asp?id=<%=uid%>'>编辑该信息</a> <a href="#" ONCLICK="window.open('user_mail.asp?username=<%=rs("username")%>&email=<%=rs("email")%>','Sample','toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=yes,resizable=no,copyhistory=no,width=700,height=700,left=300,top=100')">发送邮件</a>  <a href='message2.asp?username=<%=trim(rs("username"))%>'>发送短信</a></font>         
          <%if rs("yz")=0 then%>
              <a href="info_yz.asp?selectedid=<%=uid%>">通过验证</a>
              <%else%>
 <a href="#" ONCLICK="window.open('info_tj.asp?id=<%=uid%>','Sample','toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=no,resizable=no,copyhistory=no,width=250,height=150,left=300,top=100')">信息推荐</a>                
              <a href="info_yz.asp?selectedid=<%=uid%>&qxyz=no">取消验证</a>              
              <%end if%>
			  <a href="javascript:DoEmpty('info_del.asp?selectedid=<%=uid%>&urlpage=infolist.asp');">删除该信息</a> <a href="javascript:DoEmpty('info_replydel.asp?selectedid=<%=uid%>');">删除所有回复</a>
            &nbsp;</div></td>
        </tr>
    </table>
      </center>
    </div>
    </td>
  </tr>
  
  <tr bgcolor="#ffffff" id="left_<%=rs("id")%>" style="display:none;">
    <td height="22" colspan="14" align="right">
	<table width="95%" border="0" cellpadding="2" cellspacing="1" bgcolor="#999999">
      <tr>
        <td width="4%" height="25" align="center" bgcolor="#cccccc">id</td>
        <td width="8%" align="center" bgcolor="#cccccc">用户名称</td>
        <td width="45%" align="center" bgcolor="#cccccc">回复内容</td>
        <td width="13%" align="center" bgcolor="#cccccc">回复时间</td>
		 <td width="15%" align="center" bgcolor="#cccccc">IP地址</td>
        <td width="5%" align="center" bgcolor="#cccccc">编辑</td>
        <td width="5%" align="center" bgcolor="#cccccc">审核</td>
        <td width="5%" align="center" bgcolor="#cccccc">操作</td>
      </tr>
	  <%if rs1.eof and rs1.bof then%>
		 <tr>
        <td height="25" colspan="14" align="center" bgcolor="#E3EBF9">暂时还没有回复</td>
        </tr>
		<%else
		do while not rs1.eof%>
      <tr>
        <td height="25" align="center" bgcolor="#E3EBF9"><%=rs1("id")%></td>
        <td align="center" bgcolor="#E3EBF9"><%if rs1("username")<>"" then%><a target="_blank" href=../preview.asp?username=<%=rs1("username")%>><%=rs1("username")%><%else%>匿名<%end if%></td>
        <td bgcolor="#E3EBF9"><%=rs1("neirong")%></td>
        <td align="center" bgcolor="#E3EBF9"><%=rs1("hfsj")%></td>
		<td align="center" bgcolor="#E3EBF9"><%=rs1("ip")%></td>
		<td align="center" bgcolor="#E3EBF9"><a href="chkcomment.asp?edit=replay&id=<%=rs1("id")%>">编辑</a></td>		
       <td align="center" bgcolor="#E3EBF9">
	<%if rs1("hfyz")=1 then%><a href="info_replyyz.asp?dick=delyz&selectedid=<%=rs1("id")%>">√</a><%else%><b><a href="info_replyyz.asp?selectedid=<%=rs1("id")%>"><font color="#FF0000">ⅹ</font></a></b><%end if%></td>
        <td align="center" bgcolor="#E3EBF9"><%if session("aleave")="check" then%>删除<%else%><a href="javascript:DoEmpty('info_replydel.asp?selectedid=<%=rs1("id")%>&xid=<%=uid%>&back=xinxi');">删除</a><%end if%></td>
      </tr>
     
	  <%rs1.movenext
		loop
		rs1.close
		set rs1=nothing
		end if%>
    </table>	</td>
  </tr>

  <tr bgcolor="#ffffff" id="left2_<%=rs("id")%>" style="display:none;">
    <td height="22" colspan="14" align="right">
	<table width="95%" border="0" cellpadding="2" cellspacing="1" bgcolor="#999999">
      <tr>
        <td width="5%" height="25" align="center" bgcolor="#cccccc">id</td>
        <td width="10%" align="center" bgcolor="#cccccc">举报者信箱</td>
        <td width="12%" align="center" bgcolor="#cccccc">举报类型</td>
        <td width="40%" align="center" bgcolor="#cccccc">举报内容</td>
		<td width="13%" align="center" bgcolor="#cccccc">举报时间</td>
		<td width="10%" align="center" bgcolor="#cccccc">IP地址</td>        
        <td width="5%" align="center" bgcolor="#cccccc">审核</td>
        <td width="5%" align="center" bgcolor="#cccccc">操作</td>
      </tr>
	  <%

		if rs2.eof and rs2.bof then
		
		%>
		 <tr>
        <td height="25" colspan="14" align="center" bgcolor="#E3EBF9">暂时还没有举报</td>
        </tr>
		<%else
		do while not rs2.eof 
	  %>
      <tr>
        <td height="25" align="center" bgcolor="#E3EBF9"><%=rs2("id")%></td>
        <td align="center" bgcolor="#E3EBF9"><a href="#" onClick="window.open('Bad_info_mail.asp?email=<%=rs2("Bad_infomail")%>&biaoti=<%=rs("biaoti")%>','Sample','toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=no,resizable=no,copyhistory=no,width=455,height=328,left=300,top=100')"><%=rs2("Bad_infomail")%></a></td>
        <td bgcolor="#E3EBF9"><%=rs2("Bad_infolx")%></td>
		<td bgcolor="#E3EBF9"><%=HTMLEncode(rs2("Bad_info"))%></td>
        <td align="center" bgcolor="#E3EBF9"><%=rs2("addsj")%></td>
		<td align="center" bgcolor="#E3EBF9"><%=rs2("ip")%></td>			
       <td align="center" bgcolor="#E3EBF9">
	<%if rs2("sh")=1 then%><a href="Bad_info_yz.asp?dick=delyz&selectedid=<%=rs2("id")%>">√</a><%else%><b><a href="Bad_info_yz.asp?selectedid=<%=rs2("id")%>"><font color="#FF0000">ⅹ</font></a></b><%end if%></td>
        <td align="center" bgcolor="#E3EBF9"><a href="javascript:DoEmpty('Bad_info_del.asp?selectedid=<%=rs2("id")%>');">删除</a></td>
      </tr>
     
	  <%rs2.movenext
		loop
		rs2.close
		set rs2=nothing
		end if%>
    </table>	</td>
  </tr>
<%
k=k+1
rs.movenext
if k>=Pagesize then exit do
loop
Call CloseRs(rs)
Call CloseDB(conn)
%>
  <tr>
    <td width="100%" height="1" colspan="14" style="border-top-style: solid; border-top-width: 1">
<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%" height="1">
<tr> 
<td width="100%" height="25" align="center" bgcolor="#eeeeee"><div align="center">
      <input onClick=CheckAll(this.form) type=checkbox value=on name=chkall>
      选中所有
<input onclick=javascript:showoperatealert(1) type="submit" value="验证通过" name="B1">
<input onclick=javascript:showoperatealert(2) type="submit" value="取消验证" name="B2">	  
<input onclick=javascript:showoperatealert(3) type="submit" value="信息推荐" name="B3">
<input onclick=javascript:showoperatealert(4) type="submit" value="图片推荐" name="B4">
<input onclick=javascript:showoperatealert(5) type="submit" value="取消推荐" name="B5">
<input name='submit1' type='submit' value=' 批量删除 ' onClick="return test();"  style="color: #FF0000">

<br>共有&nbsp;<font color="#CC5200"><%=Allrecord%></font>&nbsp;条记录 共 <font color="#CC5200"><%=Allpage%></font> 页 每页 <font color="#CC5200"><%=Pagesize%></font> 条 现在是第 <font color="#CC5200"><%=ThisPage%></font> 页
    <%
if ThisPage<2 then     
response.write "<font color=""#808080"">首页</font>&nbsp;"
response.write "<font color=""#808080"">上一页</font>&nbsp;"     
else     
response.write "<a href=?page=1&key1="&trim(request("key1"))&"&key2="&trim(request("key2"))&"&key3="&trim(request("key3"))&"&skey="&trim(request("skey"))&"&t1="&trim(request("t1"))&"&t2="&trim(request("t2"))&"&t3="&trim(request("t3"))&"&dick="&dick&"&city_one="&city_oneid&"&city_two="&city_twoid&"&city_three="&city_threeid&"&type_one="&type_oneid&"&type_two="&type_twoid&"&type_three="&type_threeid&">首页</a>&nbsp;"
response.write "<a href=?page="&ThisPage-1&"&key1="&trim(request("key1"))&"&key2="&trim(request("key2"))&"&key3="&trim(request("key3"))&"&skey="&trim(request("skey"))&"&t1="&trim(request("t1"))&"&t2="&trim(request("t2"))&"&t3="&trim(request("t3"))&"&dick="&dick&"&city_one="&city_oneid&"&city_two="&city_twoid&"&city_three="&city_threeid&"&type_one="&type_oneid&"&type_two="&type_twoid&"&type_three="&type_threeid&">上一页</a>&nbsp;"     
end if
if Allpage-ThisPage<1 then     
response.write "<font color=""#808080"">下一页</font>&nbsp;"
response.write "<font color=""#808080"">尾页</font>&nbsp;"  
else     
response.write "<a href=?page="&(ThisPage+1)&"&key1="&trim(request("key1"))&"&key2="&trim(request("key2"))&"&key3="&trim(request("key3"))&"&skey="&trim(request("skey"))&"&t1="&trim(request("t1"))&"&t2="&trim(request("t2"))&"&t3="&trim(request("t3"))&"&dick="&dick&"&city_one="&city_oneid&"&city_two="&city_twoid&"&city_three="&city_threeid&"&type_one="&type_oneid&"&type_two="&type_twoid&"&type_three="&type_threeid&">下一页</a>&nbsp;"   
response.write "<a href=?page="&Allpage&"&key1="&trim(request("key1"))&"&key2="&trim(request("key2"))&"&key3="&trim(request("key3"))&"&skey="&trim(request("skey"))&"&t1="&trim(request("t1"))&"&t2="&trim(request("t2"))&"&t3="&trim(request("t3"))&"&dick="&dick&"&city_one="&city_oneid&"&city_two="&city_twoid&"&city_three="&city_threeid&"&type_one="&type_oneid&"&type_two="&type_twoid&"&type_three="&type_threeid&">尾页</a>&nbsp;"     
end if
%>
</div></td>
</tr>
</table>
    </td>
  </tr>
  </form>
  <tr>
    <td width="100%" height="1" colspan="10" style="border-top-style: solid; border-top-width: 1"></td>
  </tr>
</table></center>
</div>
