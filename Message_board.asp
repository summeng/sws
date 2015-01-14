<!--#include file="conn.asp"-->
<!--#include file="SqlIn.Asp"-->
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
<%if not ChkPost() then 
response.write "<center>禁止站外提交或访问"
response.end 
end if
if lmkg2="1" then
call errdick(50)
Call CloseDB(conn)
response.end
end if
dim ts,dick,strUserName,mybook
dim ThisPage,Pagesize,Allrecord,Allpage,tj
mybook=request.querystring("mybook")
username=request.cookies("dick")("username")
strUserName = "非注册用户"
If request.cookies("dick")("username")<>"" Then
strUserName=request.cookies("dick")("username")
End if
ts=trim(request("ts"))
if request("page")="" then
  ThisPage=1		
else
  ThisPage=request("page")
end if
%>
<link rel="stylesheet" type="text/css" href="/<%=strInstallDir%>css/css.css">
<script LANGUAGE=JavaScript>
  function textLimitCheck(thisArea, maxLength){
    if (thisArea.value.length > maxLength){
      alert(maxLength + ' 个字限制. \r超出的将自动去除.');
      thisArea.value = thisArea.value.substring(0, maxLength);
      thisArea.focus();
    }
    /*回写span的值，当前填写文字的数量*/
    messageCount.innerText = thisArea.value.length;
  }
</script>
<script language="javascript">
<!--
function formcheck(){

	   if (document.form.memo.value=="")
	{
      alert("留言内容不能为空，请重新输入！")
	  document.form.memo.focus()
	  return false
	 }

 <%if codekg1=1 then%>
    if (document.form.wenti.value=="" )
	{	  
      alert("验证答案不能为空，请重新输入！");
	  document.form.wenti.focus()
	  return false
	 }
	
    if(document.form.wenti.value != document.form.daan.value) 
	{	  
      alert("验证答案不对，请重新输入！");
	  document.form.wenti.focus()
	  return false
	 }
 <%end if%>
 <%if codekg2=1 then%>
    if (document.form.yzm.value=="" )
	{	  
      alert("数字验证码不能为空，请重新输入！");
	  document.form.yzm.focus()
	  return false
	 }
<%end if%>
 <%if codekg3=1 then%>
    if (document.form.code.value=="" )
	{	  
      alert("文字验证码不能为空，请重新输入！");
	  document.form.code.focus()
	  return false
	 }
<%end if%>
<%if codekg4=1 then%>
    if (document.form.ttdv.value=="" )
	{	  
      alert("验证星期不能为空，请重新输入！");
	  document.form.ttdv.focus()
	  return false
	 }
<%end if%>
<%if codekg5=1 then%>
    if (document.form.validate_codee.value=="" )
	{	  
      alert("算式验证码不能为空，请重新输入！");
	  document.form.validate_codee.focus()
	  return false
	 }
<%end if%>	  
}
// --></script>

<!--文字验证码调用开始-->
<SCRIPT LANGUAGE=javascript>
/*显示认证码 o start1*/
function get_Code() {
        var Dv_CodeFile = "Dv_GetCode.asp";
        if(document.getElementById("imgid"))
                document.getElementById("imgid").innerHTML = '<img src="'+Dv_CodeFile+'?t='+Math.random()+'" alt="点击刷新验证码" style="cursor:pointer;border:0;vertical-align:middle;height:30px;" onclick="this.src=\''+Dv_CodeFile+'?t=\'+Math.random()" />'
}
/*o end*/
</script>
<script language="JavaScript" type="text/javascript">
var dvajax_request_type = "GET";
</script>
<script language="JavaScript" src="dv_ajax.js" type="text/javascript"></script>
<!---文字验证码调用结束-->
</head>

<body>
<table width="980" border="0" align="center" cellpadding="0" cellspacing="0">

  <tr>
    <td width="215" height="100%" valign="top"><!--#include file=left.asp--></td>
    <td width="5">&nbsp;</td>
    <td width="760" valign="top" bgcolor="#FFFFFF">
	<table width="760" border="0" align="right" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF" class="ty1">

      <tr>
        <td  bgcolor="#FFFFFF"><table width="760" height="166" border="0" cellpadding="0" cellspacing="0" class="dq2" >
<%set rs=server.createobject("adodb.recordset")
if  strUserName="非注册用户" then
sql="select * from DNJY_gbook where hf=1 and known=0"&ttccbook&" order by fbsj "&DN_OrderType&""
Else
if mybook="ok" Then
sql="select * from DNJY_gbook where username='"&strUserName&""&ttccbook&"' order by fbsj "&DN_OrderType&""
Else
sql="select * from DNJY_gbook where (hf=1 and known=0) or username='"&strUserName&"'"&ttccbook&" order by fbsj "&DN_OrderType&""
end If
end if
rs.open sql,conn,1,1
if rs.eof Then
Response.Write "<p><br>暂无"
IF c1>0 Then Response.Write conn.Execute("Select city From DNJY_city Where id="&c1&" and twoid="&c2&" and threeid="&c3&"")(0)
Response.Write "用户留言!"
else
rs.Pagesize=7
Pagesize=rs.Pagesize
Allrecord=rs.Recordcount
Allpage=rs.Pagecount
if ThisPage<1 then                           
ThisPage=1                           
end if
rs.move (ThisPage-1)*Pagesize
%>
<%Dim num1,bg
do while not rs.eof
num1 = Num1+1 '//判断数字奇偶定背景色
If num1 Mod 2=0 Then
bg="F1F9FE"
Else
bg="F9F6FF"
End if			%>

            <tr bgcolor="#<%=bg%>">
              <td width="110" height="22" background="img/22.gif"><p style="margin-top: 0; margin-bottom: 0"> <font color="#FF0000"> <strong> &nbsp;</strong><%If rs("username")<>"" then%><%=rs("username")%><%else%>游客<%End if%>问：</font></td>
              <td width="650" height="26" background="img/22.gif"><font color="#800000"><b><%=rs("gbook1")%></b></font></td>
            </tr>
            <tr bgcolor="#<%=bg%>">
              <td width="110" valign="top" height="19"><p style="margin-top: 0; margin-bottom: 0"><strong><font color="#009900"> &nbsp;</font></strong><font color="#009900">管理员答：</font></td>
              <td height="27"><font color="#808080"><%If rs("hf")=1 then%><%=rs("gbook2")%>(回复时间：<font color="#666666"><%=rs("hfsj")%></font>)<%else%><font color="#FF0000">尚未回复</font><%End if%></font></td>
            </tr>

            <%
        tj=tj+1
        if tj>=Pagesize then exit do
        rs.movenext
        loop
        %>

            <tr>
              <td colspan="2" height="6"><table width="100%" border="0" cellpadding="0" cellspacing="0" class="dq2" style="border-collapse: collapse">
                  <tr>
                    <td height="25" width="151"><p align="center"> 共有&nbsp;<font color="#CC5200"><%=Allrecord%></font>&nbsp;条记录</td>
                    <td height="25" width="126"><p align="center">共 <font color="#CC5200"><%=Allpage%></font> 页</td>
                    <td height="25" width="118"><p align="center">现在是第 <font color="#CC5200"><%=ThisPage%></font> 页</td>
                    <td height="25" width="187"><p align="center">
                        <%
if ThisPage<2 then     
response.write "<font color=""#808080"">首页</font>&nbsp;"
response.write "<font color=""#808080"">上一页</font>&nbsp;"     
else     
response.write "<a href=?page=1&ts="&ts&">首页</a>&nbsp;"
response.write "<a href=?page="&ThisPage-1&"&ts="&ts&">上一页</a>&nbsp;"     
end if
if Allpage-ThisPage<1 then     
response.write "<font color=""#808080"">下一页</font>&nbsp;"
response.write "<font color=""#808080"">尾页</font>&nbsp;"  
else     
response.write "<a href=?page="&(ThisPage+1)&"&ts="&ts&">下一页</a>&nbsp;"   
response.write "<a href=?page="&Allpage&"&ts="&ts&">尾页</a>&nbsp;"     
end if
%>
                      </td>
                  </tr>
              </table></td>
            </tr><%end if%>
            <tr>
              <td height="27" colspan="2">&nbsp;</td>
            </tr>
            <tr bgcolor="#FAFAFA">
              <td height="27" colspan="2" align="center"><span class="style1">本留言仅作站长与用户之间的站务咨询服务</span></td>
            </tr>
            <tr bgcolor="#CCCCCC">
              <td height="1" colspan="2"></td>
            </tr>
          </table>         

            <table width="760" border="0" align="center" cellPadding="0" cellSpacing="0" bordercolor="#111111" class="font_10_e_blue" style="border-collapse: collapse">
              <tr>
                <td vAlign="top" width="760"><table class="font_10_e_black" cellSpacing="0" cellPadding="3" width="100%" align="center" border="0">
  
<form id="applyform" method="post" name="form" action="Message_board_save.asp?<%=C_ID%>" language="javascript" onsubmit="return formcheck()">
                      <tr>
                        <td width="100"><p align="left">所属地区</td>
                        <td width="600">
<%set rs=conn.execute("select * from DNJY_city where id>0 and twoid>0 and threeid=0")%>
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
</td>
                      </tr>
                      <tr>
                        <td width="100"><p align="left">留言类型</td>
                        <td width="600"><select size="1" name="gbook">
                            <option value="0">《 普通问题 》</option>
                            <option  <%if ts="1" then%> selected <%end if%> value="1">《 投诉留言 》</option>
                        </select></td>
                      </tr>
				   <%if  strUserName<>"非注册用户" then%>	 
                      <tr>
                        <td width="100"><p align="left">是否公开</td>
                        <td width="600"><input type="radio" name="known" value="0" id="known" checked>
                 公开&nbsp;&nbsp;&nbsp;&nbsp;
                <input type="radio" name="known" value="1" id="known" >
                不公开 </td>
                      </tr>					     
                    <%end if%>
                      <tr>
                        <td width="100"><p align="left">留言内容</td>
                        <td width="600"><font color="#FF0000">字数限制为100个</font></td>
                      </tr>
                      <tr>
                        <td><p align="center"> </td>
                        <td><textarea rows="8" name="memo" cols="60" onkeyUp="textLimitCheck(this, 300);"><%=request("memo")%></textarea>&nbsp;&nbsp; <br>限 300 个字符 已输入 <font color="#CC0000"><span id="messageCount">0</span></font> 个字</td>
                      </tr>
					     </tr>
 <%if codekg1=1 Or codekg2=1 Or codekg3=1 Or codekg4=1 Or codekg5=1 then%>
                       <tr>
<td height="25" colspan="2">=======================<b>为防止群发垃圾信息，请耐心填写下列验证码</b>=======================</td>
                       </tr>
 <%End If
 if codekg1=1 then
 '问答式验证
dim rscheck
Randomize 
Set rscheck= Server.CreateObject("ADODB.Recordset")
If SystemDatabaseType = 1 Then
sql="select  * from DNJY_wenda where xs=1 order BY newid()"
Else
sql="select  * from DNJY_wenda where xs=1 order BY rnd(-(ID+"& rnd() &"))"
End if
rscheck.open sql,conn,1,1
%>
                        <tr>
                          <td height="25" width="100">问答验证：</td>
                          <td height="25" width="600">问题：<%=rscheck("wenti")%>&nbsp;&nbsp;&nbsp;答案：<input type="text" name="wenti" size="12"></td>      <input type="hidden" name="daan" value="<%=rscheck("daan")%>">                    
                        </tr>
	<%rscheck.close
	set rscheck=Nothing
	End If
	if codekg2=1  then%>
                        <tr>
                          <td height="25" width="100">数字验证：</td>
                          <td height="25" width="600"> <input type="text" name="yzm" size="4" /><img src="tt_getcode.asp" alt="验证码,看不清楚?请点击刷新验证码" height="10" style="cursor : pointer;" onClick="this.src='tt_getcode.asp?t='+(new Date().getTime());" />&nbsp;&nbsp; （验证码,看不清楚?请点击刷新验证码）</td>                          
                        </tr>
   <%End if
	if codekg3=1 then%>						
                        <tr>
                          <td height="25" width="100">文字验证：</td>
                          <td height="25" width="600"><input type="text" name="code" id="code" size="4" maxlength="4" tabindex="6" onfocus="get_Code();this.onfocus=null;" onkeyup="dv_ajaxcheck('checke_dvcode','code');"  autocomplete="off"/>
    <span id="imgid" style="color:red">点击获取验证码</span><span id="isok_code"></span></td>                          
                        </tr> 
   <%End if%>
<%if codekg4=1 then%>					
                        <tr>
                          <td height="25" width="100">星期验证：</td>
                          <td height="25" width="600">今天是 <select name="ttdv">
<option value="" selected="selected">请选择</option>
<option value="1">星期一</option>
<option value="2">星期二</option>
<option value="3">星期三</option>
<option value="4">星期四</option>
<option value="5">星期五</option>
<option value="6">星期六</option>
<option value="0">星期日</option>
</select>&nbsp;&nbsp; （请选择正确的星期几）</td>                          
                        </tr> 
<%End if%>
<%if codekg5=1 then%>					
                        <tr>
                          <td height="25" width="100">算式验证：</td>
                          <td height="25" width="600"><img src="tt_getcodee.asp" alt="验证码,看不清楚?请点击刷新验证码" height="10" style="cursor : pointer;" onClick="this.src='tt_getcodee.asp?t='+(new Date().getTime());" /><input name="validate_codee" type="text" tabindex="3" value="" size="4" maxlength="4" onkeyup="value=value.replace(/[^\d|^\.]/g,'')"   onbeforepaste="clipboardData.setData('text',clipboardData.getData('text').replace(/[^\d|^\.]/g,''))"/> &nbsp;&nbsp; （验证码,看不清楚?请点击刷新验证码）</td>                          
                        </tr> 
<%End if%>		
 <tr>
                        <td width="50">　</td>
                        <td><table width="100%"  border="0" cellspacing="0" cellpadding="0">
                            <tr>
                              <td height="30"><div align="center">
                                  <input class="inputb" border="0" name="I1" type="submit"   value="确认留言">
                              </div></td>
                              <td><div align="center">
                                  <input name="I2" type="reset" class="inputb" id="I2" value="取消留言" border="0">
                              </div></td>
                            </tr>
                        </table></td>
                      </tr>
                    </form>
                  <!---------------------->
                </table></td>
              </tr>
          </table></td>
      </tr>
    </table></td>
    <td width="4" bgcolor="#E6E6E6"></td>
  </tr>
</table>
</body>
</html>
<!--#include file=footer.asp-->

