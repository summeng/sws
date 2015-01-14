<!--#include file="conn.asp"-->
<!--#include file="SqlIn.Asp"-->
<!--#include file="Include/err.asp"-->
<!--#include file=usercookies.asp-->
<!--#include file=top.asp-->

<%
'=====================================
'天天供求信息网站管理系统
'天天网络科技工作室出品
'客服中心http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>

<link rel="stylesheet" type="text/css" href="/<%=strInstallDir%>css/css.css">
<style type="text/css">
<!--
.style2 {color: #FF0000; font-weight: bold; }
-->
</style>
<head>
<SCRIPT language=JavaScript>
function CheckAll(form)
  {
  for (var i=0;i<form.elements.length;i++)
    {
    var e = form.elements[i];
    if (e.name != 'chkall' )
       e.checked = form.chkall.checked;
    }
  }
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
//-->
</SCRIPT>
<body topmargin="0" leftmargin="0" >
<div align="center">
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

	  <table width="100%" border="0" align="left" cellpadding="0" cellspacing="0" bordercolor="#111111" style="border-collapse: collapse" class="ty1">
        <FORM name=thisForm action="info_del.asp?<%=C_ID%>" method=POST>
          <tr>
            <td width="100%" height="1" colspan="7"></td>
          </tr>
          <tr bgcolor="#FAFAFA">
            <td width="6%" height="25" align="center"><p class="style2">编号</td>
            <td width="50%" height="25" align="center"><span class="style2"> 标&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 题</span></td>
			<td width="10%" height="25" align="center"><span class="style2"> 阅读权限</span></td>
            <td width="10%" height="25" align="center"><span class="style2"> 浏览/回复</span></td>
            <td width="14%" height="25" align="center"><span class="style2"> 发布时间</span></td>
            <td width="4%" height="25" align="center"><span class="style2"> 审核</span></td>
            <td width="3%" height="25" align="center"><span class="style2"> 　</span></td>
            <td width="7%" height="25" align="center"><span class="style2"> 　</span></td>
          </tr>
          <tr bgcolor="#E4E4E4">
            <td height="1" colspan="8" align="center"></td>
          </tr>
          <%
dim rs1,sql1,b,bb,ii
dim ThisPage,Pagesize,Allrecord,Allpage,rsuser

set rsuser=server.createobject("adodb.recordset")
sql = "select * from [DNJY_user] where username='"&username&"'"
rsuser.open sql,conn,1,1
if rsuser.eof then
errdick(9)
response.end
end if
vip=rs("vip")
Call CloseRs(rsuser)

i=0
ii=1
if request("page")="" then
  ThisPage=1		
else
  ThisPage=request("page")
end if
username=request.cookies("dick")("username")
set rs = Server.CreateObject("ADODB.RecordSet")
sql="select id,biaoti,llcs,hfcs,fbsj,tupian,a,b,tj,yz,hfxx,fbts,hfje,Readinfo from [DNJY_info] where username='"&username&"' order by fbsj "&DN_OrderType&",b" 
rs.open sql,conn,1,1
if rs.eof then
response.write "<tr><td></td><td><li>没有记录</td></tr>"
else
rs.Pagesize=20
Pagesize=rs.Pagesize
Allrecord=rs.Recordcount
Allpage=rs.Pagecount
if ThisPage<1 then                           
ThisPage=1
end if
On Error Resume Next
rs.move (ThisPage-1)*Pagesize
do while not rs.eof
b=trim(rs("b"))
bb=len(b)
id=rs("id")
%>
          <tr>
            <td width="6%" height="26" align="center" bordercolor="#E1E1E1" background="img/22.gif"><%=i+1%> </td>
            <td width="50%" height="26" align="center" bordercolor="#E1E1E1" background="img/22.gif"><p align="left">
			<%if rs("tupian")<>"0" then
            response.write "<img src=""images/num/pic.gif"" alt=""有贴图"">" 
           end If
           if rs("yz")=1 then%>
           <a target="_blank" href=<%=x_path("",id,"",rs("Readinfo"))%>>
            <%else%>
           <a target="_blank" href=show_user.asp?id=<%=rs("id")%>&<%=C_ID%>>
			<%end if
			if rs("a")="0" then%>
                <%=rs("biaoti")%>
                <%else%>
                <font color="#<%=rs("a")%>"><%=rs("biaoti")%></font>
                <%end if%>
                </a>
                    <%if b<>0 then%>
                    <img src="images/num/jsq.gif" alt="置顶信息">
                    <%for ii=1 to bb%>
                    <img src="images/num/<%=Mid(b,ii,1)%>.gif" alt="<%=Mid(b,ii,1)%>级置顶">
                    <%next%>
                    <%end If
                    if rs("tj")>0 then
                    response.write "<img src=""images/jian.gif"" alt=""推荐信息"">" 
                    end if%>
[<%dim zt
zt=rs("zt")
if zt=1 then
response.write "<span title=""已完成交易"" style=""cursor:help""><font color=""#ff0000"">完成</font></span>"
elseif zt=0 then
response.write "<span title=""未完成交易"" style=""cursor:help""><font color=""#0066FF"">未完</font></span>"
end if%> ]
<%if rs("hfxx")=1 Then
response.write " <span title=""竞价平均每天"&FormatNumber(rs("hfje")/rs("fbts"),2,-1)&"元"" style=""cursor:help"">["&FormatNumber(rs("hfje")/rs("fbts"),2,-1)&"]</span>" 
elseif rs("hfxx")=2 Then
response.write " <span title=""竞价尚未通过审核"" style=""cursor:help""><font color=#FF0000>["&FormatNumber(rs("hfje")/rs("fbts"),2,-1)&"]</font></span>" 
end if%>
</td>
<td width="10%" height="26" align="center" bordercolor="#E1E1E1" background="img/22.gif"">
<%if rs("Readinfo")=1 Then
response.write "普通会员"
elseif rs("Readinfo")=2 Then
response.write "VIP会员"
Else
response.write "游客"
End if
%></td>
            <td width="10%" height="26" align="center" bordercolor="#E1E1E1" background="img/22.gif""><%=rs("llcs")%>/<%=rs("hfcs")%></td>
            <td width="14%" height="26" align="center" bordercolor="#E1E1E1" background="img/22.gif"><%=datevalue(rs("fbsj"))%> </td>
            <td width="3%" height="26" align="center" bordercolor="#E1E1E1" background="img/22.gif"><%if rs("yz")=1 then%>
                <font color="#008000">√</font>
                <%else%>
                <font color="#FF0000"><span title="此条信息尚未审核,请耐心等待" style="cursor:help">ⅹ</span></font>
                <%end if%>
            </td>
            <td width="6%" height="26" align="center" bordercolor="#E1E1E1" background="img/22.gif"><font color="#666666"> <span id="followImg<%=i%>" style="CURSOR: hand" onClick="loadThreadFollow(<%=i%>,5)">管理</span></font></td>

            <td width="7%" height="26" align="center" bordercolor="#E1E1E1" background="img/22.gif"><input type="checkbox" name="selectedid" value="<%=id%>"></td>
          </tr>
          <tr style="display:none" id="follow<%=i%>">
            <td width="100%" height="25" align="center" style="border-bottom-style: none; border-bottom-width: medium" colspan="7"><table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%" height="100%">
                <tr bgcolor="#FFCCFF">
                  <td width="17%" align="center" bordercolor="#E1E1E1" style="border-bottom-style: solid; border-bottom-width: 1"><font color="#CC5200"> <a href="info_del.asp?selectedid=<%=id%>&<%=C_ID%>" onClick="return confirm('您确定进行删除操作吗？')"> <font color="#FF0000">删除</font></a></font></td>
                  <td width="17%" align="center" bordercolor="#E1E1E1" style="border-bottom-style: solid; border-bottom-width: 1"><a href="edit_info.asp?id=<%=id%>&<%=C_ID%>">修改</a></td>
                  <td width="17%" align="center" bordercolor="#E1E1E1" style="border-bottom-style: solid; border-bottom-width: 1"><a href="#" ONCLICK="window.open('edit_info_color.asp?id=<%=id%>&','Sample','toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=no,resizable=no,copyhistory=no,width=200,height=150,left=300,top=100')">标题变色</a></td>
                  <td width="17%" align="center" bordercolor="#E1E1E1" style="border-bottom-style: solid; border-bottom-width: 1"><a href="#" ONCLICK="window.open('edit_info_top.asp?id=<%=id%>','Sample','toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=no,resizable=no,copyhistory=no,width=210,height=150,left=300,top=100')">信息置顶</a></td>
                  <td width="14%" align="center" bordercolor="#E1E1E1" style="border-bottom-style: solid; border-bottom-width: 1"><a href="#" ONCLICK="window.open('edit_info_picture.asp?id=<%=id%>','Sample','toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=no,resizable=no,copyhistory=no,width=210,height=150,left=300,top=100')">内容贴图</a></td>
                 
                  <td width="18%" align="center" bordercolor="#E1E1E1" style="border-bottom-style: solid; border-bottom-width: 1"><%if rs("yz")=1 then%>
                      <font color="#0000ff">已验证</font>
                      <%else
                      if onoff>0 and (vip>0 or vipno>0) then%>
                      <a href="#" ONCLICK="window.open('edit_info_audit.asp?id=<%=id%>','Sample','toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=no,resizable=no,copyhistory=no,width=210,height=150,left=300,top=100')">使用验证道具             <%else%>
                    <font color="#CC0000">非VIP需审核</font>  
                        <%end if
                        end if%>
                    </a></td>
                </tr>
            </table></td>
          </tr>
          <%
i=i+1
rs.movenext
if i>=Pagesize then exit do
loop
end if
%>
          <tr>
            <td width="10%" height="25" colspan="7"><table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%">
                <tr>
                  <td width="25%" height="25"><p align="center"> 　</td>
                  <td width="760" height="25"><p align="center">　</td>
                  <td width="760" height="25"><p align="right">　</td>
                  <td width="30%" height="25"><p align="center">
                      <INPUT onclick=CheckAll(this.form) type=checkbox value=on name=chkall>
                    选中所有记录
                    <input  type="submit" value="删除" name="B1" onClick="return test();">
                    </td>
                </tr>
                <tr>
                  <td height="25" align="center"> 共有&nbsp;<font color="#CC5200"><%=Allrecord%></font>&nbsp;条记录</td>
                  <td height="25" align="center"> 共 <font color="#CC5200"><%=Allpage%></font> 页</td>
                  <td height="25" align="center"> 现在是第 <font color="#CC5200"><%=ThisPage%></font> 页</td>
                  <td height="25" align="center"><%
if ThisPage<2 then     
response.write "<font color=""#808080"">首页</font>&nbsp;"
response.write "<font color=""#808080"">上一页</font>&nbsp;"     
else     
response.write "<a href=?page=1>首页</a>&nbsp;"
response.write "<a href=?page="&ThisPage-1&">上一页</a>&nbsp;"     
end if
if Allpage-ThisPage<1 then     
response.write "<font color=""#808080"">下一页</font>&nbsp;"
response.write "<font color=""#808080"">尾页</font>&nbsp;"  
else     
response.write "<a href=?page="&(ThisPage+1)&">下一页</a>&nbsp;"   
response.write "<a href=?page="&Allpage&">尾页</a>&nbsp;"     
end if
%></td>
                </tr>
                
            </table></td>
          </tr>
        </form>
      </table></td>
      <td width="4" align="center" valign="top" bgcolor="#e6e6e6"></td>
    </tr>
    <tr>
<%
Call CloseRs(rs)%>
      <td width="980" height="26" colspan="3"><!--#include file=footer.asp--></td>
    </tr>
  </table>
</div>
</body>
</html>

<script language=javascript>
function test()
{
  if(!confirm('确认删除吗？')) return false;
} 
</script>
