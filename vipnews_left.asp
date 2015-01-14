
<%
'=====================================
'天天供求信息网站管理系统
'天天网络科技工作室出品
'客服中心http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>
<table width="220" border="0" cellpadding="0" cellspacing="0" class="ty1">
  <tr>
    <td height="30" align="center"class="ty3">站内搜索</td>
  </tr>
  <tr>
    <td align="center"><table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#f7f7f7" class="tab_k">
      <form action="vipnewssearch.asp?t=<%=id%> method="get" name="form1" id="form1">
        <tr>
          <td height="33" align="center"class="dq2"><input name="key" type="text" class="input" id="key" size="19" /></td>
        </tr>
        <tr>
          <td height="35" align="center" class="dq2"><select name="otype" class="input" id="otype" style="line-height:30px;">
            <option value="title" selected="selected" class="input">标题</option>
            <option value="msg" class="input">内容</option></select>		  
		    <input name="t" type="hidden" class="input" id="id" value="<%=clng(request.querystring("id"))%>" size="19" />
            <input type="submit" name="submit2" value="搜索" class="input" /></td>
        </tr>
      </form>
    </table></td>
  </tr>
</table>
      <table width="220" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td  height="7"></td>
        </tr>
      </table>
<table width="220" border="0" cellpadding="0" cellspacing="0" class="ty1">
  <tr>
    <td align="center" height="30" class="ty3">最新推荐</td>
  </tr>
  <tr>
    <td colspan="2" align="center"><%=vipnews("",0,1,1,1,1,0,0,20,15,1,15,10,0)%></td>
  </tr>     
</table>
      <table width="220" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td  height="7"></td>
        </tr>
      </table>
<table width="220" border="0" cellpadding="0" cellspacing="0"  class="ty1">
  <tr>
    <td align="center" height="30" class="ty3">热门点击</td>
  </tr>
  <tr>
    <td height="22" align="center"><%=vipnews("",0,2,1,1,0,1,0,20,15,1,15,10,0)%></td>
  </tr>
</table>
