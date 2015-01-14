
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
    <td height="30" align="center"class="ty3"><b>站内搜索</b></td>
  </tr>
  <tr>
    <td align="center"><table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#f7f7f7" class="tab_k">
      <form action="newssearch.asp" method="get" name="form1" id="form1">
        <tr>
          <td height="33" align="center"class="dq2"><input name="key" type="text" class="input" id="key" size="19" /></td>
        </tr>
        <tr>
          <td height="35" align="center" class="dq2"><select name="otype" class="input" id="otype" style="line-height:30px;">
            <option value="title" selected="selected" class="input">标题</option>
            <option value="msg" class="input">内容</option>
          </select>
		  
		    <input name="bigcid" type="hidden" class="input" id="bigcid" value="<%=clng(request.querystring("classid"))%>" size="19" />
			<input name="c1" type="hidden" class="input" id="c1" value="<%=c1%>" size="19" />
			<input name="c2" type="hidden" class="input" id="c2" value="<%=c2%>" size="19" />
			<input name="c3" type="hidden" class="input" id="c3" value="<%=c3%>" size="19" />
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
    <td align="center" height="30" class="ty3"><b>最新推荐</b></td>
  </tr>
  <tr>
    <td colspan="2" align="center"><%=news(c1,c2,c3,1,0,0,0,1,0,0,0,0,0,8,1,24,13,20)%></td>
  </tr>     
</table>
      <table width="220" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td  height="7"></td>
        </tr>
      </table>
<table width="220" border="0" cellpadding="0" cellspacing="0"  class="ty1">
  <tr>
    <td align="center" height="30" class="ty3"><b>热门点击</b></td>
  </tr>
  <tr>
    <td height="22" align="center"><%=news(c1,c2,c3,3,0,0,0,1,0,0,0,0,0,8,1,24,13,20)%></td>
  </tr>
</table>
