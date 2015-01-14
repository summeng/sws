<!--#include file="conn.asp"-->
<!--#include file="SqlIn.Asp"-->
<!--#include file=top.asp-->
<!--#include file=usercookies.asp-->
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
.style2 {color: #333333}
.style3 {color: #FF0000}
.style5 {color: #FF0000; font-weight: bold; }
-->
</style>
</head>

<body topmargin="0" leftmargin="0">

<div align="center">
  <center>
  <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="980" height="140">
    <tr>
      <td width="214" height="356" valign="top" bgcolor="#FFFFFF"><!--#include file=userleft.asp--></td>
	  <td width="5">&nbsp;</td>
      <td width="760" bgcolor="#FFFFFF"><table width="760" height="375" border="0" align="right" cellpadding="0" cellspacing="0" style="border-collapse: collapse" class="ty1">
        <tr bgcolor="#FAFAFA">
          <td height="25" colspan="7" ><span class="style5">　赠送<%=rmb_mc%>公告：本站于<%=xnpsj%> 给每个VIP会员赠送<%=rmb_mc%><%=xnp%>个。</span></td>
        </tr>
<%dim m,integrity
username=request.cookies("dick")("username")
set rs=conn.execute("select count(id) from [DNJY_info] where username='"&username&"'")
m=rs(0)
rs.close
set rs=server.createobject("adodb.recordset")
sql = "select * from [DNJY_user] where username='"&username&"'"
rs.open sql,conn,1,1
%>

        <tr>
          <td width="16%" height="28" align="right" background="img/22.gif"><img src="images/File_29.gif" width="16" height="16" border="0">帐户余额：</td>
          <td height="25" colspan="2" align="left" valign="middle" background="img/22.gif"><%=rs("Amount")%>元</td>
          <td width="15%" height="25" align="right" background="img/22.gif"><%=rmb_mc%>：</td>
          <td width="10%" height="25" align="left" valign="middle" background="img/22.gif"><%=rs("hb")%>个</td>
          <td width="16%" height="25" align="right" background="img/22.gif">积分总数：</td>
          <td width="31%" height="25" align="left" valign="middle" background="img/22.gif"><%=rs("jf")%></td>
        </tr>
		
        <tr>
          <td width="16%" height="28" align="right" background="img/22.gif"><img src="images/File_29.gif" width="16" height="16" border="0">登陆次数：</td>
          <td height="25" colspan="2" align="left" valign="middle" background="img/22.gif"><%=rs("dlcs")%>次</td>
          <td width="15%" height="25" align="right" background="img/22.gif">发布信息：</td>
          <td width="10%" height="25" align="left" valign="middle" background="img/22.gif"><%=m%>条</td>

          <td width="16%" height="25" align="right" background="img/22.gif">会员级别：</td>
          <td width="31%" height="25" align="left" valign="middle" background="img/22.gif"><%if vip>0 then
  response.write "<font color=#FF0000>VIP会员</font>"
if sj>0 then
response.Write "<font color=#FF0000>(资格剩余"&sj&"</b>天)</font>"
elseif sj=0 then
response.Write "<font color=""#414141"">(资格今日到期)</font>"
elseif sj<0 then
response.Write "<font color=""#FF0000""> (资格已经过期)</font>"
end If
else
response.write "<font color=#FF0000>普通会员</font>"
response.write "(建议升级<a href='upgradevip.asp?"&C_ID&"'><font color=#0000ff>VIP会员</font></a>)"
end if%> <%if rs("sdname")<>"" then%><a  href="user/com.asp?com=<%=rs("username")%>&<%=C_ID%>" target="_blank"><img src="img/dian.gif" title="您已开有网店" border="0"></a><%end if%> <%if rs("mpname")<>"" then%> <a  href="company.asp?username=<%=rs("username")%>&<%=C_ID%>" target="_blank"><img src="img/qy.gif" title="您已发布企业名片" border="0"></a><%end if%></td>
        </tr>
        <tr>
          <td colspan="7" width="16%" height="28"  background="img/22.gif">&nbsp;&nbsp;&nbsp;<img src="images/File_29.gif" width="16" height="17" border="0">注册时间：<%=rs("zcdata")%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;登陆时间：<%=rs("dldata")%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;最近登录：<%=request.cookies("dick")("dldata")%></td>
        </tr>
        <tr>
          <td colspan="7" width="16%" height="28"  background="img/22.gif">&nbsp;&nbsp;&nbsp;<img src="images/File_29.gif" width="16" height="17" border="0">诚信记录：<%=rs("goodfaith")%> 次 <img src="img/hp.jpg">&nbsp;&nbsp;&nbsp;&nbsp; 失信记录：<%=rs("bpromises")%> 次 <img src="img/cp.jpg"></td>
        </tr>    		
        <tr>
          <td width="16%" height="28" align="right" background="img/22.gif" valign="top"><img src="images/File_29.gif" width="16" height="16" border="0">网友参评：</td>
          <td height="25"  align="left" valign="middle" background="img/22.gif" valign="top"></td><%Dim df,cs,pj
df=rs("wevaluation")
cs=rs("participants")
If rs("wevaluation")=0 Then df=5
If rs("participants")=0 Then cs=1
pj=Formatnumber(df/(cs*5)*100,0,0,0,0)%>
          <td width="40%" align="left" background="img/22.gif" valign="top"> <%=cs%> 人次&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;平均得分 <%=Formatnumber(df/cs,1,0,0,0)%> 分</td>
          
          <td height="25" align="left" valign="middle" background="img/22.gif" colspan="6">	<style type="text/css"> 
body{ 
text-align:center; 
} 
.graph{ 
width:355px; 
border:1px solid #F8B3D0; 
height:15px; 
} 
#bar{ 
display:block; 
background:#FFE7F4; 
float:left; 
height:100%; 
text-align:center; 
} 
#barNum{ 
position:absolute; 
} 
</style> 
<script type="text/javascript"> 
function $(obj){ 
return document.getElementById(obj); 
} 
function go(){ 
$("bar").style.width = parseInt($("bar").style.width) + 1 + "%"; 
$("bar").innerHTML = $("bar").style.width; 
if($("bar").style.width =="<%=pj%>%"){ 
window.clearInterval(bar); 
} 

} 
var bar = window.setInterval("go()",40); 
window.onload = function(){ 
bar; 
} 
</script> 
<div class="graph"><strong id="bar" style="width:1%;"></strong></div>
<img src="img/pj.jpg" width="370"></td>
        </tr>
    
        <tr>
          <td height="10" colspan="7"></td>
        </tr>

        <tr>
          <td height="25" colspan="7"><table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="760" height="26">
              <tr>
                <td width="100%" height="26"><table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%">
                    <tr>
                      <td width="100%" height="10"></td>
                    </tr>
                    <tr>
                      <td height="5" bgcolor="#66CC00"></td>
                    </tr>
                    <tr>
                      <td height="8" bgcolor="#EEEEEE"></td>
                    </tr>
                    <tr>
                      <td height="10"></td>
                    </tr>
                </table></td>
              </tr>
          </table></td>
        </tr>
        <tr>
          <td height="25" colspan="7"><table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="99%" height="17">
              <tr>
                <td height="20"><div align="center">注册资料完整度(建议完善资料方便以后发布信息) [ <a href="user_data.asp?<%=C_ID%>">修改资料</a>]</div></td>
              </tr>
              <tr>
                <td height="10">
				密码保护<%If rs("problem")<>"" Then
                response.Write "<font color=#66CC00><b>√</b></font>"
				integrity=integrity+1
				Else
				response.Write "×"
				End if%>
				地区分类<%If rs("city_oneid")>0 Then
                response.Write "<font color=#66CC00><b>√</b></font>"
				integrity=integrity+1
				Else
				response.Write "×"
				End if%>
				信息分类<%If rs("type_oneid")>0 Then
                response.Write "<font color=#66CC00><b>√</b></font>"
				integrity=integrity+1
				Else
				response.Write "×"
				End if%>
				真实姓名<%If rs("name")<>"" Then
                response.Write "<font color=#66CC00><b>√</b></font>"
				integrity=integrity+1
				Else
				response.Write "×"
				End if%>
				性&nbsp;&nbsp;&nbsp;&nbsp;别
				<%response.Write "<font color=#66CC00><b>√</b></font>"
				integrity=integrity+1%>
				身 份 证<%If rs("idcard")<>"" Then
                response.Write "<font color=#66CC00><b>√</b></font>"
				integrity=integrity+1
				Else
				response.Write "×"
				End if%>
				固定电话<%If rs("dianhua")<>"" Then
                response.Write "<font color=#66CC00><b>√</b></font>"
				integrity=integrity+1
				Else
				response.Write "×"
				End if%>
				手机号码<%If rs("CompPhone")<>"" Then
                response.Write "<font color=#66CC00><b>√</b></font>"
				integrity=integrity+1
				Else
				response.Write "×"
				End if%>
				传 真 机<%If rs("fax")<>"" Then
                response.Write "<font color=#66CC00><b>√</b></font>"
				integrity=integrity+1
				Else
				response.Write "×"
				End if%>
				电子邮箱<%If rs("email")<>"" Then
                response.Write "<font color=#66CC00><b>√</b></font>"
				integrity=integrity+1
				Else
				response.Write "×"
				End if%>
				联系 QQ<%If rs("qq")<>"" Then
                response.Write "<font color=#66CC00><b>√</b></font>"
				integrity=integrity+1
				Else
				response.Write "×"
				End if%>
				<br>微信号码<%If rs("weixin")<>"" Then
                response.Write "<font color=#66CC00><b>√</b></font>"
				integrity=integrity+1
				Else
				response.Write "×"
				End if%>
				公司名称<%If rs("mpname")<>"" Then
                response.Write "<font color=#66CC00><b>√</b></font>"
				integrity=integrity+1
				Else
				response.Write "×"
				End if%>
				公司网址<%If rs("http")<>"" Then
                response.Write "<font color=#66CC00><b>√</b></font>"
				integrity=integrity+1
				Else
				response.Write "×"
				End if%>
				通信地址<%If rs("dizhi")<>"" Then
                response.Write "<font color=#66CC00><b>√</b></font>"
				integrity=integrity+1
				Else
				response.Write "×"
				End if%>
				邮政编码<%If rs("youbian")<>"" Then
                response.Write "<font color=#66CC00><b>√</b></font>"
				integrity=integrity+1
				Else
				response.Write "×"
				End if%>
				地图标注<%If rs("Emap")<>"" Then
                response.Write "<font color=#66CC00><b>√</b></font>"
				integrity=integrity+1
				Else
				response.Write "×"
				End if%>
				订阅资讯<%response.Write "<font color=#66CC00><b>√</b></font>"
				integrity=integrity+1%>				
				</td>
              </tr>
              <tr>
                <td height="20">
<%Dim wzd
wzd=Formatnumber(integrity/18*100,0,0,0,0)%>
<style type="text/css"> 
body{ 
text-align:center; 
} 
.graph2{ 
width:650px; 
border:1px solid #F8B3D0; 
height:15px; 
} 
#wzd{ 
display:block; 
background:#FFE7F4; 
float:left; 
height:100%; 
text-align:center; 
} 
#wzdNum{ 
position:absolute; 
} 
</style> 
<script type="text/javascript"> 
function $(objwzd){ 
return document.getElementById(objwzd); 
} 
function go2(){ 
$("wzd").style.width = parseInt($("wzd").style.width) + 1 + "%"; 
$("wzd").innerHTML = $("wzd").style.width; 
if($("wzd").style.width =="<%=wzd%>%"){ 
window.clearInterval(wzd); 
} 

} 
var wzd = window.setInterval("go2()",40); 
window.onload = function(){ 
wzd; 
} 
</script> 
<div class="graph2"><strong id="wzd" style="width:1%;"></strong></div>
<img src="img/wzd.jpg" width="675"></td>
              </tr>
          </table></td>
        </tr>

        <tr>
          <td height="31" colspan="7"><table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%">
              <tr>
                <td width="100%" height="5" bgcolor="#FFCC33"></td>
              </tr>
              <tr>
                <td height="8" bgcolor="#EEEEEE"></td>
              </tr>
              <tr>
                <td height="10"></td>
              </tr>
          </table></td>
        </tr>
        <tr>
          <td width="16%" height="25" valign="top"><div align="right">标题<font color="#CC5200">变色</font>道具：</div></td>
          <td height="25" colspan="2" align="right" valign="top"><div align="center"><%=rs("a")%>&nbsp;个</div></td>
          <td height="25" colspan="4"><span class="style2"> <img src="images/form2_r3_c3.gif"> 当你拥有此道具时可以设定发布的信息标题颜色，以次来获取浏览用户的视线！</span></td>
        </tr>
        <tr>
          <td width="16%" height="25" valign="top"><div align="right">信息<font color="#CC5200">置顶</font>道具：</div></td>
          <td height="25" colspan="2" align="right" valign="top"><div align="center"><%=rs("b")%>&nbsp;个</div></td>
          <td height="25" colspan="4"><span class="style2"> <img src="images/form2_r3_c3.gif"> 使用越多你的信息将在首页（分类页）最前面显示，但有效期只有一天！</span></td>
        </tr>
        <tr>
          <td width="16%" height="25" valign="top"><div align="right">内容<font color="#CC5200">贴图</font>道具：</div></td>
          <td height="25" colspan="2" align="right" valign="top"><div align="center"><%=rs("c")%>&nbsp;个</div></td>
          <td height="25" colspan="4"><span class="style2"> <img src="images/form2_r3_c3.gif"> 使用后可以在信息内容里插入图片，以便用户了解你的商品！</span></td>
        </tr>
        <tr>
          <td width="16%" height="25" valign="top"><div align="right">通过<font color="#CC5200">验证</font>道具：</div></td>
          <td height="25" colspan="2" align="right" valign="top"><div align="center"><%=rs("d")%>&nbsp;个</div></td>
          <td height="25" colspan="4"><span class="style2"> <img src="images/form2_r3_c3.gif"> 使用后不经过管理员验证而直接显示在我们网站里，自助操作但购买时价格较高！</span></td>
        </tr>
        <tr>
          <td height="25" colspan="7"><table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="760" height="26">
              <tr>
                <td width="100%" height="26"><table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%">
                    <tr>
                      <td width="100%" height="10"></td>
                    </tr>
                    <tr>
                      <td height="5" bgcolor="#66CC00"></td>
                    </tr>
                    <tr>
                      <td height="8" bgcolor="#EEEEEE"></td>
                    </tr>
                    <tr>
                      <td height="10"></td>
                    </tr>
                </table></td>
              </tr>
          </table></td>
        </tr>
        <tr>
          <td height="25" colspan="7"><table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%" height="17">
              <tr>
                <td width="185" height="48" background="img/alipaybutton.gif" ><div align="center"><a href="props.asp?<%=C_ID%>"><br>我要用人民币兑换<%=rmb_mc%></a></div></td>
                <td width="75%" height="13" valign="top"><span class="style2"> 使用道具您将获得更多权限，道具要使用<%=rmb_mc%>兑换，如果您的<%=rmb_mc%>不足，可以用人民币兑换<%=rmb_mc%>。一元人民币可兑换<font color=#ff0000><%=rmb_hb%></font>个<%=rmb_mc%>。请点击左边按钮兑换。</span> </td>
              </tr>
          </table></td>
        </tr>
        <tr>
          <td height="25" colspan="7"><table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="760" height="1">
              <tr>
                <td height="26" colspan="2"><table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%">
                    <tr>
                      <td width="100%" height="10"></td>
                    </tr>
                    <tr>
                      <td height="5" bgcolor="#CC99FF"></td>
                    </tr>
                    <tr>
                      <td height="8" bgcolor="#EEEEEE"></td>
                    </tr>
                    <tr>
                      <td height="10"><script src="<%=adjs_path("ads/js","info1",c1&"_"&c2&"_"&c3)%>"></script></td>
                    </tr>
                </table></td>
              </tr>

          </table></td>
        </tr>
      </table></td>
      <td width="10" valign="top" bgcolor="#e6e6e6"></td>
    </tr>
    <tr><%Call CloseRs(rs)%>
      <td height="26" colspan="3"><!--#include file=footer.asp--></td>
    </tr>
  </table>
  </center>
</div>
</body>
</html>

