<table width="980" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td height="5"></td>
  </tr>
</table>
    <TABLE width=980 border=0 align="center" cellPadding=0 cellSpacing=0 background="skin/ltby/btmbg.gif"  style="BORDER-COLLAPSE: collapse">
      <TBODY>
        <TR> 
          <TD height=10 width="758"> </TD>
        </TR>
        <TR> 
          
      <TD align="center" class="p9black"> <a href="help.asp?id=0&<%=C_ID%>">关于本站</a><a href="admin/" target="_blank"><font color="#948A80">‖</font></a><a href="help.asp?id=1&<%=C_ID%>">会员帮助</a>‖<a href="help.asp?id=2&<%=C_ID%>">服务条款</a>‖<a href="help.asp?id=3&<%=C_ID%>">免责声明</a>‖<a href="help.asp?id=4&<%=C_ID%>">广告服务</a>‖<a href="help.asp?id=5&<%=C_ID%>">联系我们</a>‖<a href="Message_board.asp">客户留言</a><%If contribute>0 Then Response.Write "‖<a href=""#"" ONCLICK=""window.open('user/contribute.asp','Sample','toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=yes,resizable=no,copyhistory=no,width=780,height=700,left=300,top=20')"">我要投稿</a>"%>‖<a href="#" onClick="javascript:window.external.AddFavorite('http://<%=web%>', '<%=title%>')" target="_self">收藏本站</a>‖<a href="cityadmin/login.asp?<%=C_ID%>" target="_blank">分站管理登录</a><br>
      </TD>
        </TR>
        <TR> 
          <TD height=90 > 
          <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0"  style="line-height:130%;letter-spacing:1px">
          <tr>
              <td height="22" align="center"><span class="p9black">          <p  style="line-height: 150%; margin-top: 0; margin-bottom: 0">
版权所有 &copy; 2006-2015 <b><a href="http://<%=cityweb%>"><%=title%></a>  <%=cityCompany%></b><br>
	<%if Cellular_phone<>"" then%>手机:<%=Cellular_phone%><%end if%> &nbsp; <%if tel<>"" then%>电话:<%=tel%><%end if%><%if fax<>"" then%>    传真:<%=fax%><%end if%><%if mymail<>"" then%>    

<%dim ttemail,ttmail,fqq,fqq_name
fqq=qq
fqq_name=qq_name
	ttemail=Replace(""&mymail&"","@","&")
ttmail=Replace(""&mymail&"","@","<IMG SRC=images/@.gif BORDER=0 width=5 height=8>")%> 客服信箱：<a href="javascript://" onClick="sendmail('<%=ttemail%>?subject=我的建设和意见');return false"><%=ttmail%></a><%end if%>    
<%'客服QQ此文件不需要作任何修改
if fqq<>"" Then 
fqq=replace(fqq,"，",",")
if isnull(fqq_name) or fqq_name="" then fqq_name=fqq
fqq_name=replace(fqq_name,"，",",")
fQQ_NAME=split(fqq_name,",")
fQQ=split(fqq,",")
%> 服务QQ：<%for N=0 to UBound(fQQ)%>
<a class='c' target=blank href='tencent://message/?uin=<%=fQQ(n)%>&Site=<%=title%>&Menu=yes'><%=fQQ(n)%></a>
</script >
<%next%><%end if%>     <%if serve<>"" then%><br>办公地址：<%=serve%><%end if%>  <%if yzbm<>"" then%> 邮政编码：<%=yzbm%><%end if%> <a href="http://www.miibeian.gov.cn" target=_blank><%=icp%></a>  </span></td>
            </tr>
            
            <tr>
              <td height="22" align="center"><SCRIPT language="">
var message="法律声明：本站免费提供信息交流平台，交易者自行甄别信息真伪，责任自负，如有损失，本站概不负责。"
var n=0; 
if (document.all){ 
document.write('<font size="12px" face="Arial" color="#EAEAEA">') 
for (m=0;m<message.length;m++) 
document.write('<span id="neonlight" style="font-size:12px">'+message.charAt(m)+'</span>') 
document.write('</font>') 
var tempref=document.all.neonlight 
} 
else 
document.write(message) 
function neon(){ 
if (n==0){ 
for (m=0;m<message.length;m++) 
tempref[m].style.color="#0000FF" 
} 
tempref[n].style.color="#FF0000" 
if (n<tempref.length-1) 
n++ 
else{ 
n=0 
clearInterval(flashing) 
setTimeout("beginneon()",3000) 
return 
} 
} 
function beginneon(){ 
if (document.all) 
flashing=setInterval("neon()",50) 
} 
beginneon() 
 </SCRIPT></td>
            </tr>
          </table>
          </TD>
        </TR>
      </TBODY>
</TABLE>
<%Call CloseDB(conn)%>