<!--#include file="conn.asp"-->
<!--#include file="SqlIn.Asp"-->
<!--#include file=top.asp-->

<%
'=====================================
'天天供求信息网站管理系统
'天天网络科技工作室出品
'客服中心http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>

<style type="text/css">
<!--
.style1 {color: #FF0000}
-->
</style>
</head>


<table width="980" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td width="980" valign="top"><table width="980" border="0" cellpadding="0" cellspacing="0"  bgcolor="#ffffff">
      <tr>
        <td height="414" valign="top">   <div>
            <br />
            <table width="100%" height="354" border="0" align="center" cellpadding="0" cellspacing="1" class="PayPut1">
              <tr valign="middle">
                <td height="25" nowrap="nowrap" bgcolor="#EDDAFF"><div class="style36"> <span class="aa  TopicShow5"> &nbsp;&gt;&gt; 发布信息提示：请先选择发布类型 </span></div></td>
              </tr>
			  <tr><td width=980 height=120 align="center"><img src=img/reg.jpg border=0></td></tr>
              <tr valign="middle">
                <td height="40" nowrap="nowrap"><table width="543" height="283"  border="0" align="center" cellpadding="0" cellspacing="16">
                  <tr>
                    <td width="198" height="45" background="img/cardTop.gif"><div align="center" class="style34 STYLE31 STYLE31"><a href="<%=addinfo%>?<%=CT_ID%>&ttop=2"><span style="color:#FF6600 ">游客发布</span></a></div></td>
                    <td width="83"  >&nbsp;</td>
                    <td width="198" background="img/cardTop.gif"><div align="center" class="style34" ><a href="<%=reg%>?<%=CT_ID%>&ttop=10"><span style="color:#FF6600 ">注册发布</span></a></div></td>
                    <td width="83"  >&nbsp;</td>
                    <td width="198" background="img/cardTop.gif"><div align="center" class="style34" ><a href="help.asp?id=1&<%=CT_ID%>&ttop=10"><span style="color:#FF6600 ">VIP发布</span></a></div></td>
                  </tr>
                  <tr>
                    <td width="198" height="126" valign="top" bgcolor="#F7FFFF" class="tb1"><span style="line-height:200%;margin-left:5px"><span>
                      <table width="180"  border="0" align="center" cellpadding="0" cellspacing="0">
                        <tr>
                          <td height="10"></td>
                        </tr>
                        <tr>
                          <td>
                            <li>游客发布信息<%If onoff =1 And xinxish=1 then%>免审核<%else%>须审核<%End if%></li>
                            <li>无积分机制</li>
                            <li>管理员随时无条件删除信息</li>
                            <li>无置顶功能，本站不作推荐</li>
                            <li>标题无变色的功能</li>
							<li>无上传图片功能，贴图无效</li>
                            <li>发布信息时内容框非HTM编辑器</li>
                            <li>已发布信息不能修改和删除</li>
							<li>发布信息不能按用户名查阅</li>
							<li>不能申请开通网店或企业名片</li>
                            <li>每次发布信息都要填联络信息</li></td>
                        </tr>
                      </table>
                    </span></span> </td>
                    <td width="83" valign="top"  ></td>
                    <td width="198" valign="top" bgcolor="#F7FFFF" class="tb1"><span style="line-height:200%;margin-left:5px">
                      <table width="180"  border="0" align="center" cellpadding="0" cellspacing="0">
                        <tr>
                          <td height="10"></td>
                        </tr>
                        <tr>
                          <td>
                            <li>普通会员发布信息<%If onoff =1 And vipno=1 then%>免审核<%else%>须审核<%End if%></li>
                            <li>有积分机制</li>
                            <li>信息不过期不违规不会被删除</li>
                            <li>有置顶功能，本站不作推荐</li>
                            <li>标题可变色，突出显示</li>
							<li>可上传图片一张，无贴图功能</li>
                            <li>发布信息时内容框为HTM编辑器</li>
                            <li>已发布信息可随时修改删除</li>
							<li>信息可按名查阅，但仅显示<%=huiyuanxx%>条</li>
							<li>可开网店名片,信息仅显示<%=huiyuansp%>条</li>
                            <li>每次发布信息不用填联络信息</li></td>
                        </tr>
                      </table>
                    </span></td>
                    <td width="83" valign="top"  ></td>
                    <td width="198" valign="top" bgcolor="#F7FFFF" class="tb1"><span style="line-height:200%;margin-left:5px">
                      <table width="180"  border="0" align="center" cellpadding="0" cellspacing="0">
                        <tr>
                          <td height="10"></td>
                        </tr>
                        <tr>
                          <td>
                            <li>VIP会员发布信息<%If onoff=1 then%>免审核<%else%>须审核<%End if%></li>
                            <li>有积分机制</li>
                            <li>信息不过期不违规不会被删除</li>
                            <li>有置顶功能，可获得本站推荐</li>
                            <li>标题可变色，突出显示</li>
							<li>可上传图片，并有贴图权限</li>
                            <li>发布信息时内容框为HTM编辑器</li>
                            <li>已发布信息可随时修改删除</li>
							<li>信息可按名查阅，无限显示</li>
							<li>可开网店名片,信息无限显示</li>
                            <li>每次发布信息不用填联络信息</li></td>
                        </tr>
                      </table>
                    </span></td>
                  </tr>
                  <tr valign="middle">
                    <td height="30" ><div align="center"><a href="<%=addinfo%>?<%=CT_ID%>&ttop=2"><img src="img/PostSelect1.gif"   border="0" /></a> </div></td>
                    <td><div align="center"> </div></td>
                    <td><div align="center"><a href="<%=reg%>?<%=CT_ID%>&ttop=10"><img src="img/PostSelect2.gif" border="0"   /></a> </div></td>
                  <td><div align="center"> </div></td>
                    <td><div align="center"><a href="help.asp?id=1&<%=CT_ID%>&ttop=10"><img src="img/PostSelect3.gif" border="0"   /></a> </div></td>
                  </tr>
                </table>                  </td>

              </tr>
            </table> 
            <br />
<center>
<!--JS调用方式广告代码开始-->
  <script src="<%=adjs_path("ads/js","syzb730",c1&"_"&c2&"_"&c3)%>"></script>
  <!--广告代码结束-->
      
            </div></td>
        </tr>
    </table>	</td>
  </tr>
</table>
 
   <table width="980" cellpadding="0" cellspacing="0" align="center" >
     <tr>
       <td><!--#include file=footer.asp--></td>
     </tr>
</table>
</body>
</html>
