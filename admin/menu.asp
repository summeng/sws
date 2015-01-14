<!--#include file=cookies.asp-->
<!--#include file="../config.asp"-->
<%
'=====================================
'天天供求信息网站管理系统
'天天网络科技工作室出品
'客服中心http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />

<title>管理菜单</title>
<script type="text/javascript" src="inc/js_left.js"></script>

<base target="mainFrame" />
<link rel="stylesheet" href="inc/style_left.css" type="text/css" />
</head>
<body>
<div id="body">

<div class="d">控制面板</div>
<div class="e"><a href="right.asp" TARGET=right>管理首页</a> ｜ <a HREF="exit.asp" TARGET="_parent">安全退出</a></div>

<div class="a" id="AA0" onclick="TAB(0)" >供求信息管理</div>
<div class="c" id="BB0" style="display:block;">
<p><A href="infolist.asp" TARGET=right>信息管理</A> ｜ <A href="info_add.asp" TARGET=right>发布管理</A></p>
<p><A href="xxhflist.asp" TARGET=right>回复信息</A> ｜ <A href="Bad_info_list.asp" TARGET=right>举报信息</A></p>
<p><a href="siterss.asp" TARGET=right>总站RSS</a> ｜ <a href="Substationrss.asp" TARGET=right>分站RSS</a></p>
<p><a href="Substationrss_Batches.asp" TARGET=right>批量RSS</a> ｜ <a href="info_html.asp" TARGET=right>html管理</a></p>
<p><a href="siteinfo.asp" TARGET=right>运程调用</a> ｜ <a href="sitenew.asp" TARGET=right>JS调用</a></p>
</div>

<div class="b" id="AA1" onclick="TAB(1)">会员资料管理</div>
<div class="c" id="BB1" style="display:none;">
<p><A href="users_List.asp" TARGET=right>会员管理</A> ｜ <A href="add_user.asp" TARGET=right>添加会员</A></p>
<p><A href="admin_order.asp" TARGET=right>订单管理</A> ｜ <A href=" admin_Amount.asp" TARGET=right>财务日志</A></p>
<p><A href="xnp.asp" TARGET=right>赠<%=rmb_mc%></A> ｜ <a href="sitehy.asp" TARGET=right>店企调用</a></p>
<p><A href="UserNews_ModSelClass.asp" TARGET=right>会员资讯</A> ｜ <A href="usernews_ClassManage.asp" TARGET=right>栏目管理</A></p>
<p><A href="sms_loglist.asp?isse=0" TARGET=right><font color="#4144C1">验证记录</font></A> <%If mailqr=1 then%>｜ <A href="user_Temporary.asp" TARGET=right>临时会员</A><%End if%></p>
</div>

<div class="b" id="AA2" onclick="TAB(2)">互动信息管理</div>
<div class="c" id="BB2" style="display:none;">
<p><A href="book.asp" TARGET=right>留言管理</A> ｜ </p>
<!--p><a href="#" ONCLICK="window.open('e_mail.asp','Sample','toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=yes,resizable=no,copyhistory=no,width=730,height=700,left=300,top=20')" target=_self>发送杂志</font></A> ｜ <a href="#" ONCLICK="window.open('sendmail_user.asp','Sample','toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=yes,resizable=no,copyhistory=no,width=730,height=700,left=300,top=20')" target=_self>杂志管理</font></A></p-->
<p><a href="#" ONCLICK="window.open('AdminSendEmail.asp','Sample','toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=yes,resizable=no,copyhistory=no,width=730,height=700,left=300,top=20')" target=_self>邮件群发</font></A> ｜ <A href="mail_loglist.asp?isse=0" TARGET=right><font color="#4144C1">邮件日志</font></A></p>
<p><A href="Message1.asp" TARGET=right><font color="#1DA2DB">接收短信</font></A> ｜ <A href="Message2.asp" TARGET=right><font color="#1DA2DB">撰写短信</font></A></p>
<p><A href="Message3.asp" TARGET=right><font color="#1DA2DB">短信管理</font></A> ｜ <A href="Message4.asp" TARGET=right><font color="#1DA2DB">删除短信</font></A></p>
</div>

<div class="b" id="AA3" onclick="TAB(3)">新闻资讯管理</div>
<div class="c" id="BB3" style="display:none;">
<p><A href="admin_article_add.asp"  TARGET=right>添加新闻资讯</A></p>
<p><A href="admin_article_list.asp"  TARGET=right>管理新闻资讯</A></p>
<p><A href="admin_article_comment.asp"  TARGET=right>文章评论管理</A></p>
<p><A href="admin_article_Class1.asp"  TARGET=right>文章栏目管理</A></p>
</div>

<div class="b" id="AA4" onclick="TAB(4)">电子杂志管理</div>
<div class="c" id="BB4" style="display:none;">
<p><A href="../magazine/magazine_issue.asp"  TARGET=right>杂志期号管理</A></p>
<p><A href="../magazine/magazine_edition.asp"  TARGET=right>杂志版面管理</A></p>
</div>

<div class="b" id="AA5" onclick="TAB(5)">信息采集管理</div>
<div class="c" id="BB5" style="display:none;">
<p><A href="collect_management.asp"  TARGET=right>项目管理</A></p>
<p><A href="collect_start.asp"  TARGET=right>数据采集</A></p>
<p><A href="collect_Settings.asp"  TARGET=right>过滤设置</A></p>
<p><A href="collect_record.asp"  TARGET=right>历史记录</A></p>
<p><A href="collect_laoy.asp?Action=import"  TARGET=right>导入规则</A></p>
<p><A href="collect_laoy.asp"  TARGET=right>导出规则</A></p>
</div>

<div class="b" id="AA6" onclick="TAB(6)">新闻采集管理</div>
<div class="c" id="BB6" style="display:none;">
<p><A href="cj_management.asp"  TARGET=right>项目管理</A></p>
<p><A href="cj_start.asp"  TARGET=right>数据采集</A></p>
<p><A href="cj_Settings.asp"  TARGET=right>过滤设置</A></p>
<p><A href="cj_record.asp"  TARGET=right>历史记录</A></p>
<p><A href="cj_laoy.asp?Action=import"  TARGET=right>导入规则</A></p>
<p><A href="cj_laoy.asp"  TARGET=right>导出规则</A></p>
</div>

<div class="b" id="AA7" onclick="TAB(7)">网站系统管理</div>
<div class="c" id="BB7" style="display:none;">
<p><A href="admin_parameters.asp?id=1" TARGET=right>基本参数</A> ｜ <A href="add_administrator.asp" TARGET=right>管理员权限</A></p>
<p><a href="pay1.asp" TARGET=right>支付平台</a> ｜ <a href="pay2.asp" TARGET=right>银行帐号</a></p>
<p><a href="api.asp" TARGET=right>API设置</a> ｜ <A href="Mail_system.asp?id=1" TARGET=right>邮件系统</A> </p>
<p><a href="type_Level1.asp"  TARGET=right>信息分类</a> ｜ <a href="city_Level1.asp"  TARGET=right>地区分类</a></p>
<p><a href="store_Level1.asp"  TARGET=right>店企分类</a> ｜ <a href="admin_help.asp"  TARGET=right>帮助设置</A></p>
<p><a href="admin_announce.asp"  TARGET=right>本站公告</A> ｜ <a href="#" ONCLICK="window.open('notice.asp','Sample','toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=yes,resizable=no,copyhistory=no,width=730,height=500,left=300,top=20')" target=_self>通知分站</a></p>
</div>

<div class="b" id="AA8" onclick="TAB(8)">网站模板管理</div>
<div class="c" id="BB8" style="display:none;">
<p><a href="admin_Template_Management.asp" TARGET=right>模板设置</a> ｜ <a href="Admin_Template_Main.asp?Action=Main" TARGET=right>模板回收站</a></p>
<p><a href="Admin_Template.asp?Action=Import" TARGET=right>模板导入</a> ｜ <a href="Admin_Template.asp?Action=Export" TARGET=right>模板导出</a></p>
<!--<p><a href="sitemap.asp" TARGET=right>创建 sitemap</a></p>-->
</div>

<div class="b" id="AA9" onclick="TAB(9)">附属功能管理</div>
<div class="c" id="BB9" style="display:none;">
<p><a href="bmcs.asp" TARGET=right>便民信息</a> ｜ <a href="admin_114.asp" TARGET=right>都市114</a></p>
<p><a href="admin_link.asp" TARGET=right>友情连接</a> ｜ </p>
</div>

<div class="b" id="AA10" onclick="TAB(10)">标准广告管理</div>
<div class="c" id="BB10" style="display:none;">
<p><a href="../ads/add.asp" TARGET=right>添加广告</a> ｜ <a href="../ads/list.asp?all=ok" TARGET=right>广告列表</a></p>
<p><a href="../ads/list.asp?action=stop&all=ok" TARGET=right>过期广告</a> ｜ <a href="../ad.asp" TARGET=right>默认广告</a></p>
<p><a href="../ads/code.asp" TARGET=right>调用示例</a> ｜ <a href="../ads/Leading_in_out.asp" TARGET=right>导入导出</a></p>
</div>

<div class="b" id="AA11" onclick="TAB(11)">特殊广告管理</div>
<div class="c" id="BB11" style="display:none;">
<p><a href="adv_slide.asp" TARGET=right><font color="#ff0000">幻灯广告A</font></a> ｜ <a href="adv.asp" TARGET=right><font color="#ff0000">幻灯广告B</font></a></p>
<p> <a href="adfp.asp" TARGET=right>翻牌广告</a>&nbsp;&nbsp;&nbsp;｜ <a href="admin_flv.asp" TARGET=right>FLV视频</a></p>
</div>

<div class="b" id="AA12" onclick="TAB(12)">安防过滤清理</div>
<div class="c" id="BB12" style="display:none;">
<p><A href="admin_ip.asp" TARGET=right>字符IP限制｜<A href="admin_check.asp" TARGET=right>问答验证</A></p>
<!--p><a href="admin_backupdata.asp" TARGET=right>备份数据</a>｜<a href="admin_Compressdata.asp" TARGET=right>压缩数据</a></p-->
<p><a href="delcache.asp" TARGET=right>清理缓存</a>｜<a href="../SpaceSize.asp" TARGET=right>空间占用</a></p>
<p><a href="delpic.asp" TARGET=right>清理无用上传图片和文件</a></p>
<p><a href="admin_server.asp" TARGET=right>服务器探针</a>｜<a href="admin_data.asp" TARGET=right>数据库管理</a></p>
<!--<p><a href="fileUpdate.asp" TARGET=right>随机更改入口文件名</a></p>-->
<%If sys=True then%><p><a href="admin_server.asp" TARGET=right>服务器探针</a>｜<a href="sqladmin.asp" TARGET=right>注入记录</a></p><%End if%>
</div>

<div class="f">系统版权信息</div>
<div class="c">
<p><a href="http://www.ip126.com/" target="_blank">客服中心</a> ｜ <a href="http://www.ip126.com/bbs/index.asp?boardid=49" target="_blank">客服论坛</a></p>
<p><a href="http://www.ip126.com" target="_blank">天天网络科技工作室</a></p>
</div>

</div>
</body>
</html>
