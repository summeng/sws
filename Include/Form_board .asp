<%
Function M_top(Moldboard)

M_top=Moldboard
End Function

Function M_Index(Moldboard,M)
Dim TextFile,i,Text
M=split(M,"|")
TextFile="<!--"&"#include file="
Moldboard=replace(Moldboard,"{$搜索条}","<%=f_search("&m(0)&","&m(1)&")%"&">")
Moldboard=replace(Moldboard,"{$最新信息}","<%=f_info(1,0,0,0,c1,c2,c3,"&m(2)&","&m(3)&","&m(4)&","&m(5)&","&m(6)&","&m(7)&","&m(8)&","&m(9)&","&m(10)&","&m(11)&","&m(12)&","&m(13)&","&m(14)&","&m(15)&","&m(16)&")%"&">")
Moldboard=replace(Moldboard,"{$竞价信息}","<%=f_info(2,0,0,0,c1,c2,c3,"&m(17)&","&m(18)&","&m(19)&","&m(20)&","&m(21)&","&m(22)&","&m(23)&","&m(24)&","&m(25)&","&m(26)&","&m(27)&","&m(28)&","&m(29)&","&m(30)&","&m(31)&")%"&">")
Moldboard=replace(Moldboard,"{$推荐信息}","<%=f_info(3,0,0,0,c1,c2,c3,"&m(32)&","&m(33)&","&m(34)&","&m(35)&","&m(36)&","&m(37)&","&m(38)&","&m(39)&","&m(40)&","&m(41)&","&m(42)&","&m(43)&","&m(44)&","&m(45)&","&m(46)&")%"&">")
Moldboard=replace(Moldboard,"{$热门信息}","<%=f_info(4,0,0,0,c1,c2,c3,"&m(47)&","&m(48)&","&m(49)&","&m(50)&","&m(51)&","&m(52)&","&m(53)&","&m(54)&","&m(55)&","&m(56)&","&m(57)&","&m(58)&","&m(59)&","&m(60)&","&m(61)&")%"&">")
Moldboard=replace(Moldboard,"{$分类信息}","<%=xxfl(c1,c2,c3,"&m(62)&","&m(63)&","&m(64)&","&m(65)&","&m(66)&","&m(67)&","&m(68)&","&m(69)&","&m(70)&","&m(71)&","&m(72)&","&m(73)&")%"&">")
Moldboard=replace(Moldboard,"{$方框显示信息}","<%=fkss(c1,c2,c3,"&m(74)&","&m(75)&","&m(76)&","&m(77)&","&m(78)&","&m(79)&","&m(80)&","&m(81)&","&m(82)&","&m(83)&","&m(84)&","&m(85)&","&m(86)&","&m(87)&")%"&">")
Moldboard=replace(Moldboard,"{$图片信息推荐}","<%=f_tptj(c1,c2,c3,3,"&m(88)&","&m(89)&")%"&">")
Moldboard=replace(Moldboard,"{$最新信息图文}","<%=tw_info(1,0,0,0,c1,c2,c3,"&m(90)&","&m(91)&","&m(92)&","&m(93)&","&m(94)&","&m(95)&","&m(96)&","&m(97)&","&m(98)&","&m(99)&","&m(100)&","&m(101)&","&m(102)&","&m(103)&","&m(104)&","&m(105)&","&m(106)&","&m(107)&")%"&">")
Moldboard=replace(Moldboard,"{$推荐信息图文}","<%=tw_info(3,0,0,0,c1,c2,c3,"&m(108)&","&m(109)&","&m(110)&","&m(111)&","&m(112)&","&m(113)&","&m(114)&","&m(115)&","&m(116)&","&m(117)&","&m(118)&","&m(119)&","&m(120)&","&m(121)&","&m(122)&","&m(123)&","&m(124)&","&m(125)&")%"&">")
Moldboard=replace(Moldboard,"{$热门信息图文}","<%=tw_info(4,0,0,0,c1,c2,c3,"&m(126)&","&m(127)&","&m(128)&","&m(129)&","&m(130)&","&m(131)&","&m(132)&","&m(133)&","&m(134)&","&m(135)&","&m(136)&","&m(137)&","&m(138)&","&m(139)&","&m(140)&","&m(141)&","&m(142)&","&m(143)&")%"&">")
Moldboard=replace(Moldboard,"{$名店推荐}","<%=f_dmtj(c1,c2,c3,"&m(144)&","&m(145)&","&m(146)&","&m(147)&","&m(148)&","&m(149)&","&m(150)&","&m(151)&","&m(152)&","&m(153)&")%"&">")
Moldboard=replace(Moldboard,"{$店铺资源共享}","<%=vipnews("""","&m(154)&","&m(155)&","&m(156)&","&m(157)&","&m(158)&","&m(159)&","&m(160)&","&m(161)&","&m(162)&","&m(163)&","&m(164)&","&m(165)&","&m(166)&")%"&">")
Moldboard=replace(Moldboard,"{$企业文字名片}","<%=f_qymp(c1,c2,c3,"&m(167)&","&m(168)&","&m(169)&","&m(170)&","&m(171)&","&m(172)&","&m(173)&","&m(174)&","&m(175)&","""&mid(m(176),2)&""","&m(177)&","&m(178)&")%"&">")
Moldboard=replace(Moldboard,"{$企业图片名片}","<%=qymp(c1,c2,c3,"&m(179)&","&m(180)&","&m(181)&","&m(182)&")%"&">")

Moldboard=replace(Moldboard,"{$新闻资讯一}","<%=news(c1,c2,c3,"&m(183)&","&m(184)&","&m(185)&","&m(186)&","&m(187)&","&m(188)&","&m(189)&","&m(190)&","&m(191)&","&m(192)&","&m(193)&","&m(194)&","&m(195)&","&m(196)&","&m(197)&")%"&">")
Moldboard=replace(Moldboard,"{$新闻资讯二}","<%=news(c1,c2,c3,"&m(198)&","&m(199)&","&m(200)&","&m(201)&","&m(202)&","&m(203)&","&m(204)&","&m(205)&","&m(206)&","&m(207)&","&m(208)&","&m(209)&","&m(210)&","&m(211)&","&m(212)&")%"&">")
Moldboard=replace(Moldboard,"{$新闻资讯三}","<%=news(c1,c2,c3,"&m(213)&","&m(214)&","&m(215)&","&m(216)&","&m(217)&","&m(218)&","&m(219)&","&m(220)&","&m(221)&","&m(222)&","&m(223)&","&m(224)&","&m(225)&","&m(226)&","&m(227)&")%"&">")
Moldboard=replace(Moldboard,"{$新闻资讯四}","<%=news(c1,c2,c3,"&m(228)&","&m(229)&","&m(230)&","&m(231)&","&m(232)&","&m(233)&","&m(234)&","&m(235)&","&m(236)&","&m(237)&","&m(238)&","&m(239)&","&m(240)&","&m(241)&","&m(242)&")%"&">")
Moldboard=replace(Moldboard,"{$新闻资讯相册}","<%=news_photo(c1,c2,c3,"&m(243)&","&m(244)&","&m(245)&","&m(246)&","&m(247)&","&m(248)&","&m(249)&","&m(250)&","&m(251)&","&m(252)&","&m(253)&","&m(254)&","&m(255)&",request(""page""),"""","""")%"&">")
Moldboard=replace(Moldboard,"{$电子杂志}","<%=magazine(c1,c2,c3,"&m(256)&","&m(257)&","&m(258)&","&m(259)&","&m(260)&")%"&">")
Moldboard=replace(Moldboard,"{$都市114}","<%=ds114(c1,c2,c3,"&m(261)&","&m(262)&")%"&">")
Moldboard=replace(Moldboard,"{$便民服务}","<%=bmfw(c1,c2,c3,"&m(263)&","&m(264)&","&m(265)&","&m(266)&","&m(267)&","&m(268)&","&m(269)&","&m(270)&","&m(271)&","&m(272)&")%"&">")
Moldboard=replace(Moldboard,"{$友情链接}","<%=f_link(c1,c2,c3,"&m(273)&","&m(274)&","&m(275)&","&m(276)&","&m(277)&","&m(278)&","&m(279)&","&m(280)&")%"&">")
Moldboard=replace(Moldboard,"{$用户留言}","<%=book(c1,c2,c3,"&m(281)&","&m(282)&","&m(283)&","&m(284)&","&m(285)&","&m(286)&")%"&">")
Moldboard=replace(Moldboard,"{$本站公告}","<%=announce(c1,c2,c3,"&m(287)&","&m(288)&","&m(289)&","&m(290)&","&m(291)&","&m(292)&","&m(293)&","&m(294)&","&m(295)&")%"&">")

Text=TextFile&"""conn.asp""-->"&vbCrLf&TextFile&"""SqlIn.Asp""-->"&vbCrLf&TextFile&"""top.asp""-->"&vbCrLf&Moldboard&vbCrLf&TextFile&"""footer.asp"" -->"&vbCrLf&TextFile&"""qq_kefu.asp""-->"&vbCrLf&"<script language='javascript' src='/Resetting.js'></script>"
M_Index=Text
End Function



Function M_xxIndex(Moldboard,M)
Dim TextFile,i,Text
M=split(M,"|")
TextFile="<!--"&"#include file="
Moldboard=replace(Moldboard,"{$搜索条}","<%=f_search("&m(0)&","&m(1)&")%"&">")
Moldboard=replace(Moldboard,"{$最新信息}","<%=f_info(1,0,0,0,c1,c2,c3,"&m(2)&","&m(3)&","&m(4)&","&m(5)&","&m(11)&","&m(6)&","&m(7)&","&m(8)&","&m(9)&","&m(10)&","&m(12)&","&m(13)&","&m(14)&","&m(15)&","&m(16)&")%"&">")
Moldboard=replace(Moldboard,"{$竞价信息}","<%=f_info(2,0,0,0,c1,c2,c3,"&m(17)&","&m(18)&","&m(19)&","&m(20)&","&m(26)&","&m(21)&","&m(22)&","&m(23)&","&m(24)&","&m(25)&","&m(27)&","&m(28)&","&m(29)&","&m(30)&","&m(31)&")%"&">")
Moldboard=replace(Moldboard,"{$推荐信息}","<%=f_info(3,0,0,0,c1,c2,c3,"&m(32)&","&m(33)&","&m(34)&","&m(35)&","&m(41)&","&m(36)&","&m(37)&","&m(38)&","&m(39)&","&m(40)&","&m(42)&","&m(43)&","&m(44)&","&m(45)&","&m(46)&")%"&">")
Moldboard=replace(Moldboard,"{$热门信息}","<%=f_info(4,0,0,0,c1,c2,c3,"&m(47)&","&m(48)&","&m(49)&","&m(50)&","&m(56)&","&m(51)&","&m(52)&","&m(53)&","&m(54)&","&m(55)&","&m(57)&","&m(58)&","&m(59)&","&m(60)&","&m(61)&")%"&">")
Moldboard=replace(Moldboard,"{$推荐图片信息}","<%=f_tptj(c1,c2,c3,3,"&m(62)&","&m(63)&")%"&">")
Moldboard=replace(Moldboard,"{$最新图片信息}","<%=f_tptj(c1,c2,c3,1,"&m(64)&","&m(65)&")%"&">")
Moldboard=replace(Moldboard,"{$竞价图片信息}","<%=f_tptj(c1,c2,c3,2,"&m(66)&","&m(67)&")%"&">")
Moldboard=replace(Moldboard,"{$热门图片信息}","<%=f_tptj(c1,c2,c3,4,"&m(68)&","&m(69)&")%"&">")
Moldboard=replace(Moldboard,"{$置顶图片信息}","<%=f_tptj(c1,c2,c3,5,"&m(70)&","&m(71)&")%"&">")
Moldboard=replace(Moldboard,"{$方框显示信息}","<%=fkss(c1,c2,c3,"&m(72)&","&m(73)&","&m(74)&","&m(75)&","&m(76)&","&m(77)&","&m(78)&","&m(79)&","&m(80)&","&m(81)&","&m(82)&","&m(83)&","&m(84)&","&m(85)&")%"&">")
Moldboard=replace(Moldboard,"{$信息分类}","<%=s_more("&m(86)&","&m(87)&","&m(88)&","&m(89)&","&m(90)&","&m(91)&","&m(92)&","&m(93)&","&m(94)&","&m(95)&","&m(96)&","&m(97)&")%"&">")
Moldboard=replace(Moldboard,"{$信息列表}","<%=xxfl(c1,c2,c3,"&m(98)&","&m(99)&","&m(100)&","&m(101)&","&m(102)&","&m(103)&","&m(104)&","&m(105)&","&m(106)&","&m(107)&","&m(108)&","&m(109)&")%"&">")


Text=TextFile&"""conn.asp""-->"&vbCrLf&TextFile&"""SqlIn.Asp""-->"&vbCrLf&TextFile&"""top.asp""-->"&vbCrLf&Moldboard&vbCrLf&TextFile&"""footer.asp"" -->"&vbCrLf&TextFile&"""qq_kefu.asp""-->"&vbCrLf&"<script language='javascript' src='/Resetting.js'></script>"
M_xxIndex=Text
End Function

Function M_Footer(Moldboard,M)
M=split(M,"|")
Moldboard=replace(Moldboard,"{$关于本站}","<a href=""help.asp?id=0&<%=C_ID%"&">"">关于本站</a>")
Moldboard=replace(Moldboard,"{$会员帮助}","<a href=""help.asp?id=1&<%=C_ID%"&">"">会员帮助</a>")
Moldboard=replace(Moldboard,"{$服务条款}","<a href=""help.asp?id=2&<%=C_ID%"&">"">服务条款</a>")
Moldboard=replace(Moldboard,"{$免责声明}","<a href=""help.asp?id=3&<%=C_ID%"&">"">免责声明</a>")
Moldboard=replace(Moldboard,"{$广告服务}","<a href=""help.asp?id=4&<%=C_ID%"&">"">广告服务</a>")
Moldboard=replace(Moldboard,"{$联系我们}","<a href=""help.asp?id=5&<%=C_ID%"&">"">联系我们</a>")
Moldboard=replace(Moldboard,"{$客户留言}","<a href=""Message_board.asp"">客户留言</a>")
Moldboard=replace(Moldboard,"{$程序版本}","<a href=""http://www.ip126.com/"" target=_blank>Version <%=SystemVersion%"&"></a>")
Moldboard=replace(Moldboard,"{$公司名称}","<%=cityCompany%"&">")
Moldboard=replace(Moldboard,"{$网站地址}","<%=cityweb%"&">")
Moldboard=replace(Moldboard,"{$网站名称}","<a href=""http://<%=cityweb%"&">""><%=title%"&"></a>")
Moldboard=replace(Moldboard,"{$联系电话}","电话:<%=tel%"&">")
Moldboard=replace(Moldboard,"{$联系手机}","手机:<%=Cellular_phone%"&">")
Moldboard=replace(Moldboard,"{$传真}","传真:<%=fax%"&">")
Moldboard=replace(Moldboard,"{$办公地址}","办公地址：<%=serve%"&">")
Moldboard=replace(Moldboard,"{$邮政编码}","邮政编码：<%=yzbm%"&">")
Moldboard=replace(Moldboard,"{$网站ICP}","<a href=""http://www.miibeian.gov.cn"" target=_blank><%=icp%"&"></a>")
Moldboard=replace(Moldboard,"{$统计代码}",""&conn.Execute("Select tjdm From DNJY_config")(0)&"")
M_Footer=Moldboard
End Function

Function M_show(Moldboard,M)
Dim TextFile,Text,i,shows9
M=split(M,"|")
Moldboard=replace(Moldboard,"{$最新信息}","<SCRIPT src=""Include/info_money.asp?id=<%=Request(""id"")%"&">&xxsx=1&<%=C_ID%"&">&n="&m(1)&"&l="&m(2)&"&s="&m(3)&"&gdxx="&m(4)&"&gd="&m(5)&"&dj="&m(6)&"&zs="&m(7)&"&zt="&m(8)&"&hg="&m(9)&"&d="&m(10)&"""></SCRIPT>")
Moldboard=replace(Moldboard,"{$竞价信息}","<SCRIPT src=""Include/info_money.asp?id=<%=Request(""id"")%"&">&xxsx=2&<%=C_ID%"&">&n="&m(11)&"&l="&m(12)&"&s="&m(13)&"&gdxx="&m(14)&"&gd="&m(15)&"&dj="&m(16)&"&zs="&m(17)&"&zt="&m(18)&"&hg="&m(19)&"&d="&m(20)&"""></SCRIPT>")
Moldboard=replace(Moldboard,"{$推荐信息}","<SCRIPT src=""Include/info_money.asp?id=<%=Request(""id"")%"&">&xxsx=3&<%=C_ID%"&">&n="&m(21)&"&l="&m(22)&"&s="&m(23)&"&gdxx="&m(24)&"&gd="&m(25)&"&dj="&m(26)&"&zs="&m(27)&"&zt="&m(28)&"&hg="&m(29)&"&d="&m(30)&"""></SCRIPT>")
Moldboard=replace(Moldboard,"{$热门信息}","<SCRIPT src=""Include/info_money.asp?id=<%=Request(""id"")%"&">&xxsx=4&<%=C_ID%"&">&n="&m(31)&"&l="&m(32)&"&s="&m(33)&"&gdxx="&m(34)&"&gd="&m(35)&"&dj="&m(36)&"&zs="&m(37)&"&zt="&m(38)&"&hg="&m(39)&"&d="&m(40)&"""></SCRIPT>")
TextFile="<!--"&"#include file="
Text=TextFile&"""conn.asp""-->"&vbCrLf&TextFile&"""SqlIn.Asp""-->"&vbCrLf&TextFile&"""top.asp""-->"&vbCrLf
Text=Text&"<%dim zuo1,zuo2,zuo3,info1,info2,info3,info4,info5,xxtp,biaoti,leixin,biaotixs,fbsj,scjiage,jiage,sdmba,usernameid,namea,dianhua,mobile,userqq,email,xxmpname,dizhi,youbian,wcolor,ncolor,memo,hfcs,tail,cksj"&vbcrlf
Text=Text&"XinxiData%"&">"&vbcrlf
Text=Text&Moldboard&vbCrLf&TextFile&"""footer.asp"" -->"
Text=replace(Text,"{$网站名称}","<%=title%"&">")
Text=replace(Text,"{$网站地址}","<%=web%"&">")
Text=replace(Text,"{$信息标题}","<%=biaoti%"&">")
Text=replace(Text,"{$信息标题显示}","<%=biaotixs%"&">")
Text=replace(Text,"{$网站关键词}","<%=keywords%"&">")
Text=replace(Text,"{$网站简介}","<%=description%"&">")
Text=replace(Text,"{$信息编号}","<%=Id%"&">")
Text=replace(Text,"{$地区ID}","city_oneid=<%=c1%"&">&city_twoid=<%=c2%"&">&city_threeid=<%=c3%"&">")
Text=replace(Text,"{$发布时间}","<%=fbsj%"&">")
Text=replace(Text,"{$有效时间}","<SCRIPT src=""user/xinfodt.asp?action=dqsj&id=<%=id%"&">""></SCRIPT>")
Text=replace(Text,"{$浏览次数}","<SCRIPT src=""user/xinfodt.asp?action=llcs&id=<%=id%"&">""></SCRIPT>")
Text=replace(Text,"{$评论次数}","<SCRIPT src=""user/xinfodt.asp?action=hfcs&id=<%=id%"&">""></SCRIPT>")
Text=replace(Text,"{$不良记录}","<SCRIPT src=""user/xinfodt.asp?action=Bad&id=<%=id%"&">""></SCRIPT>")
Text=replace(Text,"{$不良举报}","<a href=""user/Bad_info_list.asp?infoid=<%=cstr(id)%"&">&biaoti=<%=biaoti%"&">"" target=blank>查</a>")
Text=replace(Text,"{$信息类别}","<%=belongtype%"&">")
Text=replace(Text,"{$交易类型}","[<%=leixin%"&">]")
Text=replace(Text,"{$交易状态}","<SCRIPT src=""user/xinfodt.asp?action=zt&id=<%=id%"&">""></SCRIPT>")
Text=replace(Text,"{$市场价格}","<%=scjiage%"&">")
Text=replace(Text,"{$交易价格}","<%=jiage%"&">")
Text=replace(Text,"{$交易地区}","<%=belongcity%"&">")
Text=replace(Text,"{$IP地址}","<%=userip%"&">")
Text=replace(Text,"{$商家店企}","<%=sdmba%"&">")
Text=replace(Text,"{$会员ID号}","<%=username%"&">")
Text=replace(Text,"{$会员ID显示}","<%=usernameid%"&">")
Text=replace(Text,"{$信息图片}","<%=xxtp%"&">")
Text=replace(Text,"{$联系人}","<%=name%"&">")
Text=replace(Text,"{$联系人显示}","<%=namea%"&">")
Text=replace(Text,"{$固定电话}","<%=dianhua%"&">")
Text=replace(Text,"{$移动电话}","<%=mobile%"&">")
Text=replace(Text,"{$QQ号码}","<%=userqq%"&">")
Text=replace(Text,"{$电子信箱}","<%=email%"&">")
Text=replace(Text,"{$公司名称}","<%=xxmpname%"&">")
Text=replace(Text,"{$联系地址}","<%=dizhi%"&">")
Text=replace(Text,"{$邮政编码}","<%=youbian%"&">")
Text=replace(Text,"{$外观颜色}","<%=wcolor%"&">")
Text=replace(Text,"{$内饰颜色}","<%=ncolor%"&">")
Text=replace(Text,"{$详细说明}","<%=memo%"&">")
Text=replace(Text,"{$本站声明}",""&m(0)&"")
Text=replace(Text,"{$转载声明}","【<%=title%"&">】http://<%=web%"&">")
Text=replace(Text,"{$评论表单开始}","<form name=""form"" id=""form"" onSubmit=""return form_onsubmit()"" action=""user/reply.asp?id=<%=id%"&">&Readinfo=<%=Readinfo%"&">&xxusername=<%=username%"&">&biaoti=<%=biaoti%"&">&email=<%=email%"&">&dick=chk&city_oneid=<%=c1%"&">&city_twoid=<%=c2%"&">&city_threeid=<%=c3%"&">"" method=""POST"">")
'Text=replace(Text,"{$直接评论}","<input type=""radio"" name=""r1"" value=""0"" checked>")
'Text=replace(Text,"{$邮箱评论}","<input type=""radio"" value=""1"" name=""r1"">")
Text=replace(Text,"{$评论用户名表单}","<input type=""text"" name=""username"" size=""16"" maxlength=""50"" value=""<%=request.cookies(""dick"")(""username"")%"&">"">")
Text=replace(Text,"{$匿名发布}","<input type=""checkbox"" name=""nm"" value=""on"">")
Text=replace(Text,"{$评论邮件标题表单}","<input type=""text"" name=""mbt"" size=""39"" maxlength=""50"">")
'Text=replace(Text,"{$评论邮箱表单}","<input type=""text"" name=""mdz"" size=""30"" maxlength=""40"" value=""<%=request.cookies(""dick"")(""email"")%"&">"">")
'Text=replace(Text,"{$评论者邮箱表单}","<input type=""hidden"" name=""et"" size=""15"" value=""<%=email%"&">"">")
Text=replace(Text,"{$评论内容文本框}","<textarea rows=""4"" name=""neirong"" cols=""50"" maxlength=""200""></textarea>")
Text=replace(Text,"{$验证码}","<tr><td height=""29""><p style=""line-height: 100%; margin-left: 20px"">算式验证(答案见下)：<input type=""text"" name=""validate_codee"" size=""4"" onkeyup=""value=value.replace(/[^\d|^\.]/g,'')""   onbeforepaste=""clipboardData.setData('text',clipboardData.getData('text').replace(/[^\d|^\.]/g,''))""></td></tr>")
Text=replace(Text,"{$评论提交按钮}","<input type=""submit"" value=""提交评论"" name=""hf"">")
Text=replace(Text,"{$评论表单结束}","</Form>")
Text=replace(Text,"{$评论显示区}","<SCRIPT src=""user/rq_message.asp?id=<%=id%"&">&hfcs=<%=hfcs%"&">""></SCRIPT>")
Text=replace(Text,"{$评论数量}","<SCRIPT src=""user/rq_messagecs.asp?id=<%=id%"&">""></SCRIPT>")
If shows8=1 then
Text=replace(Text,"{$广告zuo1}","<script language=javascript src=""<%=adjs_path(""ads/js"",""zuo1"",c1&""_""&c2&""_""&c3)%"&">""></script>")
Text=replace(Text,"{$广告zuo2}","<script language=javascript src=""<%=adjs_path(""ads/js"",""zuo2"",c1&""_""&c2&""_""&c3)%"&">""></script>")
Text=replace(Text,"{$广告zuo3}","<script language=javascript src=""<%=adjs_path(""ads/js"",""zuo3"",c1&""_""&c2&""_""&c3)%"&">""></script>")
Text=replace(Text,"{$广告info1}","<script language=javascript src=""<%=adjs_path(""ads/js"",""info1"",c1&""_""&c2&""_""&c3)%"&">""></script>")
Text=replace(Text,"{$广告info2}","<script language=javascript src=""<%=adjs_path(""ads/js"",""info2"",c1&""_""&c2&""_""&c3)%"&">""></script>")
Text=replace(Text,"{$广告info3}","<script language=javascript src=""<%=adjs_path(""ads/js"",""info3"",c1&""_""&c2&""_""&c3)%"&">""></script>")
Text=replace(Text,"{$广告info4}","<script language=javascript src=""<%=adjs_path(""ads/js"",""info4"",c1&""_""&c2&""_""&c3)%"&">""></script>")
Text=replace(Text,"{$广告info5}","<script language=javascript src=""<%=adjs_path(""ads/js"",""info5"",c1&""_""&c2&""_""&c3)%"&">""></script>")
Text=replace(Text,"{$广告tail}","<script language=javascript src=""<%=adjs_path(""ads/js"",""tail"",c1&""_""&c2&""_""&c3)%"&">""></script>")
Text=replace(Text,"{$广告tc}","<script language=javascript src=""<%=adjs_path(""ads/js"",""tc"",c1&""_""&c2&""_""&c3)%"&">""></script>")
Else
Text=replace(Text,"{$广告zuo1}","<script src=""<%=adjs_path(""ads/js"",""zuo1"",""0_0_0"")%"&">""></script>")
Text=replace(Text,"{$广告zuo2}","<script src=""<%=adjs_path(""ads/js"",""zuo2"",""0_0_0"")%"&">""></script>")
Text=replace(Text,"{$广告zuo3}","<script src=""<%=adjs_path(""ads/js"",""zuo3"",""0_0_0"")%"&">""></script>")
Text=replace(Text,"{$广告info1}","<script src=""<%=adjs_path(""ads/js"",""info1"",""0_0_0"")%"&">""></script>")
Text=replace(Text,"{$广告info2}","<script src=""<%=adjs_path(""ads/js"",""info2"",""0_0_0"")%"&">""></script>")
Text=replace(Text,"{$广告info3}","<script src=""<%=adjs_path(""ads/js"",""info3"",""0_0_0"")%"&">""></script>")
Text=replace(Text,"{$广告info4}","<script src=""<%=adjs_path(""ads/js"",""info4"",""0_0_0"")%"&">""></script>")
Text=replace(Text,"{$广告info5}","<script src=""<%=adjs_path(""ads/js"",""info5"",""0_0_0"")%"&">""></script>")
Text=replace(Text,"{$广告tail}","<script src=""<%=adjs_path(""ads/js"",""tail"",""0_0_0"")%"&">""></script>")
Text=replace(Text,"{$广告tc}","<script src=""<%=adjs_path(""ads/js"",""tc"",""0_0_0"")%"&">""></script>")
End if
M_show=Text
End Function

Function M_list(Moldboard,M)
Dim TextFile,Text,i
M=split(M,"|")
Moldboard=replace(Moldboard,"{$搜索条}","<%=f_search("&m(0)&","&m(1)&")%"&">")
Moldboard=replace(Moldboard,"{$信息分类列表}","<%=F_vassal_c_t(c1,c2,c3,t1,t2,t3,"&m(2)&","&m(3)&","&m(4)&","&m(5)&","&m(6)&",1)%"&">")
Moldboard=replace(Moldboard,"{$地区分类列表}","<%=F_vassal_c_t(c1,c2,c3,t1,t2,t3,"&m(7)&","&m(8)&","&m(9)&","&m(10)&","&m(11)&",0)%"&">")
Moldboard=replace(Moldboard,"{$最新信息}","<%=f_info(1,0,0,0,c1,c2,c3,"&m(12)&","&m(13)&","&m(14)&","&m(15)&","&m(16)&","&m(17)&","&m(18)&","&m(19)&","&m(20)&","&m(21)&","&m(22)&","&m(23)&","&m(24)&","&m(25)&","&m(26)&")%"&">")
Moldboard=replace(Moldboard,"{$竞价信息}","<%=f_info(2,0,0,0,c1,c2,c3,"&m(27)&","&m(28)&","&m(29)&","&m(30)&","&m(31)&","&m(32)&","&m(33)&","&m(34)&","&m(35)&","&m(36)&","&m(37)&","&m(38)&","&m(39)&","&m(40)&","&m(41)&")%"&">")
Moldboard=replace(Moldboard,"{$推荐信息}","<%=f_info(3,0,0,0,c1,c2,c3,"&m(42)&","&m(43)&","&m(44)&","&m(45)&","&m(46)&","&m(47)&","&m(48)&","&m(49)&","&m(50)&","&m(51)&","&m(52)&","&m(53)&","&m(54)&","&m(55)&","&m(56)&")%"&">")
Moldboard=replace(Moldboard,"{$热门信息}","<%=f_info(4,0,0,0,c1,c2,c3,"&m(57)&","&m(58)&","&m(59)&","&m(60)&","&m(61)&","&m(62)&","&m(63)&","&m(64)&","&m(65)&","&m(66)&","&m(67)&","&m(68)&","&m(69)&","&m(70)&","&m(71)&")%"&">")
Moldboard=replace(Moldboard,"{$栏目信息}","<%=f_typeinfo(xxsx,c1,c2,c3,t1,t2,t3,strint(request(""page"")),"&m(72)&",tw,s,"&m(75)&","&m(76)&","&m(77)&","&m(78)&")%"&">")
Moldboard=replace(Moldboard,"{$信息图文列表}","<%=tw_info(xxsx,t1,t2,t3,c1,c2,c3,"&m(79)&","&m(80)&","&m(81)&","&m(82)&","&m(83)&","&m(84)&","&m(85)&","&m(86)&","&m(87)&","&m(88)&","&m(89)&","&m(90)&","&m(91)&","&m(92)&","&m(93)&","&m(94)&","&m(95)&","&m(96)&")%"&">")
TextFile="<!--"&"#include file="
Text=TextFile&"""conn.asp""-->"&vbCrLf&TextFile&"""SqlIn.Asp""-->"&vbCrLf
Text=Text&TextFile&"""top.asp""-->"&vbCrLf
Text=Text&Moldboard&vbCrLf&TextFile&"""footer.asp"" -->"
M_list=Text
End Function

Function News000(Moldboard,M)
Dim Text,TextFile,i
M=split(M,"|")

Moldboard=replace(Moldboard,"{$固顶新闻}","<SCRIPT src=""Include/topnews.asp?t=0&<%=C_ID%"&">&n="&m(1)&"&s="&m(3)&"&l="&m(2)&"&zs="&m(4)&"&zt="&m(6)&"&hg="&m(7)&"&d="&m(5)&"""></SCRIPT>")
TextFile="<!--"&"#include file="
Text=TextFile&"""top.asp""-->"&vbCrLf
Text=Text&"<SCRIPT src=""Showjs.asp?id=<%=id%"&">""></SCRIPT>"&vbcrlf
Text=Text&"<%dim title,belongcity,text,pic,username,tel,QQ,web,email,adddate,stopdate,ip"&vbcrlf
Text=Text&"newsData%"&">"&vbcrlf
Text=Text&Moldboard&vbCrLf&TextFile&"""footer.asp"" -->"
Text=replace(Text,"{$标题}","<%=title%"&">")
Text=replace(Text,"{$所属区域}","<%=belongcity%"&">")
Text=replace(Text,"{$新闻类别}","<%=belongtype%"&">")
Text=replace(Text,"{$新闻编号}","<%=Id%"&">")
Text=replace(Text,"{$人气指数}","<script>document.write(hits)</SCRIPT>")
Text=replace(Text,"{$新闻内容}","<%=text%"&">")
Text=replace(Text,"{$新闻图片}","<%=Pic%"&">")
Text=replace(Text,"{$发布时间}","<%=adddate%"&">")
Text=replace(Text,"{$新闻栏目}","<%=News_Type%"&">")
Text=replace(Text,"{$自定义广告代码}",""&m(0)&"")
News=Text
End Function



Function News_More000(Moldboard,M)
Dim Text,TextFile,i
M=split(M,"|")
for i=1 to 50
Moldboard=replace(Moldboard,"{$广告"&i&"}","<%=f_advertisement(piclink("&i&"),picname("&i&"),picpass("&i&"),picw("&i&"),pich("&i&"))%"&">")
next
Moldboard=replace(Moldboard,"{$新闻列表}","<%=F_news_more(city_oneid,city_twoid,city_threeid,"&m(1)&",strint(Request(""pageid"")),"&m(3)&",Strint(Request(""ID"")))%"&">")

Moldboard=replace(Moldboard,"{$新闻栏目}","<%=News_Type%"&">")

Moldboard=replace(Moldboard,"{$固顶新闻}","<%=f_topnews(city_oneid,city_twoid,city_threeid,"&m(4)&","&m(6)&","&m(5)&","&m(7)&","&m(9)&","&m(10)&","&m(8)&")%"&">")

TextFile="<!--"&"#include file="
Text=TextFile&"""top.asp""-->"&vbCrLf
Text=Text&Moldboard&vbCrLf&TextFile&"""footer.asp"" -->"
News_More=Text
End Function

Function M_css(Moldboard)
M_css=Moldboard
End Function
%>