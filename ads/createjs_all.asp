<!--#include file="../conn.asp"-->
<!--#include file="../config.asp"-->
<!--#include file=cookies.asp-->
<!--#include file="../SqlIn.Asp"-->
<!--#include file=Html2Js.asp-->
<%
'=====================================
'天天供求信息网站管理系统
'天天网络科技工作室出品
'客服中心http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>
<%call checkmanage("13")%>
<%dim objStream
	Dim city,adkg,ADID,HostURL,ScrCode,JsCode,AD_ADHits,ADjs
	Dim AD_Type,AD_Src,AD_Width,AD_Height,AD_Link,AD_Alt,AD_Views,ADViews,AD_Hits,ADHits,AD_Stop_Views,AD_Stop_Hits,AD_Stop_Date,city_oneid,city_twoid,city_threeid,InstallDir
If strInstallDir<>"" Then
InstallDir=StrReverse(Mid(StrReverse(strInstallDir),2))+"\/"
Else
InstallDir=""
End if
Response.Expires = 0
city=Request("city")
city_oneid=TRIM(Request("city_oneid"))
city_twoid=TRIM(Request("city_twoid"))
city_threeid=TRIM(Request("city_threeid"))
If city_oneid="" Then city_oneid=0
If city_twoid="" Then city_twoid=0
If city_threeid="" Then city_threeid=0
		HostURL="http://"&Request.Servervariables("server_name")&":"&Request.ServerVariables("SERVER_PORT")&replace(Request.Servervariables("url"),"/createjs_all.asp","")
Call OpenConn
	Set rs=server.createobject("adodb.recordset")
		sql = "Select ADID,adkg,ADType,ADSrc,ADLink,ADWidth,ADHeight,ADAlt,ADViews,ADHits,ADStopViews,ADStopHits,ADjs,ADStopDate,city_oneid,city_twoid,city_threeid from [DNJY_ad]"
If city="ok" Then
sql=sql&" where  and city_oneid="&city_oneid&" and city_twoid="&city_twoid&" and city_threeid="&city_threeid&""
End if
		rs.open sql,conn,1,1
	If rs.bof or rs.eof Then'=================a===================
    response.write "没有广告记录，无法生成广告代码！"
    rs.close
	Set rs=nothing
	conn.Close
	Set conn=nothing
	response.end
    Else
  do while not rs.eof  
	    ADID=rs("ADID")
		AD_Type=rs("ADType")
		AD_Src=rs("ADSrc")
		AD_Link=rs("ADLink")
		AD_Width=rs("ADWidth")
		AD_Height=rs("ADHeight")
		AD_Alt=rs("ADAlt")
		AD_Views=rs("ADViews")
		AD_ADHits=rs("ADHits")
		AD_Stop_Views=rs("ADStopViews")
		AD_Stop_Hits=rs("ADStopHits")
		AD_Stop_Date=rs("ADStopDate")
		ADjs=rs("ADjs")	
		adkg=rs("adkg")
		city_oneid=rs("city_oneid")
		city_twoid=rs("city_twoid")
		city_threeid=rs("city_threeid")

	If ( AD_Stop_Views <> 0 and AD_Views > AD_Stop_Views) Then AD_Type = 0
	If ( AD_Stop_Hits <> 0 and AD_Hits > AD_Stop_Hits) Then AD_Type = 0
	If ( NOW() > AD_Stop_Date) Then AD_Type = 0
	If adkg=0 Then AD_Type = 0

	If InStr(1,AD_Src,".swf",1)>0 Then
		ScrCode="<EMBED src="""& AD_Src &""" quality=high WIDTH="""& AD_Width &""" HEIGHT="""& AD_Height &""" TYPE=""application/x-shockwave-flash"" PLUGINSPAGE=""http://www.macromedia.com/go/getflashplayer""></EMBED>"
	Else
		ScrCode="<a href="""& HostURL &"/openad.asp?adid="& ADID &"&c="&city_oneid&","&city_twoid&","&city_threeid&""" target=_blank><img src="""& AD_Src &""" border=0 width="""& AD_Width &""" height="""& AD_Height &""" alt="""& AD_Alt &""" align=top></a>"
	End If
    JsCode=""
	Select Case AD_Type
	Case 0
		JsCode= "document.write('<!-- 过期或暂停广告 ID"&ADID&" -->')"
	Case 1
		JsCode= "document.write('"& ScrCode &"');"
	Case 2
		JsCode= "ns4=(document.layers)?true:false;" & vbCrLf & _
				"ie4=(document.all)?true:false;" & vbCrLf & _
				"if(ns4){document.write('<layer id=DGbanner2 width="& AD_Width &" height="& AD_Height &" onmouseover=stopme(""DGbanner2"") onmouseout=movechip(""DGbanner2"")>"& ScrCode &"</layer>');}" & vbCrLf & _
				"else{document.write('<div id=DGbanner2 style=""position:absolute; width:"& AD_Width &"px; height:"& AD_Height &"px; z-index:9; filter: Alpha(Opacity=90)"" onmouseover=stopme(""DGbanner2"") onmouseout=movechip(""DGbanner2"")>"& ScrCode &"<DIV onclick=""hidead();"" style=""FONT-SIZE: 9pt; CURSOR: hand"" align=right>关闭×</DIV></div>');}" & vbCrLf & _
				"document.write('<script language=javascript src="& HostURL &"/images/ad_float_fullscreen.js></script>');"
	Case 3
		JsCode= "if (navigator.appName == 'Netscape')" & vbCrLf & _
				"{document.write('<layer id=DGbanner3 top=150 width="& AD_Width &" height="& AD_Height &">"& ScrCode &"</layer>');}" & vbCrLf & _
				"else{document.write('<div id=DGbanner3 style=""position: absolute;width:"& AD_Width &";top:150;visibility: visible;z-index: 1"">"& ScrCode &"</div>');}" & vbCrLf & _
				"document.write('<script language=javascript src="& HostURL &"/images/ad_float_upanddown.js></script>');"
	Case 4
		JsCode= "if (navigator.appName == 'Netscape')" & vbCrLf & _
				"{document.write('<layer id=DGbanner10 top=150 width="& AD_Width &" height="& AD_Height &">"& ScrCode &"</layer>');}" & vbCrLf & _
				"else{document.write('<div id=DGbanner10 style=""position: absolute;width:"& AD_Width &";top:150;visibility: visible;z-index: 1"">"& ScrCode &"</div>');}" & vbCrLf & _
				"document.write('<script language=javascript src="& HostURL &"/images/ad_float_upanddown_L.js></script>');"
	Case 5
		JsCode= "ns4=(document.layers)?true:false;" & vbCrLf & _
				"if(ns4){document.write('<layer id=DGbanner4Cont onLoad=""moveToAbsolute(layer1.pageX-160,layer1.pageY);clip.height="& AD_Height &";clip.width="& AD_Width &"; visibility=show;""><layer id=DGbanner4News position:absolute; top:0; left:0>"& ScrCode &"</layer></layer>');}" & vbCrLf & _
				"else{document.write('<div id=DGbanner4 style=""position:absolute;top:0; left:0;""><div id=DGbanner4Cont style=""position:absolute; width:"& AD_Width &"; height:"& AD_Height &";clip:rect(0,"& AD_Width &","& AD_Height &",0)""><div id=DGbanner4News style=""position:absolute;top:0; left:0; right:820"">"& ScrCode &"</div></div></div>');} " & vbCrLf & _
				"document.write('<script language=javascript src="& HostURL &"/images/ad_fullscreen.js></script>');"
	Case 6
		JsCode= "window.showModalDialog('"& AD_Src &"','','dialogWidth:"& AD_Width &"px;dialogHeight:"& AD_Height &"px;scroll:no;status:no;help:no');"
	Case 7
		JsCode= "document.write('<script language=javascript src="& HostURL &"/images/ad_dialog.js></script>'); " & vbCrLf & _
				"document.write('<div style=""position:absolute;left:300px;top:150px;width:"& AD_Width &"; height:"& AD_Height &";z-index:1;solid;filter:alpha(opacity=90)"" id=DGbanner5 onmousedown=""down1(this)"" onmousemove=""move()"" onmouseup=""down=false""><table cellpadding=0 border=0 cellspacing=1 width="& AD_Width &" height="& AD_Height &" bgcolor=#000000><tr><td height=18 bgcolor=#5A8ACE align=right style=""cursor:move;""><a href=# style=""font-size: 9pt; color: #eeeeee; text-decoration: none"" onClick=clase(""DGbanner5"") >关闭>>><img border=""0"" src="""&HostURL&"/images/close_o.gif""></a>&nbsp;</td></tr><tr><td bgcolor=f4f4f4 >&nbsp;<a href="""& HostURL &"/openad.asp?adid="& ADID &""" target=_blank><img src="& AD_Src &"></a></td></tr></table></div>');"
	Case 8
		JsCode= "window.open('"& AD_Link &"','_blank');"
	Case 9
		JsCode= "window.open('"& AD_Link &"','DGBANNER7','width="& AD_Width &",height="& AD_Height &",scrollbars=1')"
	Case 10
		JsCode= "function closeAd(){" & vbCrLf & _
						"huashuolayer2.style.visibility='hidden';" & vbCrLf & _
						"huashuolayer3.style.visibility='hidden';}" & vbCrLf & _
						"function winload()" & vbCrLf & _
						"{" & vbCrLf & _
						"huashuolayer2.style.top=20;" & vbCrLf & _
						"huashuolayer2.style.left=5;" & vbCrLf & _
						"huashuolayer3.style.top=20;" & vbCrLf & _
						"huashuolayer3.style.right=5;" & vbCrLf & _
						"}" & vbCrLf & _
						"if(document.body.offsetWidth>800){" & vbCrLf & _
						"{" & vbCrLf & _
						"document.write('<div id=huashuolayer2 style=""position: absolute;visibility:visible;z-index:1""><table width=100  border=0 cellspacing=0 cellpadding=0><tr><td height=10 align=right bgcolor=666666><a href=javascript:closeAd()><img src=ads/images/close.gif width=12 height=10 border=0></a></td></tr><tr><td>" & ScrCode & "</td></tr></table></div>'" & vbCrLf & _
						"+'<div id=huashuolayer3 style=""position: absolute;visibility:visible;z-index:1""><table width=100  border=0 cellspacing=0 cellpadding=0><tr><td height=10 align=right bgcolor=666666><a href=javascript:closeAd()><img src=ads/images/close.gif width=12 height=10 border=0></a></td></tr><tr><td>" & ScrCode & "</td></tr></table></div>');" & vbCrLf & _
						"}" & vbCrLf & _
						"winload()" & vbCrLf & _
						"}"
						Case 11						
JsCode= "//符合web标准且可关闭的多幅对联广告" & vbCrLf & _
"lastScrollY = 0;" & vbCrLf & _
"function heartBeat(){" & vbCrLf & _
"var diffY;" & vbCrLf & _
"if (document.documentElement && document.documentElement.scrollTop)" & vbCrLf & _
"diffY = document.documentElement.scrollTop;" & vbCrLf & _
"else if (document.body)" & vbCrLf & _
"diffY = document.body.scrollTop" & vbCrLf & _
"else" & vbCrLf & _
"{/*Netscape stuff*/}" & vbCrLf & _
"//alert(diffY);" & vbCrLf & _
"percent=.1*(diffY-lastScrollY);" & vbCrLf & _
"if(percent>0)percent=Math.ceil(percent);" & vbCrLf & _
"else percent=Math.floor(percent);" & vbCrLf & _
"document.getElementById(""leftDiv"").style.top = parseInt(document.getElementById(""leftDiv"").style.top)+percent+""px"";" & vbCrLf & _
"document.getElementById(""rightDiv"").style.top = parseInt(document.getElementById(""rightDiv"").style.top)+percent+""px"";" & vbCrLf & _
"lastScrollY=lastScrollY+percent;" & vbCrLf & _
"//alert(lastScrollY);" & vbCrLf & _
"}" & vbCrLf & _
"//下面这段删除后，对联将不跟随屏幕而移动。" & vbCrLf & _
"window.setInterval(""heartBeat()"",1);" & vbCrLf & _
"//-->" & vbCrLf & _
"//关闭按钮" & vbCrLf & _
"function close_left1(){" & vbCrLf & _
"    left1.style.visibility='hidden';" & vbCrLf & _
"}" & vbCrLf & _
"function close_left2(){" & vbCrLf & _
"    left2.style.visibility='hidden';" & vbCrLf & _
"}" & vbCrLf & _
"function close_right1(){" & vbCrLf & _
"    right1.style.visibility='hidden';" & vbCrLf & _
"}" & vbCrLf & _
"function close_right2(){" & vbCrLf & _
"    right2.style.visibility='hidden';" & vbCrLf & _
"}" & vbCrLf & _
"//显示样式" & vbCrLf & _
"document.writeln(""<style type=\""text\/css\"">"");" & vbCrLf & _
"document.writeln(""#leftDiv,#rightDiv{width:100px;height:100px;position:absolute;}"");" & vbCrLf & _
"document.writeln("".itemFloat{width:100px;height:auto;line-height:5px}"");" & vbCrLf & _
"document.writeln(""<\/style>"");" & vbCrLf & _
"//以下为主要内容" & vbCrLf & _
"document.writeln(""<div id=\""leftDiv\"" style=\""top:10px;left:5px\"">"");" & vbCrLf & _
"//------左侧各块开始" & vbCrLf & _
"//---L1" & vbCrLf & _
"document.writeln(""<div id=\""left1\"" class=\""itemFloat\"">"");" & vbCrLf & _
"document.writeln('"&ScrCode&"');" & vbCrLf & _
"document.writeln(""<br><a href=\""javascript:close_left1();\"" title=\""关闭上面的广告\""><img src="&HostURL&"/images/close1.gif width=100 height=35 border=0><\/a><br><br><br><br>"");" & vbCrLf & _
"document.writeln(""<\/div>"");" & vbCrLf & _
"//------左侧各块结束" & vbCrLf & _
"document.writeln(""<\/div>"");" & vbCrLf & _
"document.writeln(""<div id=\""rightDiv\"" style=\""top:10px;right:5px\"">"");" & vbCrLf & _
"//------右侧各块结束" & vbCrLf & _
"//---R1" & vbCrLf & _
"document.writeln(""<div id=\""right1\"" class=\""itemFloat\"">"");" & vbCrLf & _
"document.writeln('"&ScrCode&"');" & vbCrLf & _
"document.writeln(""<br><a href=\""javascript:close_right1();\"" title=\""关闭上面的广告\""><img src="&HostURL&"/images/close1.gif width=100 height=35 border=0><\/a><br><br><br><br>"");" & vbCrLf & _
"document.writeln(""<\/div>"");" & vbCrLf & _
"//------右侧各块结束" & vbCrLf & _
"document.writeln(""<\/div>"");" & vbCrLf & _
 ""
 Case 12
JsCode=JsCode&"document.write(""<marquee onmouseover=\""this.stop()\"" onmouseout=\""this.start()\"" scrollAmount=\"""&AD_Width&"\""><a href=\"""& HostURL &"/openad.asp?adid="& ADID &"\"" target=_blank><b><font color=#FF0000>"&AD_Alt&"</font></b><\/a></mrquee>"")"

 Case 13
JsCode=JsCode&"document.write(""<a href=\"""& HostURL &"/openad.asp?adid="& ADID &"\"" target=_blank><b><font color=#FF0000>"&AD_Alt&"</font></b><\/a>"")"

 Case 14
JsCode=JsCode&"document.write('"&Html2Js(ADjs)&"')"

 Case 15
JsCode="document.write('<certer><object classid=clsid:D27CDB6E-AE6D-11cf-96B8-444553540000 id=scriptmain name=scriptmain codebase=http:\/\/download.macromedia.com\/pub\/shockwave\/cabs\/')" & vbCrLf & _
"document.write('flash\/swflash.cab#version=6,0,29,0 width="&AD_Width&" height="&AD_Height&">')" & vbCrLf & _
"document.write('<param name=movie value=\/"&InstallDir&"ads\/pic\/"&Replace(AD_Src,"/"&strInstallDir&"ads/pic/","")&">')" & vbCrLf & _
"document.write('<param name=quality value=high>')" & vbCrLf & _
"document.write('<param name=scale value=noscale>')" & vbCrLf & _
"document.write('<param name=menu value=false>')" & vbCrLf & _
"document.write('<param name=wmode value=transparent>')" & vbCrLf & _
"document.write('<embed src=\/"&InstallDir&"ads\/pic\/"&Replace(AD_Src,"/"&strInstallDir&"ads/pic/","")&" width="&AD_Width&" height="&AD_Height&" quality=high pluginspage=http:\/\/www.macromedia.com\/go\/getflashplayer type=application\/x-shockwave-flash salign=T name=scriptmain menu=false wmode=transparent><\/embed>')" & vbCrLf & _
"document.write('<\/object></certer>')" & vbCrLf & _
""
Case 16
JsCode="document.write('<object id=""player"" height="""&AD_Height&""" width="""&AD_Width&""" classid=""CLSID:6BF52A52-394A-11d3-B153-00C04F79FAA6"">')" & vbCrLf & _
"document.write('<param NAME=""AutoStart"" VALUE=""-1"">')" & vbCrLf & _
"document.write('<param NAME=""Balance"" VALUE=""0"">')" & vbCrLf & _
"document.write('<param name=""enabled"" value=""-1"">')" & vbCrLf & _
"document.write('<param NAME=""EnableContextMenu"" VALUE=""-1"">')" & vbCrLf & _
"document.write('<param NAME=""url"" value="""& AD_Src &""">')" & vbCrLf & _
"document.write('<param NAME=""PlayCount"" VALUE=""1"">')" & vbCrLf & _
"document.write('<param name=""rate"" value=""1"">')" & vbCrLf & _
"document.write('<param name=""currentPosition"" value=""0"">')" & vbCrLf & _
"document.write('<param name=""currentMarker"" value=""0"">')" & vbCrLf & _
"document.write('<param name=""defaultFrame"" value="""">')" & vbCrLf & _
"document.write('<param name=""invokeURLs"" value=""0"">')" & vbCrLf & _
"document.write('<param name=""baseURL"" value="""">')" & vbCrLf & _
"document.write('<param name=""stretchToFit"" value=""0"">')" & vbCrLf & _
"document.write('<param name=""volume"" value=""50"">')" & vbCrLf & _
"document.write('<param name=""mute"" value=""0"">')" & vbCrLf & _
"document.write('<param name=""uiMode"" value=""Full"">')" & vbCrLf & _
"document.write('<param name=""windowlessVideo"" value=""0"">')" & vbCrLf & _
"document.write('<param name=""fullScreen"" value=""0"">')" & vbCrLf & _
"document.write('<param name=""enableErrorDialogs"" value=""-1"">')" & vbCrLf & _
"document.write('<param name=""SAMIStyle"" value>')" & vbCrLf & _
"document.write('<param name=""SAMILang"" value>')" & vbCrLf & _
"document.write('<param name=""SAMIFilename"" value>')" & vbCrLf & _
"document.write('</object>')" & vbCrLf & _

""
	End Select


dim JsFileName
	If JsCode<>"" Then'==========n============
	On Error Resume Next
		JsFileName = Server.MapPath("JS\"&ADID&"_"&city_oneid&"_"&city_twoid&"_"&city_threeid&".js")
		If AD_Type=11 then
		JsCode = "" & vbCrLf & JsCode & vbCrLf &"" 
		else
		JsCode = "<!--" & vbCrLf & JsCode & vbCrLf &"//-->"  
		end if
	set objStream = Server.CreateObject("ADODB.Stream")
	With objStream
		.Open
		.Type = 2
		.Charset = "GB2312"
		.WriteText JsCode
		.SaveToFile JsFileName, 2
		.Close
	End With
	Set objStream = Nothing


	End If'============n===============
rs.movenext
Loop
Response.Write ("<script>alert('已成功生成所有广告JS代码 !');history.back();</script>")
    rs.close
	Set rs=nothing
	conn.Close
	Set conn=nothing
End If'====================a==================%>