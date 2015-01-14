<!--#include file="../conn.asp"-->
<!--#include file="../config.asp"-->
<!--#include file="../Include/head.asp"-->
<!--#include file="../Include/Function.asp"-->
<!--#include file=Html2Js.asp-->
<%
'=====================================
'天天供求信息网站管理系统
'天天网络科技工作室出品
'客服中心http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>
<%Call OpenConn
	Dim ADID,HostURL,AD_Type,AD_Src,AD_Width,AD_Height,AD_Link,AD_Alt,AD_Views,ADViews,AD_Hits,ADHits,AD_Stop_Views,AD_Stop_Hits,AD_Stop_Date,ScrCode,AD_ADHits,ADjs
		Response.Expires = 0

		ADID=Request.QueryString("ADID")
		HostURL="http://"&Request.Servervariables("server_name")&":"&Request.ServerVariables("SERVER_PORT")&replace(Request.Servervariables("url"),"/ad.asp","")
		host = "http://"&Request.ServerVariables("HTTP_HOST")
If shows8=0 Then ttccads=""
	Set rs=server.createobject("adodb.recordset")
		sql = "Select ADType,ADSrc,ADLink,ADWidth,ADHeight,ADAlt,ADViews,ADHits,ADStopViews,ADStopHits,ADStopDate,ADjs,city_oneid,city_twoid,city_threeid from [DNJY_ad] where ADID='" & ADID & "' and adkg=1"&ttccads&""
		rs.open sql,conn,1,3
	If Not (rs.bof or rs.eof) Then
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
		city_oneid=rs("city_oneid")
		city_twoid=rs("city_twoid")
		city_threeid=rs("city_threeid")
		rs("ADViews")= AD_Views + 1
		rs.Update
	Else
		response.write "document.write('<!--无此地区广告。公益广告：<br>地球是我们共同的家！--> ');"
	End If
		rs.Close
	Set rs=nothing
		conn.Close
	Set conn=nothing

	If ( AD_Stop_Views <> 0 and AD_Views > AD_Stop_Views) Then AD_Type = 0
	If ( AD_Stop_Hits <> 0 and AD_Hits > AD_Stop_Hits) Then AD_Type = 0
	If ( NOW() > AD_Stop_Date) Then AD_Type = 0

	If InStr(1,AD_Src,".swf",1)>0 Then
		ScrCode="<EMBED src="""& AD_Src &""" quality=high WIDTH="""& AD_Width &""" HEIGHT="""& AD_Height &""" TYPE=""application/x-shockwave-flash"" PLUGINSPAGE=""http://www.macromedia.com/go/getflashplayer""></EMBED>"
	Else
		ScrCode="<a href="""&host&"/"&strInstallDir&"ads/openad.asp?adid="& ADID &"&c="&city_oneid&","&city_twoid&","&city_threeid&""" target=_blank><img src="""& AD_Src &""" border=0 width="""& AD_Width &""" height="""& AD_Height &""" alt="""& AD_Alt &""" align=top></a>"
	End If

	Select Case AD_Type
	Case 1
		response.write ("document.write('"& ScrCode &"')")
	Case 2
		response.write ("ns4=(document.layers)?true:false;" & _
						"ie4=(document.all)?true:false;" & _
						"if(ns4){document.write('<layer id=DGbanner2 width="& AD_Width &" height="& AD_Height &" onmouseover=stopme(""DGbanner2"") onmouseout=movechip(""DGbanner2"")>"& ScrCode &"</layer>');}" & _
						"else{document.write('<div id=DGbanner2 style=""position:absolute; width:"& AD_Width &"px; height:"& AD_Height &"px; z-index:9; filter: Alpha(Opacity=90)"" onmouseover=stopme(""DGbanner2"") onmouseout=movechip(""DGbanner2"")>"& ScrCode &"</div>');}" & _
						"document.write('<script language=javascript src="& HostURL &"/images/ad_float_fullscreen.js></script>');")
	Case 3
		response.write ("if (navigator.appName == 'Netscape')" & _
						"{document.write('<layer id=DGbanner3 top=150 width="& AD_Width &" height="& AD_Height &">"& ScrCode &"</layer>');}" & _
						"else{document.write('<div id=DGbanner3 style=""position: absolute;width:"& AD_Width &";top:150;visibility: visible;z-index: 1"">"& ScrCode &"</div>');}" & _
						"document.write('<script language=javascript src="& HostURL &"/images/ad_float_upanddown.js></script>');")
	Case 4
		response.write ("if (navigator.appName == 'Netscape')" & _
						"{document.write('<layer id=DGbanner10 top=150 width="& AD_Width &" height="& AD_Height &">"& ScrCode &"</layer>');}" & _
						"else{document.write('<div id=DGbanner10 style=""position: absolute;width:"& AD_Width &";top:150;visibility: visible;z-index: 1"">"& ScrCode &"</div>');}" & _
						"document.write('<script language=javascript src="& HostURL &"/images/ad_float_upanddown_L.js></script>');")
	Case 5
		response.write ("ns4=(document.layers)?true:false;" & _
						"if(ns4){document.write('<layer id=DGbanner4Cont onLoad=""moveToAbsolute(layer1.pageX-160,layer1.pageY);clip.height="& AD_Height &";clip.width="& AD_Width &"; visibility=show;""><layer id=DGbanner4News position:absolute; top:0; left:0>"& ScrCode &"</layer></layer>');}" & _
						"else{document.write('<div id=DGbanner4 style=""position:absolute;top:0; left:0;""><div id=DGbanner4Cont style=""position:absolute; width:"& AD_Width &"; height:"& AD_Height &";clip:rect(0,"& AD_Width &","& AD_Height &",0)""><div id=DGbanner4News style=""position:absolute;top:0; left:0; right:820"">"& ScrCode &"</div></div></div>');} " & _
						"document.write('<script language=javascript src="& HostURL &"/images/ad_fullscreen.js></script>');")
	Case 6
		response.write ("window.showModalDialog('"& AD_Src &"','','dialogWidth:"& AD_Width &"px;dialogHeight:"& AD_Height &"px;scroll:no;status:no;help:no');")
	Case 7
response.write ("document.write('<script language=javascript src="&host&"/"&strInstallDir&"ads/images/ad_dialog.js></script>')" & vbCrLf & _ 
"document.write('<div style=""position:absolute;left:300px;top:150px;width:195; height:60;z-index:1;solid;filter:alpha(opacity=90)"" id=DGbanner5 onmousedown=""down1(this)"" onmousemove=""move()"" onmouseup=""down=false""><table cellpadding=0 border=0 cellspacing=1 width=195 height=60 bgcolor=#000000><tr><td height=18 bgcolor=#5A8ACE align=right style=""cursor:move;""><a href=# style=""font-size: 9pt; color: #eeeeee; text-decoration: none"" onClick=clase(""DGbanner5"") >关闭>>><img border=""0"" src="""&host&"/"&strInstallDir&"ads/images/close_o.gif""></a>&nbsp;</td></tr><tr><td bgcolor=f4f4f4 >&nbsp;<a href="""&host&"/"&strInstallDir&"ads/openad.asp?adid=zuo1"" target=_blank><img src="&host&"/"&strInstallDir&"ads/pic/zuo1.gif></a></td></tr></table></div>')")

	Case 8
		response.write ("window.open('"& AD_Link &"','_blank');")
	Case 9
		response.write ("window.open('"& AD_Link &"','DGBANNER7','width="& AD_Width &",height="& AD_Height &",scrollbars=1')")
	Case 10
		response.write ("function closeAd(){" &_
						"huashuolayer2.style.visibility='hidden';" &_
						"huashuolayer3.style.visibility='hidden';}" &_
						"function winload()" & _
						"{" & _
						"huashuolayer2.style.top=20;" & _
						"huashuolayer2.style.left=5;" & _
						"huashuolayer3.style.top=20;" & _
						"huashuolayer3.style.right=5;" & _
						"}" & _
						"if(document.body.offsetWidth>800){" & _
						"{" & _
						"document.write('<div id=huashuolayer2 style=""position: absolute;visibility:visible;z-index:1""><table width=100  border=0 cellspacing=0 cellpadding=0><tr><td height=10 align=right bgcolor=666666><a href=javascript:closeAd()><img src=images/close.gif width=12 height=10 border=0></a></td></tr><tr><td>" & ScrCode & "</td></tr></table></div>'" & _
						"+'<div id=huashuolayer3 style=""position: absolute;visibility:visible;z-index:1""><table width=100  border=0 cellspacing=0 cellpadding=0><tr><td height=10 align=right bgcolor=666666><a href=javascript:closeAd()><img src=images/close.gif width=12 height=10 border=0></a></td></tr><tr><td>" & ScrCode & "</td></tr></table></div>');" & _
						"}" & _
						"winload()" & _
						"}")
						
Case 11						
response.write ("var delta=0.015;" & _
 "var collection;" & _
 "var closeB=false;" & _
 "function floaters() {" & _
  "this.items = [];" & _
  "this.addItem = function(id,x,y,content)" & _
      "{" & _
     "document.write('<DIV id='+id+' style=""Z-INDEX: 10; POSITION: absolute;  width:80px;height:60px;left:'+(typeof(x)=='string'?eval(x):x)+';top:'+(typeof(y)=='string'?eval(y):y)+'"">'+content+'</DIV>');" & _     
     "var newItem    = {};" & _
     "newItem.object   = document.getElementById(id);" & _
     "newItem.x    = x;" & _
     "newItem.y    = y; " & _
     "this.items[this.items.length]  = newItem;" & _
      "}" & _
  "this.play = function()" & _
      "{" & _
     "collection    = this.items" & _
     "setInterval('play()',15);" & _
      "}" & _
  "}" & _
  "function play()" & _
  "{" & _
   "if(screen.width<=600 || closeB)" & _
   "{" & _
    "for(var i=0;i<collection.length;i++)" & _
    "{" & _
     "collection[i].object.style.display = 'none';" & _
    "}" & _
    "return;" & _
   "}" & _
   "for(var i=0;i<collection.length;i++)" & _
   "{" & _
    "var followObj  = collection[i].object;" & _
    "var followObj_x  = (typeof(collection[i].x)=='string'?eval(collection[i].x):collection[i].x);" & _
    "var followObj_y  = (typeof(collection[i].y)=='string'?eval(collection[i].y):collection[i].y);" & _
    "if(followObj.offsetLeft!=(document.body.scrollLeft+followObj_x)) {" & _
     "var dx=(document.body.scrollLeft+followObj_x-followObj.offsetLeft)*delta;" & _
     "dx=(dx>0?1:-1)*Math.ceil(Math.abs(dx));" & _
     "followObj.style.left=followObj.offsetLeft+dx;" & _
     "}" & _
    "if(followObj.offsetTop!=(document.body.scrollTop+followObj_y)) {" & _
     "var dy=(document.body.scrollTop+followObj_y-followObj.offsetTop)*delta;" & _
     "dy=(dy>0?1:-1)*Math.ceil(Math.abs(dy));" & _
     "followObj.style.top=followObj.offsetTop+dy;" & _
     "}" & _
    "followObj.style.display = '';" & _
   "}" & _
  "} " & _
  "function closeBanner()" & _
  "{" & _
   "closeB=true;" & _
   "return;" & _
  "}" & _
 "var theFloaters  = new floaters();" & _
 "//" & _
 "theFloaters.addItem('followDiv1','document.body.clientWidth-100',0,'" & ScrCode & "<br><br><img src="&host&"/"&strInstallDir&"ads/images/close1.gif onClick=""closeBanner();"">');" & _
"theFloaters.addItem('followDiv2',0,0,'" & ScrCode & "<br><br><img src="&host&"/"&strInstallDir&"ads/images/close1.gif onClick=""closeBanner();"">');" & _
 "theFloaters.play();" & _
"}")

 Case 12
response.write "document.write(""<marquee onmouseover=\""this.stop()\"" onmouseout=\""this.start()\"" scrollAmount=\"""&AD_Width&"\""><a href=\"""& HostURL &"/openad.asp?adid="& ADID &"\"" target=_blank><b><font color=#FF0000>"&AD_Alt&"</font></b><\/a></mrquee>"")"

 Case 13
response.write "document.write(""<a href=\"""& HostURL &"/openad.asp?adid="& ADID &"\"" target=_blank><b><font color=#FF0000>"&AD_Alt&"</font></b><\/a>"")"

 Case 14
response.write ("document.write('"&Html2Js(ADjs)&"')")

 Case 15

response.write (" document.write('<object classid=clsid:D27CDB6E-AE6D-11cf-96B8-444553540000 id=scriptmain name=scriptmain codebase=http:\/\/download.macromedia.com\/pub\/shockwave\/cabs\/')" & vbCrLf & _
"document.write('flash\/swflash.cab#version=6,0,29,0 width="&AD_Width&" height="&AD_Height&">')" & vbCrLf & _
"document.write('<param name=movie value=\/"&strInstallDir&"ads\/pic\/"&Replace(AD_Src,"/"&strInstallDir&"ads/pic/","")&">')" & vbCrLf & _
"document.write('<param name=quality value=high>')" & vbCrLf & _
"document.write('<param name=scale value=noscale>')" & vbCrLf & _
"document.write('<param name=menu value=false>')" & vbCrLf & _
"document.write('<param name=wmode value=transparent>')" & vbCrLf & _
"document.write('<embed src=\/"&strInstallDir&"ads\/pic\/"&Replace(AD_Src,"/"&strInstallDir&"ads/pic/","")&"> width="&AD_Width&" height="&AD_Height&" quality=high pluginspage=http:\/\/www.macromedia.com\/go\/getflashplayer type=application\/x-shockwave-flash salign=T name=scriptmain menu=false wmode=transparent><\/embed>')" & vbCrLf & _
"document.write('<\/object>')")

End Select
%>