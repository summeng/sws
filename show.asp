<!--#include file="conn.asp"-->
<!--#include file="SqlIn.Asp"-->
<!--#include file="top.asp"-->
<%dim zuo1,zuo2,zuo3,info1,info2,info3,info4,info5,xxtp,biaoti,leixin,biaotixs,fbsj,scjiage,jiage,sdmba,usernameid,namea,dianhua,mobile,userqq,email,xxmpname,dizhi,youbian,wcolor,ncolor,memo,hfcs,tail,cksj
XinxiData%>
<!--
'=====================================
'���칩����Ϣ��վ����ϵͳ
'��������Ƽ������Ұ�Ȩ����
'�ٷ��ͷ����� http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================-->

<script language="JavaScript">
//�������۱���
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
		if (targetDiv.style.display!='block'){
			targetDiv.style.display="block";

		}else{
			targetDiv.style.display="none";
		}
	}
}
</script>
<script LANGUAGE="JavaScript">
<!--
function checkMobile(){
	var sMobile = document.mobileform.mobile.value
	if(!(/^((\(\d{3}\))|(\d{3}\-))?13[0-9]\d{8}|15[0-9]\d{8}|18[0-9]\d{8}/.test(sMobile))){
		alert("����������11λ�ֻ��Ż�����ȷ���ֻ���ǰ��λ");
		document.mobileform.mobile.focus();
		return false;
	}
	window.open('', 'mobilewindow', 'height=197,width=350,status=yes,toolbar=no,menubar=no,location=no')
}
//-->
</script>
<script type="text/javascript" src="Include/copy.js"></script>
<link href="skin/ltby/css.css" type="text/css" rel="stylesheet">
<style type="text/css">
<!--
img {width:expression(this.width > 760 ? "760px" : this.width);overflow:hidden;max-width: 760px}
-->
</style>
<body>  
    <script>
    document.body.oncopy=function(){
    event.returnValue=false;
    var text = clipboardData.getData("text");
    var s=document.selection.createRange().text;
    t = s+'\n�������ԡ�<%=title%>��http://<%=web%> ,ת����ע�����������Ĳο���'+location.href; clipboardData.setData("text", text);
    clipboardData.setData('Text',t);
    }
    </script>
<table width="980" height="101" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
<td width="200" height="255" align="center" valign="top">
<table width="200" border="0" cellspacing="0" cellpadding="0">
<tr><td  height="7"></td></tr></table>

<table width="200"  cellpadding="1" cellspacing="1" class="tablest">
        <tr>
          <td height="30" class="tablest2" valign="top" ><a href="#" >
            <div class="infoft10">������Ϣ</div>
          </a></td>
        </tr>
        <tr>
          <td class="tablest1"  valign="top"><SCRIPT src="Include/info_money.asp?id=<%=Request("id")%>&xxsx=2&<%=C_ID%>&n=0&l=1&s=8&gdxx=1&gd=0&dj=0&zs=20&zt=13&hg=11&d=0"></SCRIPT></td>
        </tr>
</table>
<table width="200" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td  height="7"></td>
        </tr>
</table>
<table width="200" border="1" cellpadding="0" cellspacing="0" class="tablest" align="center">
        <tr>
          <td  style="border-style:none; padding:1px;" valign="top"><script src="<%=adjs_path("ads/js","zuo1","0_0_0")%>"></script></td>
</table>
<table width="200" border="0" cellspacing="0" cellpadding="0">
<tr><td  height="7"></td></tr></table>

<table width="200"  cellpadding="1" cellspacing="1" class="tablest">
        <tr>
          <td height="30" class="tablest2" valign="top" ><a href="#" >
            <div class="infoft10">������Ϣ</div>
          </a></td>
        </tr>
        <tr>
          <td class="tablest1"  valign="top"><SCRIPT src="Include/info_money.asp?id=<%=Request("id")%>&xxsx=1&<%=C_ID%>&n=0&l=1&s=8&gdxx=1&gd=1&dj=0&zs=20&zt=13&hg=11&d=0"></SCRIPT></td>
        </tr>
</table>
<table width="200" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td  height="7"></td>
        </tr>
</table>
<table width="200" border="1" cellpadding="0" cellspacing="0" class="tablest" align="center">
        <tr>
          <td  style="border-style:none; padding:1px;" valign="top"><script src="<%=adjs_path("ads/js","zuo2","0_0_0")%>"></script></td>
</table>
<table width="200" border="0" cellspacing="0" cellpadding="0">
<tr><td  height="7"></td></tr></table>

<table width="200"  cellpadding="1" cellspacing="1" class="tablest">
        <tr>
          <td height="30" class="tablest2" valign="top" ><a href="#" >
            <div class="infoft10">�Ƽ���Ϣ</div>
          </a></td>
        </tr>
        <tr>
          <td class="tablest1"  valign="top"><SCRIPT src="Include/info_money.asp?id=<%=Request("id")%>&xxsx=3&<%=C_ID%>&n=0&l=1&s=8&gdxx=1&gd=0&dj=0&zs=20&zt=13&hg=11&d=0"></SCRIPT></td>
        </tr>
</table>
<table width="200" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td  height="7"></td>
        </tr>
</table>
<table width="200" border="1" cellpadding="0" cellspacing="0" class="tablest" align="center">
        <tr>
          <td  style="border-style:none; padding:1px;" valign="top"><script src="<%=adjs_path("ads/js","zuo3","0_0_0")%>"></script></td>
</tr></table>
<table width="200" border="0" cellspacing="0" cellpadding="0">
<tr><td  height="7"></td></tr></table>

<table width="200"  cellpadding="1" cellspacing="1" class="tablest">
        <tr>
          <td height="30" class="tablest2" valign="top" ><a href="#" >
            <div class="infoft10">������Ϣ</div>
          </a></td>
        </tr>
        <tr>
          <td class="tablest1"  valign="top"><SCRIPT src="Include/info_money.asp?id=<%=Request("id")%>&xxsx=4&<%=C_ID%>&n=0&l=1&s=8&gdxx=1&gd=0&dj=1&zs=18&zt=13&hg=11&d=0"></SCRIPT></td>
        </tr>
</table>

	 </td>

	 <td width="10">&nbsp;</td>
    <td width="770" valign="top">
			
			  <table width="770" border="0" align="center" cellpadding="0" cellspacing="0">
                <tr>
                  <td width="770" align="center"><script src="<%=adjs_path("ads/js","info1","0_0_0")%>"></script></td>
                </tr>
		</table>
<table width="770" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td  height="7"></td>
</tr>
</table>
   <table width="770" border="0" align="center" cellpadding="0" cellspacing="0" class="tablest">
                <tr>
                  <td height="68" >
				
              <table width="560" height="63" border="0" align="center" cellpadding="0" cellspacing="0">

                <tr> 
                  <td height="41"> <table width="550" height="18" border="0" align="center" cellpadding="0" cellspacing="0">
                      <tr> 
                        <td align="center" ><div style="height:28px;padding-left:10px;font-family:����;font-size:20px;line-height:28px;font-weight:normal;letter-spacing:-1px;"><%=biaotixs%></div>                   </td>
                      </tr>
                  </table>                  </td>
                </tr>
              </table>
              <table width="770" border="0" align="center" cellpadding="0" cellspacing="0">
                <tr> 
                  <td height="5" colspan="2"></td>
                </tr>
                <tr> 
                  <td align="center"><b>����ʱ�䣺</b><%=fbsj%> <SCRIPT src="user/xinfodt.asp?action=dqsj&id=<%=id%>"></SCRIPT>��<b>���/���ۣ�</b><SCRIPT src="user/xinfodt.asp?action=llcs&id=<%=id%>"></SCRIPT>��/<SCRIPT src="user/xinfodt.asp?action=hfcs&id=<%=id%>"></SCRIPT>��  ������¼��<SCRIPT src="user/xinfodt.asp?action=Bad&id=<%=id%>"></SCRIPT>�� <a href="user/Bad_info_list.asp?infoid=<%=cstr(id)%>&biaoti=<%=biaoti%>" target=blank>��</a> [<a title=ƾɾ���������ɾ������Ϣ href="#" ONCLICK="window.open('../user/user_delxx.asp?id=<%=id%>&key=del','Sample','toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=no,resizable=no,copyhistory=no,width=300,height=200,left=300,top=100')">ɾ��</a>]</td>
                  <td align="center">��</td>
                </tr>
                <tr> 
                  <td height=10 colspan="2"></td>
                </tr>
              </table>
              <table width="750" border="0" align="center" cellpadding="0" cellspacing="0">
                <tr> 
                  <td width="260" valign="top"> <table width="250" height="161" border="0" cellpadding="0" cellspacing="2" bgcolor="#E4E4E4">
                      <tr> 
                        <td height="250" align="center" bgcolor="#FFFFFF"><%=xxtp%></td>
                      </tr>
                    </table></td>
                  <td valign="top"><table width="100%" border="0" cellpadding="2" cellspacing="1" bgcolor="#E4E4E4">
                      <tr bgcolor="#FFFFFF"> 
                        <td width="50%">&nbsp;��Ϣ���<%=belongtype%></td>
                        <td>&nbsp;�� ϵ �ˣ�<%=namea%></td>
                      </tr>
                      <tr bgcolor="#FFFFFF"> 
                        <td>&nbsp;�������ͣ�[<%=leixin%>]</td>
                        <td>&nbsp;��ԱID�ţ�<%=usernameid%></td>
                      </tr>
                      <tr bgcolor="#FFFFFF"> 
                        <td>&nbsp;����״̬��<SCRIPT src="user/xinfodt.asp?action=zt&id=<%=id%>"></SCRIPT></td>
                        <td>&nbsp;�̶��绰��<%=dianhua%></td>
						
                      </tr>
                      <tr bgcolor="#FFFFFF"> 
                        <td>&nbsp;���׼۸�<%=jiage%></td>
                         <form action="http://www.ip138.com:8080/search.asp" method="get" onSubmit="return checkMobile();" target="mobilewindow" name="mobileform"><td>&nbsp;�ƶ��绰��<input type="hidden" name="action" value="mobile"><%=mobile%> <input type="hidden" name="mobile" value="<%=mobile%>"> <INPUT type="submit" value="����" name="Submit"><br><%=cksj%>
                        </td></form>
                      </tr>
                      <tr bgcolor="#FFFFFF"> 
                        <td>&nbsp;���׵�����<%=belongcity%></td>
                        <td>&nbsp;�����ʼ���<%=email%> <a href="#" ONCLICK="window.open('user/user_mail.asp?username=<%=username%>&email=<%=email%>&mailzt=<%=biaoti%>','Sample','toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=yes,resizable=no,copyhistory=no,width=650,height=300,left=300,top=20')">���Ÿ���</a></td>
                      </tr>
                      <tr bgcolor="#FFFFFF"> 
                        <td>&nbsp;�̼ҵ���<%=sdmba%></td>
<td>&nbsp;Q Q ���룺<%=userqq%><a target=blank href=tencent://message/?Uin=<%=userqq%>&Site=<%=web%>&Menu=yes><img border='0' SRC=http://wpa.qq.com/pa?p=1:<%=userqq%>:7 alt=''��������'></a></td>
                      </tr>
                      <tr bgcolor="#FFFFFF"> 
                        <form method="get" action="http://www.ip138.com/ips1388.asp" name="ipform" onsubmit="return checkIP();" target="_blank">
	<td>&nbsp;I P ��ַ��<input type="text" name="ip" size="15" value="<%=userip%>" style="border: 1px solid #FFFFFF"> <input type="submit" value="����"><input type="hidden" name="action" value="2"></td></td></form>
                        <td>&nbsp;�������룺<%=youbian%></td>
                      </tr>
<tr bgcolor="#FFFFFF"> 
<td height="32" colspan="2">&nbsp;��˾���ƣ�<%=xxmpname%></td>
</tr>
                      <tr bgcolor="#FFFFFF"> 
                        <td height="32" colspan="2">&nbsp;ͨѶ��ַ��<%=dizhi%></td>
                      </tr>
                    </table></td>
                </tr>
                <tr> 
                  <td colspan="2" height="7"></td>
                </tr>
              </table>
 			  <table width="750" border="0" align="center" cellpadding="0" cellspacing="0">
                <tr> 
                  <td valign="top" bgcolor="#FFFFFF" class="a14">
				  <table width="100%" border="0" cellpadding="2" cellspacing="1" bgcolor="#E4E4E4">
                      <tr bgcolor="#FFFFFF"> 
                        <td width="760" bgcolor="#FCFCFC"><div style="margin:2px; font-size:14px; letter-spacing:1px; line-height:140%"><%=memo%></div></td></tr></table></td>
                </tr>
                <tr> 
                  <td height="7"></td>
                </tr>
              </table>
<table width="760" border="1" cellpadding="0" cellspacing="0" align="center">
        <tr>
          <td  style="border-style:none; padding:1px;" valign="top"><script src="<%=adjs_path("ads/js","info2","0_0_0")%>"></script></td>
</tr></table>             
              <table width="750" border="0" align="center" cellPadding="5" cellSpacing="1" bgcolor="#E4E4E4">
                <tr>
                  <td width="396" bgcolor="#FFFFFF" class="hei18"><font color="#FF0000"><b>��վ������</b><br>    �����칩�����������Ķ����е���Ϣ���������ģ���վ���κε�����δ����Ϣ��ʵ�Ժ�ʵ�����Ѳ�����Ϣʱ��һ��Ҫ��ϸ�б���٣������ϵ���ƭ����վ�����κε����;��׵��⡣����ͬ�ǽ��ס�</font></td>
                  <td width="301" bgcolor="#FFFFFF" class="hei18"> <table border="0" cellpadding="0" style="border-style:solid; border-width:1px; padding:1px; border-collapse: collapse" bordercolor="#CCCCCC" width="360"  height="50">
<TR><TD><p><a href="javascript:(function(){window.open('http://v.t.qq.com/share/share.php?title='+encodeURIComponent(document.title)+'&url='+encodeURIComponent(location.href)+'&source=bookmark','_blank','width=610,height=350');})()" title="����QQ΢��">��Ѷ΢��</a> | <a title='����QQ��ǩ' href="javascript:window.open('http://shuqian.qq.com/post?from=3&title='+encodeURIComponent(document.title)+'&uri='+encodeURIComponent(document.location.href)+'&jumpback=2&noui=1','favit','width=930,height=470,left=50,top=50,toolbar=no,menubar=no,location=no,scrollbars=yes,status=yes,resizable=yes');void(0)" style="background:url('http://shuqian.qq.com/img/add.gif') no-repeat 0px 0px;height:23px;width:60px;padding:2px 2px 0px 20px;font-size:12px;">QQ��ǩ</a> | <a title='��������΢��' href="javascript:window.open('http://v.t.sina.com.cn/share/share.php?title='+encodeURIComponent(document.title.substring(0,76))+'&url='+encodeURIComponent(location.href)+'&rcontent=','_blank','scrollbars=no,width=600,height=450,left=75,top=20,status=no,resizable=yes'); void 0" style="color:#000000;text-decoration:none;font-size:12px;font-weight:normal">����΢��</a> | <A title="�ղص�POCO ��ժhttp://share.poco.cn" href="javascript:d=document;t=d.selection?(d.selection.type!='None'?d.selection.createRange().text:''):(d.getSelection?d.getSelection():'');void(vivi=window.open('http://my.poco.cn/fav/storeIt.php?t='+escape(d.title)+'&u='+escape(d.location.href)+'&c='+escape(t),'_blank','scrollbars=no,width=630,height=575,left=75,top=20,status=no,resizable=yes'));"><FONT color=#4cb509>POCO��ժ</FONT></A> |<A title='�����ٶ���ժ' style="FONT-WEIGHT: normal; FONT-SIZE: 12px; COLOR: #000000; TEXT-DECORATION: none" href="javascript:window.open('http://cang.baidu.com/do/add?it='+encodeURIComponent(document.title.substring(0,76))+'&iu='+encodeURIComponent(location.href)+'&fr=ien#nw=1','_blank','scrollbars=no,width=600,height=450,left=75,top=20,status=no,resizable=yes'); void 0"><SPAN style="PADDING-RIGHT: 5px; PADDING-LEFT: 5px; FONT-SIZE: 12px; PADDING-BOTTOM: 0px; MARGIN-LEFT: 10px; CURSOR: pointer; PADDING-TOP: 5px">�ٶ���ժ</A></SPAN></TD></TR>
</TABLE>
<TABLE width="360"  height="20" align="center">
<TR><TD ><INPUT class="inputb" TYPE=button VALUE="վ�ڶ�����ϵ" ONCLICK="window.open('messh.asp?name=<%=username%>','Sample')">
                <INPUT class="inputb" TYPE=button VALUE="��Ա�ղ�" ONCLICK="window.open('user/collection.asp?id=<%=Id%>','Sample','toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=no,resizable=no,copyhistory=no,width=170,height=80,left=300,top=100')">
                <input type="button" name="Submit" onClick='copyToClipBoard()' value="�����Ƽ�">
		<INPUT class="inputb" TYPE=button VALUE="�ٱ���Ϣ" ONCLICK="window.open('user/Bad_info.asp?id=<%=Id%>&city_oneid=<%=c1%>&city_twoid=<%=c2%>&city_threeid=<%=c3%>','Sample','toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=no,resizable=no,copyhistory=no,width=500,height=600,left=300,top=100')"></TR>
</TABLE></td>
                </tr>
              </table>
<table width="760" border="1" cellpadding="0" cellspacing="0" align="center">
        <tr>
          <td  style="border-style:none; padding:1px;" valign="top"><script src="<%=adjs_path("ads/js","info3","0_0_0")%>"></script></td>
</tr></table>    
              <br>
                <table width="750" border="0" align="center" cellPadding="5" cellSpacing="1" bgcolor="#E4E4E4"><%if newspl<3 then%>
                <tr> 
                  <td height="20" align="center" bgcolor="#FFFFFF" colspan="3"> <table width="99%" border="0" cellpadding="0" cellspacing="0">
                      <tr> 
                        <TD width="98%" align="center" bgcolor="#FBFBFB"style="border-top:1 dotted #99B509;border-bottom:1 dotted #99B509;border-left:1 dotted #99B509;"><SCRIPT src="user/rq_messagecs.asp?id=<%=id%>"></SCRIPT></TD>                                               
                      </tr>                      
			<TR align="center"> 
                        <TD colspan="3" ><strong><SCRIPT src="user/rq_message.asp?id=<%=id%>&hfcs=<%=hfcs%>"></SCRIPT></strong></TD>
                      </TR>                      
                      <TR> 
                        <TD height="10" colspan="3" ></TD>
                      </TR> 
                  </table></td>
                </tr><%end if%>
		<tr><td bgcolor="#FFFFFF"></td></tr>
                <tr> 
                  <td height="20" align="center" bgcolor="#FFFFFF"> 
 <script language="javascript">
<!--
//��֤������ȷ��
function checkemail(mdz){
var str=mdz;
//��JavaScript�У�������ʽֻ��ʹ��"/"��ͷ�ͽ���������ʹ��˫����
var Expression=/\w+([-+.']\w+)*@\w+([-.]\w+)*\.\w+([-.]\w+)*/;
var objExp=new RegExp(Expression);
if(objExp.test(str)==true){
return true;
}else{
return false;
}
}

//�����ֽ������ƣ����������=========
String.prototype.Length=function(){
temp=this.replace(/([^\x00-\xff])/g,"$1$1");
return temp.length
}

function form_onsubmit() {
    if((!checkemail(form.mymail.value)) && (document.form.mymail.value!=0))
	{
    alert("������Email��ַ����ȷ������������!");
    document.form.mymail.focus();
    return false;
    }
if (document.form.neirong.value==0)
	{
	  alert("������������")
	  document.form.neirong.focus()
	  return false
	 }
if (document.form.neirong.value.Length()>200)
     {
	 alert("�������ݳ�����200�ֽ��ڣ����������룡");
	  document.form.neirong.focus()
	  return false
  }
if (document.form.validate_codee.value==0)
	{
	  alert("������֤��")
	  document.form.validate_codee.focus()
	  return false
	 }
}
// --></script> 
<%if newspl>2 Then
ElseIf newspl>1 Then
Response.Write "��ֹ������"
else%>                 
<form name="form" id="form" onSubmit="return form_onsubmit()" action="user/reply.asp?id=<%=id%>&Readinfo=<%=Readinfo%>&xxusername=<%=username%>&biaoti=<%=biaoti%>&email=<%=email%>&dick=chk&city_oneid=<%=c1%>&city_twoid=<%=c2%>&city_threeid=<%=c3%>" method="POST">
<table border="0" cellpadding="0" style="border-style:solid; border-width:1px; padding:1px; border-collapse: collapse" bordercolor="#CCCCCC" width="100%">
<tr><td height="22" bgcolor="#FFE9D2" class="tablest2">
<p align="center"><b><a name="hf"></a>������Ϣ�������۲��ʼ�֪ͨ������</b></td></tr>
<tr><td height="29"><p style="line-height: 100%; margin-left: 20px">���Ļ�Ա����<input type="text" name="username" size="16" maxlength="50" value="<%=request.cookies("dick")("username")%>"> 
<SCRIPT src="../user/xinfodt.asp?action=username"></SCRIPT> <input type="checkbox" name="nm" value="on">�������� </td></tr>
<tr><td height="29"><p style="line-height: 100%; margin-left: 20px">���������䣺<input type="text" name="mymail" size="16" maxlength="50" value=""> <SCRIPT src="../user/xinfodt.asp?action=email"></SCRIPT></td></tr>
<tr><td><p style="line-height: 100%; margin-left: 20px">���۵�����(��200�ַ�)��<br><textarea rows="4" name="neirong" cols="50" maxlength="200"></textarea></td><tr><td height="29"><p style="line-height: 100%; margin-left: 20px">��ʽ��֤(�𰸼���)��<input type="text" name="validate_codee" size="4" onkeyup="value=value.replace(/[^\d|^\.]/g,'')"   onbeforepaste="clipboardData.setData('text',clipboardData.getData('text').replace(/[^\d|^\.]/g,''))"></td></tr>
<td height="22"><p align="center"><input type="submit" value="�ύ����" name="hf"></td></tr>
</table> </Form>  </td>
<td width="250" align="center" bgcolor="#FFFFFF"><script src="<%=adjs_path("ads/js","info5","0_0_0")%>"></script></td>
                </tr>
              </table>
<table width="200" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<img src="tt_getcodee.asp" alt="��֤��,�������?����ˢ����֤��" height="10" width="70" style="cursor : pointer;" onClick="this.src='tt_getcodee.asp?t='+(new Date().getTime());" /></td>
  </tr>
</table>
<%End if%>
 <table width="200" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td height="3"></td>
  </tr>
</table>
</td>
                </tr>

              </table>
			
    </td>
  </tr>
  </table>
  </td>
    </tr>
    <tr> 
      <td>&nbsp;</td>
    </tr>
  </table>
<table width="980" border="0" cellspacing="0" cellpadding="0">
        <tr><td  height="7"></td></tr>
</table>
<table width="980" border="0" align="center" cellpadding="1" cellspacing="1" class="tablest">
<tr><td height=10 colspan="2"> <script src="<%=adjs_path("ads/js","tc","0_0_0")%>"></script></td></tr>
</table>
<!-- Baidu Button BEGIN -->
<script type="text/javascript" id="bdshare_js" data="type=slide&amp;img=0&amp;uid=6475670" ></script>
<script type="text/javascript" id="bdshell_js"></script>
<script type="text/javascript">
	document.getElementById("bdshell_js").src = "http://bdimg.share.baidu.com/static/js/shell_v2.js?cdnversion=" + new Date().getHours();
</script>
<!-- Baidu Button END -->
<!--#include file="footer.asp" -->