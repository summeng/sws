<!--#include file="../conn.asp"-->
<!--#include file="../SqlIn.Asp"-->
<!--#include file="../config.asp"-->
<!--#include file="../Include/head.asp"-->
<!--#include file=../Include/Function.asp-->
<!--#include file=cookies.asp-->
<%
'=====================================
'���칩����Ϣ��վ����ϵͳ
'��������Ƽ������ҳ�Ʒ
'�ͷ�����http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>
<%call checkmanage("04")%>
<%
Function inHTML(str)
	Dim sTemp
	sTemp = str
	inHTML = ""
	If IsNull(sTemp) = True Then
		Exit Function
	End If
	sTemp = Replace(sTemp, "&", "&amp;")
	sTemp = Replace(sTemp, "<", "&lt;")
	sTemp = Replace(sTemp, ">", "&gt;")
	sTemp = Replace(sTemp, Chr(34), "&quot;")
	inHTML = sTemp
End Function	
Dim isse,page,ClassName,se_ClassName,searchkeys,selecttype,Descripm,Descrip,Content,addtime,author,origin,SavePathFileName,newsShared,ifshow,newstop,tj,rsnew
isse=ReplaceUnsafe(request.QueryString("isse"))
page=ReplaceUnsafe(Request.QueryString("page"))
ClassName=ReplaceUnsafe(Request.QueryString("ClassName"))
se_ClassName=ReplaceUnsafe(Request.QueryString("se_ClassName"))
searchkeys=ReplaceUnsafe(Request.QueryString("searchkeys"))
ID=ReplaceUnsafe(Request.QueryString("id"))
selecttype=ReplaceUnsafe(Request.QueryString("selecttype"))

set rsnew = server.CreateObject ("Adodb.recordset")
sql="select * from DNJY_userNews where ID="& id
rsnew.Open sql,conn,1,1
city_oneid=rsnew("city_oneid")
city_twoid=rsnew("city_twoid")
city_threeid=rsnew("city_threeid")
ClassName=rsnew("ClassName")
keywords = rsnew("keywords")
Descrip=rsnew("Descrip")
Title=rsnew("title")
Content=trim(rsnew("Content"))
If IsNull(rsnew("Content")) Then Content="��������..."
addtime=rsnew("addtime")
author=rsnew("author")
origin=rsnew("origin")
SavePathFileName = rsnew("SavePathFileName")
newsShared=rsnew("newsShared")
tj=rsnew("tj")
newstop=rsnew("newstop")
ifshow=rsnew("ifshow")
rsnew.close
set rsnew=nothing

Dim aSavePathFileName
if SavePathFileName<>"" then 
' �Ѵ�"|"���ַ���תΪ���飬���ڳ�ʼ�������
	aSavePathFileName = Split(SavePathFileName, "|")
' ����InitSelect����������ֵ����ʼֵ��������������ִ����������startup.asp�ļ��к�����˵������
else
	SavePathFileName=""
	aSavePathFileName = Split(SavePathFileName, "|")
	
end if
	
%>
<html>
<head>
<link rel="stylesheet" type="text/css" href="../css/css.css">
<style type="text/css">
<!--
body {
	background-color: #E3EBF9;
}
.style1 {color: #FFFFFF;
	font-weight: bold;
}
.style2 {color: #FFFFFF}
-->
</style>

<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>�޸���Ϣ</title>
<script language="JavaScript">
	<!--
	
	 function   checkDate(s)   
    {   
            var   isOk   =   false;   
  			tempArray   =   s.split('-');   
      		if   (tempArray.length   ==   3)   
    	    if   (   parseInt(tempArray[0]).toString().length   ==   4)   
            if   (   parseInt(tempArray[1])   >=1   &&   parseInt(tempArray[1])   <=12)   
     		if   (   parseInt(tempArray[2])   >=1   &&   parseInt(tempArray[2])   <=   31)   
            isOk   =   true;   
    
           return isOk;  
	 }   
   
	function checkForm(){
		
		//if (checkDate(document.postart.addtime.value)==false)
		//{alert("����ȷ��д���ڣ�����:2008-1-23����");
		//return false;
		//}
		 
		if (document.postart.ClassName.value=="") {
			alert("��ѡ����Ŀ��");
			document.postart.ClassName.focus();
			return false;
		}
		
		if (document.postart.txttitle.value.length == 0) {
			alert("��������⣡");
			document.postart.txttitle.focus();
			return false;
		}
		
        var editor=KindEditor.create('#editor');
        if (editor.isEmpty())
	    {
	    alert("�������ݲ���Ϊ�գ�")	  
	     return false
	     }

	}
	function doSubmit(){
		document.postart.submit();
	}
	
	//-->
</script>
<!--kindeditor�༭��-->
	<link rel="stylesheet" href="../KindEditor/themes/default/default.css" />
	<link rel="stylesheet" href="../KindEditor/plugins/code/prettify.css" />
	<script charset="utf-8" src="../KindEditor/kindeditor.js"></script>
	<script charset="utf-8" src="../KindEditor/lang/zh_CN.js"></script>
	<script charset="utf-8" src="../KindEditor/plugins/code/prettify.js"></script>
	<script>
		KindEditor.ready(function(K) {
			var editor = K.create('textarea[name="d_content"]', {
				cssPath : '../KindEditor/plugins/code/prettify.css',
				uploadJson : '../KindEditor/asp/upload_json.asp?action=usernews',
				afterBlur: function(){this.sync();},//ʧȥ����ִ��this���ֵ
				allowFileManager : true,				
				afterCreate : function() {
					var self = this;
					K.ctrl(document, 13, function() {
						self.sync();
						K('form[name=example]')[0].submit();
					});
					K.ctrl(self.edit.doc, 13, function() {
						self.sync();
						K('form[name=example]')[0].submit();
					});
				}
			});
			prettyPrint();
		});
	</script>
<!--kindeditor�༭��END-->
</head>
<body>
<div align="center">
<form method="POST" action="UserNews_ModifySave.asp?isse=<%=isse%>&id=<%=id%>&page=<%=page%>&searchkeys=<%=searchkeys%>&selecttype=<%=selecttype%>&se_ClassName=<%=se_ClassName%>" name="postart" onSubmit="return checkForm()">
	
	<input type=hidden name=d_savepathfilename  value="<%=SavePathFileName%>">
            
     <table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#ADAED6" width="70%">
	<TR> 
    <TD height="5" colspan="8" bgcolor="#799AE1" align="center"><FONT color="#FFFFFF" style="font-size:14px">�޸���Ϣ--����Ϣ��ĿIDΪ��<b><%=ID%></FONT></TD>
    </TR>

      <tr> 
                  <td width="10%" align="right" valign="middle"><div align="center">����������</td>
      <td colspan="2" valign="top" >�� 
<%Dim rsi,count
set rsi=conn.execute("select * from DNJY_city where id>0 and twoid>0 and threeid=0")
%>
      <SCRIPT language = "JavaScript">
var onecount;
onecount=0;
subcat = new Array();
        <%
		//dim count:
		count = 0
        do while not rsi.eof 
        %>
subcat[<%=count%>] = new Array("<%=rsi("city")%>","<%=rsi("id")%>","<%=rsi("twoid")%>");
        <%
        count = count + 1
        rsi.movenext
        loop
        rsi.close
		set rsi=nothing
        %>
onecount=<%=count%>;
</SCRIPT>
<%
set rsi=conn.execute("select * from DNJY_city where id>0 and twoid>0 and threeid>0")
%>
 <SCRIPT language = "JavaScript">
var onecount4;
onecount4=0;
subcat4 = new Array();
        <%
		dim count4:count4 = 0
        do while not rsi.eof 
        %>
subcat4[<%=count4%>] = new Array("<%=rsi("city")%>","<%=rsi("id")%>","<%=rsi("twoid")%>","<%=rsi("threeid")%>");
        <%
        count4 = count4 + 1
        rsi.movenext
        loop
        rsi.close
		set rsi = nothing
        %>
onecount4=<%=count4%>;

function changelocation(locationid)
    {
    document.addNEWS.city_two.length = 0; 
	document.addNEWS.city_two.options[0] = new Option('ѡ�����','');
	document.addNEWS.city_three.length = 0; 
	document.addNEWS.city_three.options[0] = new Option('ѡ�����','');
    var locationid=locationid;
    var i;
    for (i=0;i < onecount; i++)
        {
            if (subcat[i][1] == locationid)
            { 
                document.addNEWS.city_two.options[document.addNEWS.city_two.length] = new Option(subcat[i][0], subcat[i][2]);
            }        
        }
        
    }    
	
	function changelocation4(locationid,locationid1)
    {
    document.addNEWS.city_three.length = 0; 
    document.addNEWS.city_three.options[0] = new Option('ѡ�����','');
    var locationid=locationid;
	 var locationid1=locationid1;
    var i;
    for (i=0;i < onecount4; i++)
        {
            if (subcat4[i][2] == locationid)
            { 
			if (subcat4[i][1] == locationid1)
			{
                document.addNEWS.city_three.options[document.addNEWS.city_three.length] = new Option(subcat4[i][0], subcat4[i][3]);
            }        
        }
       } 
    }            
              </SCRIPT>
      <SELECT name="city_one" size="1" id="select2" onChange="changelocation(document.addNEWS.city_one.options[document.addNEWS.city_one.selectedIndex].value)">
            <OPTION value="" selected>ѡ�����</OPTION>
  <%set rsi=conn.execute("select * from DNJY_city where id>0 and twoid=0 and threeid=0 order by indexid asc")
if rsi.eof or rsi.bof then
response.write "<option value=''>û�з���</option>"
else
do until rsi.eof%>
  <OPTION value="<%=rsi("id")%>" <%if rsi("id")=city_oneid then%>selected<%end if%>><%=rsi("city")%></OPTION>
 <%rsi.movenext
    loop
	%>
	<%end if
rsi.close
set rsi = nothing
%>
      </SELECT> 
	  <SELECT name="city_two" id="city_two" onChange="changelocation4(document.addNEWS.city_two.options[document.addNEWS.city_two.selectedIndex].value,document.addNEWS.city_one.options[document.addNEWS.city_one.selectedIndex].value)">
    <OPTION value="" selected>ѡ�����</OPTION>
   <%
set rsi=conn.execute("select * from DNJY_city where id="&city_oneid&" and twoid>0 and threeid=0")
if rsi.eof and rsi.bof then
response.write "<option value=''>û�з���</option>"
else
do until rsi.eof%>
  <OPTION value="<%=rsi("twoid")%>" <%if rsi("twoid")=city_twoid then%>selected<%end if%>><%=rsi("city")%></OPTION>
 <%rsi.movenext
    loop
	%>
	<%end if
rsi.close
set rsi = nothing
%>
	</SELECT>
	     <SELECT name="city_three" id="city_three">
         <OPTION value="" selected>ѡ�����</OPTION>
		<%
set rsi=conn.execute("select * from DNJY_city where id="&city_oneid&" and twoid="&city_twoid&" and threeid<>0")
if rsi.eof and rsi.bof then
response.write "<option value=''>û�з���</option>"
else
do until rsi.eof%>
<OPTION value="<%=rsi("threeid")%>" <%if rsi("threeid")=city_threeid then%>selected<%end if%>><%=rsi("city")%></OPTION>
   <% rsi.movenext
    loop
	%>
<%end if
rsi.close
set rsi = nothing
%>
    </SELECT></td>
    </tr>
                <tr> 
                  <td width="10%" align="right" valign="middle"><div align="center">��Ϣ��Ŀ��</div></td>
                 
                  <td >

          <select name="ClassName" size="1" id="ClassName">
            <option value="" selected>ѡ����Ŀ</option>
<%set rs=conn.execute("select * From DNJY_userNewsClass")
if rs.eof or rs.bof then
	response.write "<option value=''>û�з���</option>"
else
	do until rs.eof
%>
<option value="<%=rs("classname")%>" <%if rs("ClassName")=ClassName then%>selected<%end if%>><%=rs("ClassName")%></option> 
<%   rs.movenext
     loop
	
end if
rs.close
set rs = nothing
%>
          </select>
          <span class="STYLE2">*</span></td>
                </tr>
                
                
                <tr> 
                  <td width="10%" align="right"  valign="middle" ><div align="center">��&nbsp;&nbsp;&nbsp; �⣺</div></td>
                  <td  colspan="2"><input name="txttitle" type="TEXT" id="txttitle"  value="<%=inHTML(Title)%>" size=60>
                  <span class="STYLE2">*</span></td>
                </tr>
                <tr>
                  <td align="right"  valign="middle" ><div align="center">�� �� �֣�</div></td>
                  <td  colspan="2"><input name="keywords" type="TEXT" id="keywords"  value="<%=keywords%>" size=60>
                  (��������ؼ���Ϊ:����_��վ��)</td>
                </tr>
				<tr>
                  <td align="right" valign="middle"  ><div align="center">������Ϣ��</div></td>
                  <td colspan=2 ><textarea name="Descrip" cols="75" rows="3" id="Descrip"><%=Descrip%></textarea>
                  (��������������ϢΪ����ǰ150�ַ�)</td>
                </tr>
   				<tr> 
                  <td width="10%"><div align="center">��&nbsp;&nbsp;&nbsp; �ߣ�</div></td>
                  <td  colspan="2"><input name="author"  type="TEXT" value="<%=author%>" size=20>
                  &lt;=��<FONT 
      color=blue><FONT style="CURSOR: hand" onclick='author.value="δ֪"' 
      color=#993300>δ֪</FONT></FONT>����<FONT color=blue><FONT 
      style="CURSOR: hand" onclick="author.value='����'" 
      color=#993300>����</FONT></FONT>����<FONT color=blue><FONT 
      style="CURSOR: hand" onclick="author.value='վ��'" 
      color=#993300>վ��</FONT></FONT>�� </td>
                </tr>
				<tr> 
                  <td width="10%"><div align="center">��&nbsp;&nbsp;&nbsp; Դ��</div></td>
                  <td  colspan="2"><input name="origin" type="TEXT" value="<%=origin%>" size=20>
                  &lt;=��<FONT 
      color=blue><FONT style="CURSOR: hand" onclick="origin.value='����'" 
      color=#993300>����</FONT></FONT>����<FONT color=blue><FONT 
      style="CURSOR: hand" onclick="origin.value='��վ'" 
      color=#993300>��վ</FONT></FONT>����<FONT color=blue><FONT 
      style="CURSOR: hand" onclick="origin.value='������'" 
      color=#993300>������</FONT></FONT>��</td>
                </tr >
      			<tr> 
                  <td width="10%"><div align="center">��&nbsp;&nbsp;&nbsp;&nbsp;����</div></td>
                  <td  colspan="2"> <%if newstop=1 then%>               
                <input type="radio" name="newstop" value="1" id="newstop" checked>
                 �̶�&nbsp;&nbsp;&nbsp;&nbsp;
                <input type="radio" name="newstop" value="0" id="newstop">
                ���̶�
                <%else%>                         
                <input type="radio" name="newstop" value="1" id="newstop">
                 �̶�&nbsp;&nbsp;&nbsp;&nbsp;
                <input type="radio" name="newstop" value="0" id="newstop" checked>
                ���̶�<%end if%>  
                </td>
                </tr>  
      			<tr> 
                  <td width="10%"><div align="center">��&nbsp;&nbsp;&nbsp;&nbsp;����</div></td>
                  <td  colspan="2"> <%if tj=True then%>               
                <input type="radio" name="tj" value="True" id="tj" checked>
                 �Ƽ�&nbsp;&nbsp;&nbsp;&nbsp;
                <input type="radio" name="tj" value="false" id="tj">
                ���Ƽ�
                <%else%>                         
                <input type="radio" name="tj" value="True" id="tj">
                 �Ƽ�&nbsp;&nbsp;&nbsp;&nbsp;
                <input type="radio" name="tj" value="false" id="tj" checked>
                ���Ƽ�<%end if%>
                </td>
                </tr>  
      			<tr> 
                  <td width="10%"><div align="center">������ˣ�</div></td>
                  <td  colspan="2"> <%if ifshow=1 then%>               
                <input type="radio" name="ifshow" value="1" id="ifshow" checked>
                 ͬ��&nbsp;&nbsp;&nbsp;&nbsp;
                <input type="radio" name="ifshow" value="0" id="ifshow">
                ��ͬ��
                <%else%>                         
                <input type="radio" name="ifshow" value="1" id="ifshow">
                 ͬ��&nbsp;&nbsp;&nbsp;&nbsp;
                <input type="radio" name="ifshow" value="0" id="ifshow" checked>
                ��ͬ��<%end if%>  (ͨ����˺�ſ�����վ�����Ŀ��ʾ)
                </td>
                </tr>                     
                
                <tr><td width="10%"><div align="center">��Ѷ���ݣ�</div></td> 
                  <td height="25" colspan=3 align="left" ><textarea id="editor" name="d_content" style="width:670px;height:400px;"><%=Content%></textarea><input type="checkbox" name="T_SaveImg" value="1" /> ����������Զ��ͼƬ������&nbsp;&nbsp;<input type="checkbox" name="AspJpeg" value="1" checked/>���ϴ�ͼƬ��ˮӡ<br></td>
                </tr>
                <tr>
                  <td height="20" align=center ><div align="center">�������ڣ�</div></td>
                  <td height="20" colspan="2"  ><input name="addtime" type="TEXT" size=25 maxlength=20  value="<%=addtime%>">
&lt;=�õ�ǰ���ڣ�<FONT color=blue><FONT 
      style="CURSOR: hand" onclick="addtime.value='<%=now()%>'" 
      color=#993300><%=now()%></FONT></FONT> </td>
                </tr>

                <tr> 
                  <td height="20" colspan=3 align=center > 
                    <input style="height:20px;" name="cmdok" type="submit" class="button" value=" �� �� " >
                    &nbsp; 
                    <input style="height:20px;" name="cmdcancel" type="reset" class="button" value=" �� д " >
&nbsp;&nbsp;                  </td>
                </tr>
    </table>


</form>
</div>
<%
Call CloseRs(rs)
Call CloseDB(conn)
%>