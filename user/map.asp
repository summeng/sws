<!--#include file="../conn.asp"-->
<!--#include file="../config.asp"-->
<!--#include file="../SqlIn.Asp"-->
<%
'=====================================
'���칩����Ϣ��վ����ϵͳ
'��������Ƽ������ҳ�Ʒ
'�ͷ�����http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>
<%Dim sdName,Emap,com,rsmap,sqlmap,sdfg
com=trim(Request.QueryString("com"))
Call OpenConn
set rsmap=server.createobject("adodb.recordset")
sqlmap="select * from DNJY_user where username='"&com&"'" 	      
rsmap.open sqlmap,conn,1,1
if rsmap.bof or rsmap.eof then
response.write "û�д˻�Ա"
rsmap.close
set rsmap=nothing
response.End
Else
sdName=rsmap("sdName")
Emap=rsmap("Emap")
sdfg=rsmap("sdfg")
end If
rsmap.close
set rsmap=Nothing
If Emap="" Then
response.write "��������,�޷���ʾ��ͼ��"
response.end
End if%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title><%=sdName%>���ڵ�ͼ</title>
<meta name="keywords" content="<%=sdName%>���ڵ�ͼ" />
<meta name="description" content="<%=sdName%>���ڵ�ͼ" />
<link rel='stylesheet' type='text/css' href='images/sdimg/<%=sdfg%>/style.css' />
<script language="javascript" src=" http://api.51ditu.com/js/maps.js " ></script>
<script language="javascript" src="http://api.51ditu.com/js/ezmarker.js"></script>
<body style="margin:2px;padding-top:6px">
<div style='position:absolute;bottom:500px;padding:0px;text-align:center;height:16px;width:200px;background-color: #006699;z-index:1000;font-size:12px; overflow: hidden'><font color=#FFFFFF><%= sdName%>(�Ӽ�����������С��ͼ)</font></div>

<div id="myMap" style="position:relative; width:760px; height:400px;"></div>
<script language="javascript">
var maps = new LTMaps( "myMap" );
var icon = new LTIcon();

var point = new LTPoint(<%=Emap%>)
maps.centerAndZoom (point ,3);

var marker1 = new LTMarker( new LTPoint(<%=Emap%>) , icon );
maps.addOverLay( marker1 );

var text = new LTMapText( point );
text.setLabel( "<div style='padding:6px'>���ǹ�˾��λ��</div>" ); 
maps.addOverLay( text ); 


maps.handleKeyboard(); //���̲���֧��
maps.handleMouseScroll();//������֧��
var control1 = new LTSmallMapControl();//��׼���ſؼ�
maps.addControl( control1 );
</script>
</body>
</html>