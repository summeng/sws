<%
'=====================================
'������ҵ��վ����ϵͳ  TianTianCMS
'��������Ƽ������Ұ�Ȩ����
'�ٷ���վhttp://www.ip126.com
'����֧����̳ http://bbs.ip126.com/
'QQ:530051328 mail:bbfhq@sohu.com
'=====================================
%>
<%
'===========================================================
'�汾 1.0

'w,h ��ʾ��ͼ�Ŀ�Ⱥ͸߶�
'xy ��ͼ��ʼ��γ����
'z ��ͼ��ʼ���ű���
'FormName ������ֵ�ı����Ƽ����ط�ʽ
'Input ������ֵ���ı�������
'zC ���ſؼ� StandΪ��׼�ؼ� SmallΪ���׿ؼ�

' by netboy
'===========================================================

Response.Buffer = True
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 1
Response.Expires = 0
Response.CacheControl = "no-cache"
%>
<body style="margin:0">
<script language="javascript" src=" http://api.51ditu.com/js/maps.js " ></script>
<script language="javascript" src="http://api.51ditu.com/js/ezmarker.js"></script>
<div style='position:absolute;bottom:0px;padding:2px;text-align:center;height:18px;width:150px;background-color: #006699;z-index:1000;font-size:12px; overflow: hidden'><font color=#FFFFFF>���ڵ�ͼ�ϱ�ע��˾λ��</a></div>
<%
Dim Content,w,h,xy,z,Input,zC,i,Tempxy,Text
w = Request("w")
If w = "" or Isnull(w) Then w = "100%" Else w = w&"px"
h = Request("h")
If h = "" or Isnull(h) Then h = "100%" Else h = h&"px"
z = Request("z")
If z = "" or Isnull(z) Then z = 5
xy = Request("xy")
If Request("Action") <> "Show" Then
  If xy = "" or Isnull(xy) Then xy = "10884937,3441544":z = 13
Else
  If xy = "" or Isnull(xy) Then Response.Write "<table width='100%' height='100%' border='0' cellspacing='0' cellpadding='0' background='bg.gif'><tr><td><table border='0' cellspacing='0' cellpadding='0' align=center><tr><td><img src=images/nomap.gif></td></tr></table></td></tr></table>":Response.end
End If
For i=0 to Ubound(Split(xy,","))
  If Split(xy,",")(i) <> "" and isNull(Split(xy,",")(i)) = False Then
    Tempxy = Tempxy & Split(xy,",")(i) & ","
  End If 
Next
xy = Split(Tempxy,",")(0)&","&Split(Tempxy,",")(1)
Input = Request("Input")
FormName = Request("Form")
If FormName = "" Or IsNull(FormName) Then FormName = "dialogArguments.document.thisForm"
zC = Request("zC")
If zC = "" Or IsNull(zC) Then zC = "Stand"
Text = Request("Text")
%>
<div id="myMap" style="position:relative; width:<%=w%>; height:<%=h%>;"></div>
<script language="javascript">
var maps = new LTMaps( "myMap" );
var icon = new LTIcon();
icon.setImageUrl('centerPoi.gif');
maps.centerAndZoom ( new LTPoint(<%=xy%>),<%=z%>);
if (<%=FormName%>.<%=Input%>.value != ""){
var marker1 = new LTMarker( new LTPoint(<%=xy%>) , icon );
maps.addOverLay( marker1 );
}

maps.handleKeyboard(); //���̲���֧��
maps.handleMouseScroll();//������֧��
var control1 = new LT<%=zC%>MapControl();//��׼���ſؼ�
maps.addControl( control1 );
var control2 = new LTMarkControl();//��ע��ť
maps.addControl( control2 );

function getPoi(){
var poi = control2.getMarkControlPoint();
if (confirm('����ע��λ�þ�γ�ȷֱ�Ϊ��' + poi.getLongitude() + ','+ poi.getLatitude() + '\n�Ƿ񱣴�ñ�ע��')){
<%=FormName%>.<%=Input%>.value = poi.getLongitude() + "," + poi.getLatitude();
window.close();
}
}
LTEvent.addListener( control2 , "mouseup" , getPoi );
</script>
