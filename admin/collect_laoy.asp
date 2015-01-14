<!--#include file="../conn.asp"-->
<!--#include file="../config.asp"-->
<!--#include file=cookies.asp-->
<!--#include file="inc/Function.asp"-->
<!--#include file="inc/ubbcode.asp"-->
<!--#include file="inc/clsCache.asp"-->
<!--#include file="../Include/Function.asp"-->
<%
'=====================================
'天天供求信息网站管理系统
'天天网络科技工作室出品
'客服中心http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>
<html>
<head>
<title>信息采集系统</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="../css/manage.css" type="text/css">
<script language="JavaScript">
<!--
function CheckAll(form)  {
  for (var i=0;i<form.elements.length;i++)    {
    var e = form.elements[i];
    if (e.name != 'chkall'){
	e.checked = form.chkall.checked;
	}
   }
  }
//-->
</script>
</head>
<body>
<%Call OpenConn
dim StyleConn
dim ActionName,ActionType,MdbName
dim Dr,Rows,i
Select Case Request("action")
	Case "LoadSkin"
		Call LoadSkin()
	Case "Load"
		Call Load()
	Case "ExportSkin"
		Call ExportSkin()
	Case "ClearLoadData"
		Call ClearLoadData()
	Case Else
		Call Main()
End Select

Sub Main()
	If Request("action")="import" Then
		ActionName = "导入"
		ActionType = "LoadSkin"
		MdbName = "collect_info.rules.mdb"
	Else
		Mdbname = "collect_info.rules.mdb"
		ActionName ="导出"
		ActionType ="ExportSkin"
	End If
%>
<table width="98%" border="0" cellpadding="3" cellspacing="1" bgcolor="#f7f7f7" class="admintable">
<form name="form1" method="post" action="?">
<%
    Set rs = server.CreateObject("adodb.recordset")
	sql = "Select * from DNJY_xx_Item Order by ItemID"&DN_OrderType&""   
	If ActionType = "LoadSkin" Then
		SkinConnection(MdbName)
		rs.open sql,StyleConn,1,3
	Else
		rs.open sql,conn,1,3
	End If	
    if not(rs.bof and rs.eof) then
	Dr = rs.GetRows()
%>
  <tr>
    <td colspan="6" class="admintitle">信息采集规则列表</td>
  </tr>
  <tr align="center">
  <td width="5%" class="ButtonList">选择</td>
    <td width="15%" class="ButtonList">项目名称</td>
	<td width="25%" class="ButtonList">所属地区</td>
    <td width="25%" class="ButtonList">所属栏目</td>
    <td width="25%"  class="ButtonList">上次采集</td>
    <td width="5%"  class="ButtonList">状态</td>
  </tr>
<%Dim ItemCollecDate
		Rows = ubound(Dr,2)
		for i=0 to Rows
%>
  <tr>
      <td align="center" bgcolor="#FFFFFF" ><input name="ItemID" type="checkbox" class="noborder" value="<%=Dr(0,i)%>"></td>
    <td bgcolor="#FFFFFF" ><%=Dr(1,i)%></td>
	<td align="center" bgcolor="#FFFFFF" ><%=Dr(3,i)%><%If Dr(5,i)<>"" Then Response.Write " / "%><%=Dr(5,i)%><%If Dr(7,i)<>"" Then Response.Write " / "%><%=Dr(7,i)%></td>
    <td align="center" bgcolor="#FFFFFF" ><%=Dr(9,i)%><%If Dr(11,i)<>"" Then Response.Write " / "%><%=Dr(11,i)%><%If Dr(13,i)<>"" Then Response.Write " / "%><%=Dr(13,i)%></td>
    <td align="center" bgcolor="#FFFFFF" ><%
	set rs=server.createobject("adodb.recordset")
       Set Rs=conn.execute("select Top 1 CollecDate From DNJY_xx_Histroly Where ItemID=" & Dr(0,i) & " Order by HistrolyID"&DN_OrderType&"")
       If Not Rs.Eof Then
          ItemCollecDate=rs("CollecDate")
       Else
          ItemCollecDate=""
       End if
       Set Rs=Nothing
       if ItemCollecDate<>"" then
          Response.Write ItemCollecDate
       Else
          Response.Write "尚无记录"
       End If
       %></td>
    <td align="center" bgcolor="#FFFFFF" ><b>
      <%If Dr(62,i)=True then
                 Response.write "√"
          Else
                 Response.write "<font color=red>×</font>"
          End If
        %>
    </b></td>
  </tr>
<%
        next
%>
<tr>
  <td height="25" align="center" bgcolor="#FFFFFF" class="td"><input name=chkall type=checkbox class="noborder" onClick="CheckAll(this.form)" value=on /></td>
  <td height="25" bgcolor="#FFFFFF" class="td">全选</td>
  <td height="25" colspan="4" bgcolor="#FFFFFF" class="td">&nbsp;</td>
  </tr>

<tr>
  <td height="24" colspan="6" align="center" bgcolor="#FFFFFF">
    <input type="submit" name="Submit" value=" <%=ActionName%>采集规则"><%If request("action")<>"import" then%>
    <input type="submit" name="Submit" value=" 清空导出目标库 " onClick="action.value='ClearLoadData';">
    <%end if%>
    <input type="hidden" name="action" value="<%=ActionType%>"></td>
  </tr>
<%
	else
	Call Alert ("温馨提示：该规则库里没有采集规则!","collect_management.asp")
	end if
%>
</form>
</table>
<br>
<%
End Sub

Sub Load()
%>
<form action="?action=import" method=post>
  <table border="0" cellpadding="3"  cellspacing="1" bgcolor="#f7f7f7" class="admintable">
    <tr>
      <td colspan="2" class="admintitle">导入采集规则(第一步)</td>
    </tr>
    <tr>
      <td width="20%" height="30" bgcolor="#FFFFFF" class="td">导入采集规则数据库名：</td>
      <td width="80%" bgcolor="#FFFFFF" class="td"><input type="text" name="TemplateMdb" size="30" value="collect_info.rules.mdb"></td>
    </tr>
    <tr>
      <th colspan="2" align="left" bgcolor="#FFFFFF"><input type="submit" name="submit" value="下一步"></th>
    </tr>
  </table>
</form>
<%
End Sub

Sub SkinConnection(MdbName)
	On Error Resume Next 
	Set StyleConn = Server.CreateObject("ADODB.Connection")
	StyleConn.open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("../temp/collect_info.rules.mdb")
	If Err.Number ="-2147467259"  Then 
		Call Alert (""& MdbName &"数据库不存在。",-1)
	End If
End Sub

Sub LoadSkin()
	dim ItemID,ChannelID,oChannelID
	dim tRs,tSql,nRs,nSql
	MdbName = trim(Request.form("TemplateMdb"))
	ItemID = Replace(Request.form("ItemID")," ","")

	If ItemID = "" then
		Call Alert ("请选择要导入的规则！",-1)
	End If
	
	tSql = "Select * from [DNJY_xx_Item] Where ItemID in("& ItemID &")"
	SkinConnection(MdbName)
	Set tRs = StyleConn.Execute(tSql)
	
	Set nRs = Server.CreateObject("ADODB.RecordSet")
	nRs.open "[DNJY_xx_Item]",conn,1,3
	
	Do while Not tRs.Eof
	nRs.AddNew

      nRs("ItemName")=tRs(1)'按数据表中字段出现顺序！！！	  
      nRs("city_oneid")=tRs(2)
      nRs("city_one")=tRs(3)
      nRs("city_twoid")=tRs(4)
      nRs("city_two")=tRs(5)
      nRs("city_threeid")=tRs(6)
      nRs("city_three")=tRs(7)
      nRs("type_oneid")=tRs(8)
      nRs("type_one")=tRs(9)
      nRs("type_twoid")=tRs(10)
      nRs("type_two")=tRs(11)
      nRs("type_threeid")=tRs(12)
      nRs("type_three")=tRs(13)
      nRs("WebName")=tRs(14)'
      nRs("WebUrl")=tRs(15)'
      nRs("ItemDemo")=tRs(16)'

      nRs("LoginType")=strint(tRs(17))''''''''''''''''''''''''''''''''''''''''''''''''''''''''
      nRs("LoginUrl")=tRs(18)          '登录
      nRs("LoginPostUrl")=tRs(19)
      nRs("LoginUser")=tRs(20)
      nRs("LoginPass")=tRs(21)
      nRs("LoginFalse")=tRs(22)
      nRs("ListStr")=tRs(23)            '列表地址
      nRs("LsString")=tRs(24)          '列表
      nRs("LoString")=tRs(25)
      nRs("ListPaingType")=tRs(26)
      nRs("LPsString")=tRs(27)          
      nRs("LPoString")=tRs(28)
      nRs("ListPaingStr1")=tRs(29)
      nRs("ListPaingStr2")=tRs(30)
      nRs("ListPaingID1")=tRs(31)
      nRs("ListPaingID2")=tRs(32)
      nRs("ListPaingStr3")=tRs(33)
      nRs("HsString")=tRs(34)  
      nRs("HoString")=tRs(35)
      nRs("HttpUrlType")=tRs(36)
      nRs("HttpUrlStr")=tRs(37)
      nRs("TsString")=tRs(38)          '标题
      nRs("ToString")=tRs(39)
      nRs("CsString")=tRs(40)          '正文
      nRs("CoString")=tRs(41)
      nRs("DateType")=tRs(42)      '作者
      nRs("DsString")=tRs(43)          
      nRs("DoString")=tRs(44)
      nRs("AuthorType")=tRs(45)      '作者
      nRs("AsString")=tRs(46)          
      nRs("AoString")=tRs(47)
      nRs("AuthorStr")=tRs(48)
      nRs("CopyFromType")=tRs(49)  '来源
      nRs("FsString")=tRs(50)          
      nRs("FoString")=tRs(51)
      nRs("CopyFromStr")=tRs(52)
      nRs("KeyType")=tRs(53)            '关键词
      nRs("KsString")=tRs(54)          
      nRs("KoString")=tRs(55)
      nRs("KeyStr")=tRs(56)
      nRs("NewsPaingType")=tRs(57)            '关键词
      nRs("NPsString")=tRs(58)          
      nRs("NPoString")=tRs(59)
      nRs("NewsPaingStr")=tRs(60)
      nRs("NewsPaingHtml")=tRs(61)
      nRs("Flag")=tRs(62)
      nRs("PaginationType")=tRs(63)
      nRs("MaxCharPerPage")=tRs(64)
      nRs("ReadLevel")=tRs(65)
      nRs("Stars")=tRs(66)
      nRs("ReadPoint")=tRs(67)
      nRs("Hits")=tRs(68)
      nRs("UpDateType")=tRs(69)
      nRs("UpDateTime")=tRs(70)
      nRs("IncludePicYn")=tRs(71)
      nRs("DefaultPicYn")=tRs(72)
      nRs("OnTop")=tRs(73)
      nRs("Elite")=tRs(74)
      nRs("Hot")=tRs(75)
      nRs("SkinID")=tRs(76)
      nRs("TemplateID")=tRs(77)
      nRs("Script_Iframe")=tRs(78)
      nRs("Script_Object")=tRs(79)
      nRs("Script_Script")=tRs(80)
      nRs("Script_Div")=tRs(81)
      nRs("Script_Class")=tRs(82)
      nRs("Script_Span")=tRs(83)
      nRs("Script_Img")=tRs(84)
      nRs("Script_Font")=tRs(85)
      nRs("Script_A")=tRs(86)
      nRs("Script_Html")=tRs(87)
      nRs("CollecListNum")=tRs(88)
      nRs("CollecNewsNum")=tRs(89)
      nRs("Passed")=tRs(90)
      nRs("SaveFiles")=tRs(91)
      nRs("CollecOrder")=tRs(92)
      nRs("LinkUrlYn")=tRs(93)
      nRs("InputerType")=tRs(94)
      nRs("Inputer")=tRs(95)
      nRs("EditorType")=tRs(96)
      nRs("Editor")=tRs(97)
      nRs("ShowCommentLink")=tRs(98)
      nRs("Script_Table")=tRs(99)
      nRs("Script_Tr")=tRs(100)
      nRs("Script_Td")=tRs(101)
      nRs("StopDateType")=tRs(102)
      nRs("SDsString")=tRs(103)
      nRs("SDoString")=tRs(104)
      nRs("StopDate")=tRs(105)
      nRs("TelType")=tRs(106)
      nRs("TelsString")=tRs(107)
      nRs("TeloString")=tRs(108)
      nRs("TelStr")=tRs(109)

      nRs("mailType")=tRs(110)
      nRs("mailsString")=tRs(111)
      nRs("mailoString")=tRs(112)
      nRs("mailStr")=tRs(113)

      nRs("QQType")=tRs(114)
      nRs("QsString")=tRs(115)
      nRs("Qostring")=tRs(116)
      nRs("QQStr")=tRs(117)

      nRs("leixingType")=tRs(118)
      nRs("LxsString")=tRs(119)
      nRs("LxoString")=tRs(120)
      nRs("leixingStr")=tRs(121)

      nRs("sjType")=tRs(122)
      nRs("sjsString")=tRs(123)
      nRs("sjoString")=tRs(124)
      nRs("sjStr")=tRs(125)

      nRs("jgType")=tRs(126)
      nRs("jgsString")=tRs(127)
      nRs("jgoString")=tRs(128)
      nRs("jgStr")=tRs(129)

      nRs("dzType")=tRs(130)
      nRs("dzsString")=tRs(131)
      nRs("dzoString")=tRs(132)
      nRs("dzStr")=tRs(133)

      nRs("ybType")=tRs(134)
      nRs("ybsString")=tRs(135)
      nRs("yboString")=tRs(136)
      nRs("ybStr")=tRs(137)
      nRs("cj_watermark")=tRs(138)

		nRs.Update
		tRs.MoveNext
	Loop	
	Call Alert ("采集规则导入成功！","collect_management.asp")
End Sub

Sub ExportSkin()
	dim ItemID,ChannelID,oChannelID
	dim tRs,tSql,nRs,nSql
	
	MdbName = trim(Request.form("TemplateMdb"))
	ItemID = Replace(Request.form("ItemID")," ","")

	If ItemID = "" then
		Call Alert ("请选择要导出的规则！",-1)
	End If

	'源库
	nSql = "Select * from [DNJY_xx_Item] Where ItemID in("& ItemID &")"
	Set nRs = conn.Execute(nSql)

	'目标库
	Set tRs = Server.CreateObject("ADODB.RecordSet")
	SkinConnection(MdbName)
	tRs.open "[DNJY_xx_Item]",StyleConn,1,3
	
	Do while Not nRs.Eof
		tRs.AddNew
 
      tRs("ItemName")=nRs(1)'按数据表中字段出现顺序！！！	  
      tRs("city_oneid")=nRs(2)
      tRs("city_one")=nRs(3)
      tRs("city_twoid")=nRs(4)
      tRs("city_two")=nRs(5)
      tRs("city_threeid")=nRs(6)
      tRs("city_three")=nRs(7)
      tRs("type_oneid")=nRs(8)
      tRs("type_one")=nRs(9)
      tRs("type_twoid")=nRs(10)
      tRs("type_two")=nRs(11)
      tRs("type_threeid")=nRs(12)
      tRs("type_three")=nRs(13)
      tRs("WebName")=nRs(14)'
      tRs("WebUrl")=nRs(15)'
      tRs("ItemDemo")=nRs(16)'

      tRs("LoginType")=strint(nRs(17))''''''''''''''''''''''''''''''''''''''''''''''''''''''''
      tRs("LoginUrl")=nRs(18)          '登录
      tRs("LoginPostUrl")=nRs(19)
      tRs("LoginUser")=nRs(20)
      tRs("LoginPass")=nRs(21)
      tRs("LoginFalse")=nRs(22)
      tRs("ListStr")=nRs(23)            '列表地址
      tRs("LsString")=nRs(24)          '列表
      tRs("LoString")=nRs(25)
      tRs("ListPaingType")=nRs(26)
      tRs("LPsString")=nRs(27)          
      tRs("LPoString")=nRs(28)
      tRs("ListPaingStr1")=nRs(29)
      tRs("ListPaingStr2")=nRs(30)
      tRs("ListPaingID1")=nRs(31)
      tRs("ListPaingID2")=nRs(32)
      tRs("ListPaingStr3")=nRs(33)
      tRs("HsString")=nRs(34)  
      tRs("HoString")=nRs(35)
      tRs("HttpUrlType")=nRs(36)
      tRs("HttpUrlStr")=nRs(37)
      tRs("TsString")=nRs(38)          '标题
      tRs("ToString")=nRs(39)
      tRs("CsString")=nRs(40)          '正文
      tRs("CoString")=nRs(41)
      tRs("DateType")=nRs(42)      '作者
      tRs("DsString")=nRs(43)          
      tRs("DoString")=nRs(44)
      tRs("AuthorType")=nRs(45)      '作者
      tRs("AsString")=nRs(46)          
      tRs("AoString")=nRs(47)
      tRs("AuthorStr")=nRs(48)
      tRs("CopyFromType")=nRs(49)  '来源
      tRs("FsString")=nRs(50)          
      tRs("FoString")=nRs(51)
      tRs("CopyFromStr")=nRs(52)
      tRs("KeyType")=nRs(53)            '关键词
      tRs("KsString")=nRs(54)          
      tRs("KoString")=nRs(55)
      tRs("KeyStr")=nRs(56)
      tRs("NewsPaingType")=nRs(57)            '关键词
      tRs("NPsString")=nRs(58)          
      tRs("NPoString")=nRs(59)
      tRs("NewsPaingStr")=nRs(60)
      tRs("NewsPaingHtml")=nRs(61)
      tRs("Flag")=nRs(62)
      tRs("PaginationType")=nRs(63)
      tRs("MaxCharPerPage")=nRs(64)
      tRs("ReadLevel")=nRs(65)
      tRs("Stars")=nRs(66)
      tRs("ReadPoint")=nRs(67)
      tRs("Hits")=nRs(68)
      tRs("UpDateType")=nRs(69)
      tRs("UpDateTime")=nRs(70)
      tRs("IncludePicYn")=nRs(71)
      tRs("DefaultPicYn")=nRs(72)
      tRs("OnTop")=nRs(73)
      tRs("Elite")=nRs(74)
      tRs("Hot")=nRs(75)
      tRs("SkinID")=nRs(76)
      tRs("TemplateID")=nRs(77)
      tRs("Script_Iframe")=nRs(78)
      tRs("Script_Object")=nRs(79)
      tRs("Script_Script")=nRs(80)
      tRs("Script_Div")=nRs(81)
      tRs("Script_Class")=nRs(82)
      tRs("Script_Span")=nRs(83)
      tRs("Script_Img")=nRs(84)
      tRs("Script_Font")=nRs(85)
      tRs("Script_A")=nRs(86)
      tRs("Script_Html")=nRs(87)
      tRs("CollecListNum")=nRs(88)
      tRs("CollecNewsNum")=nRs(89)
      tRs("Passed")=nRs(90)
      tRs("SaveFiles")=nRs(91)
      tRs("CollecOrder")=nRs(92)
      tRs("LinkUrlYn")=nRs(93)
      tRs("InputerType")=nRs(94)
      tRs("Inputer")=nRs(95)
      tRs("EditorType")=nRs(96)
      tRs("Editor")=nRs(97)
      tRs("ShowCommentLink")=nRs(98)
      tRs("Script_Table")=nRs(99)
      tRs("Script_Tr")=nRs(100)
      tRs("Script_Td")=nRs(101)
      tRs("StopDateType")=nRs(102)
      tRs("SDsString")=nRs(103)
      tRs("SDoString")=nRs(104)
      tRs("StopDate")=nRs(105)
      tRs("TelType")=nRs(106)
      tRs("TelsString")=nRs(107)
      tRs("TeloString")=nRs(108)
      tRs("TelStr")=nRs(109)

      tRs("mailType")=nRs(110)
      tRs("mailsString")=nRs(111)
      tRs("mailoString")=nRs(112)
      tRs("mailStr")=nRs(113)

      tRs("QQType")=nRs(114)
      tRs("QsString")=nRs(115)
      tRs("Qostring")=nRs(116)
      tRs("QQStr")=nRs(117)

      tRs("leixingType")=nRs(118)
      tRs("LxsString")=nRs(119)
      tRs("LxoString")=nRs(120)
      tRs("leixingStr")=nRs(121)

      tRs("sjType")=nRs(122)
      tRs("sjsString")=nRs(123)
      tRs("sjoString")=nRs(124)
      tRs("sjStr")=nRs(125)

      tRs("jgType")=nRs(126)
      tRs("jgsString")=nRs(127)
      tRs("jgoString")=nRs(128)
      tRs("jgStr")=nRs(129)

      tRs("dzType")=nRs(130)
      tRs("dzsString")=nRs(131)
      tRs("dzoString")=nRs(132)
      tRs("dzStr")=nRs(133)

      tRs("ybType")=nRs(134)
      tRs("ybsString")=nRs(135)
      tRs("yboString")=nRs(136)
      tRs("ybStr")=nRs(137)

      tRs("cj_watermark")=nRs(138)
		
		tRs.Update
		nRs.MoveNext
	Loop
	Call Alert ("导出采集规则成功",-1)
End Sub

Sub ClearLoadData()
	MdbName = trim(Request.form("TemplateMdb"))
	
	dim tRs
	Set tRs = Server.CreateObject("ADODB.RecordSet")
	SkinConnection(MdbName)
	tRs.open "[DNJY_xx_Item]",StyleConn,1,3
	
	Do while Not tRs.Eof
		tRs.delete
		tRs.MoveNext
	Loop
	Call CloseRs(tRs)
	Call Alert ("清空采集规则导出库完成","collect_laoy.asp")
End Sub
%>
</body>
</html>