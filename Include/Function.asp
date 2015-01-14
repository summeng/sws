<%
'=====================================
'天天供求信息网站管理系统
'天天网络科技工作室出品
'客服中心http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>
<%'获取主机地址，即http://www.ip126.com/list.asp中的www.ip126.com
Dim BaseUrl
BaseUrl = LCase(Replace(Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("URL"),Split(request.ServerVariables("SCRIPT_NAME"),"/")(ubound(Split(request.ServerVariables("SCRIPT_NAME"),"/"))),""))

'**************************************************
'函数名：GetUrl
'作  用：得到当前页面的地址，如果非80端口也会加上 
'**************************************************
Function GetUrl()   
On Error Resume Next   
Dim strTemp   
If LCase(Request.ServerVariables("HTTPS"))="off" Then   
strTemp="http://"
Else   
strTemp="https://"
End If   
strTemp=strTemp&Request.ServerVariables("SERVER_NAME")   
If Request.ServerVariables("SERVER_PORT")<>80 Then strTemp=strTemp&":"&Request.ServerVariables("SERVER_PORT")   
strTemp=strTemp&Request.ServerVariables("URL")   
If Trim(Request.QueryString)<>"" Then  strTemp=strTemp&"?"&Trim(Request.QueryString)   
GetUrl=strTemp   
End Function

'**************************************************
'函数名：Alert
'作  用：操作过程提示并决定去向 
'用法：Call Alert("提示内容","-1")
'**************************************************
Function Alert(message,gourl) 
    message = replace(message,"'","\'")
    If gourl="-1" then
        Response.Write ("<script language=javascript>alert('" & message & "');history.go(-1)</script>")
    ElseIf gourl="-2" then
        Response.Write ("<script language=javascript>alert('" & message & "');history.go(-2)</script>")
    ElseIf gourl="Close" then
		Response.Write ("<script language=javascript>alert('" & message & "');window.opener=null;window.close();</script>")
	Else
        Response.Write ("<script language=javascript>alert('" & message & "');location='" & gourl &"'</script>")
    End If
    Response.End()
End Function


'特殊字符过滤“替换”一
Function HTMLEncode(fString)
If Not IsNull(fString) Then
fString = trim(fString)
'fString = replace(fString, ";", "；")      '分号过滤
fString = replace(fString, "--", "——") '--过滤
fString = replace(fString, "%20", "")     '特殊字符过滤
fString = replace(fString, "==", "")      '==过滤
fString = replace(fString, ">", "&gt;")
fString = replace(fString, "<", "&lt;")
fString = replace(fString, "%", "&#37;")     '%过滤
fString = Replace(fString, CHR(32), " ")    '&nbsp;
fString = Replace(fString, CHR(9), " ")     '&nbsp;
fString = Replace(fString, CHR(34), "&quot;")
fString = Replace(fString, CHR(39), "&#39;") '单引号过滤
fString = Replace(fString, CHR(13), "")
fString = Replace(fString, CHR(10) & CHR(10), "</P><P> ")
fString = Replace(fString, CHR(10), "<BR> ")
HTMLEncode = fString
End If
End Function

'特殊字符过滤“替换”二
Function HtmlEncode2(Content)
IF content="" or isnull(content) Then exit function
  Content = trim(Content)
  Content = Replace(Content,"%20"  , ""       )'特殊字符过滤
  Content = Replace(Content,chr(62),  "＞"  )' > 
  Content = Replace(Content,chr(60),  "＜"  )' < 
  Content = Replace(Content,chr(39),  "＇"    )' ' 
  Content = Replace(Content,chr(37),  "％"    )' % 
  Content = Replace(Content, vbcrlf,  ""      )
  Content = Replace(Content,chr(34),  "”")' "
  Content = Replace(Content,chr(40),  "（"    )' ( 
  Content = Replace(Content,chr(41),  "）"    )' ) 
  Content = Replace(Content,chr(91),  "［"    )' [ 
  Content = Replace(Content,chr(93),  "］"    )' ] 
  Content = Replace(Content,chr(123), "｛"    )' { 
  Content = Replace(Content,chr(125), "｝"    )' } 
  Content = Replace(Content, CHR(13),   "")   
  Content = Replace(Content,CHR(10), "")
  HtmlEncode2 = content 
End Function

'**************************************************
'函数名：chkdick
'作  用：用于字符过滤
'参  数：char ---要处理的字符串'             
'**************************************************
function chkdick(char)
if instr(char,"'") then 
call errdick(13)
response.end 
end if
'if instr(char,"|") then
'call errdick(13)
'response.end 
'end if
if instr(char,"<") then
call errdick(13)
response.end 
end if
if instr(char,">") then
call errdick(13)
response.end 
end If
'if instr(char,"(") then
'call errdick(13)
'response.end 
'end if
'if instr(char,")") then
'call errdick(13)
'response.end 
'end if
if instr(char,"&") then
call errdick(13)
response.end 
end if
'if instr(char,"%") then
'call errdick(13)
'response.end 
'end if
if instr(char,";") then
call errdick(13)
response.end 
end if
if instr(char,"$") then
call errdick(13)
response.end 
end if
if instr(char,"$") then
call errdick(13)
response.end 
end if
if instr(char,"╰") then
call errdick(13)
response.end 
end if
'if instr(char,"★") then
'call errdick(13)
'response.end 
'end if
if instr(char,"╮") then
call errdick(13)
response.end 
end if
'if instr(char,"※") then
'call errdick(13)
'response.end 
'end if
'if instr(char,"◆") then
'call errdick(13)
'response.end 
'end if
'if instr(char,"◎") then
'call errdick(13)
'response.end 
'end if
'if instr(char,"○") then
'call errdick(13)
'response.end 
'end if
'if instr(char,"▲") then
'call errdick(13)
'response.end 
'end if
'if instr(char,"△") then
'call errdick(13)
'response.end 
'end if
'if instr(char,"■") then
'call errdick(13)
'response.end 
'end if
'if instr(char,"〓") then
'call errdick(13)
'response.end 
'end if
end function


'*************************************************
'函数名：RemoveHTML_ttkj
'作  用：过滤HTML代码的函数包括过滤CSS和JS，只显示文字
'参  数：RemoveHTML_ttkj()
'*************************************************

Function RemoveHTML_ttkj(strHTML)
StrHtml = Replace(StrHtml,vbCrLf,"") 
StrHtml = Replace(StrHtml,Chr(13)&Chr(10),"") 
StrHtml = Replace(StrHtml,Chr(13),"") 
StrHtml = Replace(StrHtml,Chr(10),"") 
StrHtml = Replace(StrHtml," ","") 
StrHtml = Replace(StrHtml,"    ","")

Dim objRegExp, Match, Matches 
Set objRegExp = New Regexp 
objRegExp.IgnoreCase = True 
objRegExp.Global = True

'取闭合的<> 
objRegExp.Pattern = "<style(.+?)/style>" 
'进行匹配 
Set Matches = objRegExp.Execute(strHTML) 
' 遍历匹配集合，并替换掉匹配的项目 
For Each Match in Matches 
strHtml=Replace(strHTML,Match.Value,"") 
Next

'取闭合的<> 
objRegExp.Pattern = "<script(.+?)/script>" 
'进行匹配 
Set Matches = objRegExp.Execute(strHTML) 
' 遍历匹配集合，并替换掉匹配的项目 
For Each Match in Matches 
strHtml=Replace(strHTML,Match.Value,"") 
Next

'取闭合的<> 
objRegExp.Pattern = "<.+?>" 
'进行匹配 
Set Matches = objRegExp.Execute(strHTML) 
' 遍历匹配集合，并替换掉匹配的项目 
For Each Match in Matches 
strHtml=Replace(strHTML,Match.Value,"") 
Next

RemoveHTML_ttkj=strHTML 
Set objRegExp = Nothing 
End Function 

'还原HTML代码-----------------------
Function HTMLDecode(reString)
	Dim Str:Str=reString
	If Not IsNull(Str) Then
		Str = Replace(Str, "&amp;", "&")
		Str = Replace(Str, "&gt;", ">")
		Str = Replace(Str, "&lt;", "<")
		Str = Replace(Str, "&nbsp;", CHR(32))
	    Str = Replace(Str, "&nbsp;", CHR(9))
		Str = Replace(Str, "&#160;&#160;&#160;&#160;", CHR(9))
		Str = Replace(Str, "&quot;", CHR(34))
		Str = Replace(Str, "&#39;", CHR(39))
		Str = Replace(Str, "", CHR(13))
		Str = Replace(Str, "<br>", CHR(10))
		HTMLDecode = Str
	End If
End Function


'过滤指定字符，参数strHTML：待过滤处理的HTML代码内容 参数strTAGs：为待过滤掉的HTML标记名，各标记名以英文逗号( , )为间隔，nType，用来区分过滤标签的类型（单标签还是双标签），当nType=1时，可以过滤单标记（如果：img,hr等等）调用形式：sPageCont=lFilterBadHTML(sPageCont,"script,iframe,object,table")

Function lFilterBadHTML(byval strHTML,byval strTAGs,byval nType)
Dim objRegExp,strOutput
Dim arrTAG,i
arrTAG=Split(strTAGs,",")
Set objRegExp = New Regexp 
strOutput=strHTML 
objRegExp.IgnoreCase = True
objRegExp.Global = True
If nType<>1 Then
For i=0 to UBound(arrTAG)
objRegExp.Pattern = "<"&arrTAG(i)&"[\s\S]+</"&arrTAG(i)&"*>"
strOutput = objRegExp.Replace(strOutput, "") 
Next
Else
For i=0 to UBound(arrTAG)
objRegExp.Pattern = "<"&arrTAG(i)&"[^>]+>"
strOutput = objRegExp.Replace(strOutput, "") 
Next
End If
Set objRegExp = Nothing
lFilterBadHTML = strOutput 
End Function

'替换关键词指定字符
function   Replace_Text(fString) 
if   isnull(fString)   then 
Replace_Text= " " 
exit   function 
else 
fString=trim(fString) 
fString=replace(fString, "，", ",") 
fString=replace(fString, "|", ",") 
fString=replace(fString, "｜", ",")
fString=server.htmlencode(fString) 
Replace_Text=fString 
end   if 
end   function


'**************************************************
'函数名：Check_Str
'作  用：判断字符串
'参  数：Input ---- 要处理的字符串'        
'返回值：包含符号返回True，否则返回False
'**************************************************
Function Check_Str(Input)
Dim Output
 Output = False   
 'If InStr(Input,Chr(32)) Then Output = True '空格    
 'If InStr(Input,Chr(33)) Then Output = True '!    
 If InStr(Input,Chr(34)) Then Output = True '"    
 If InStr(Input,Chr(35)) Then Output = True '#    
 If InStr(Input,Chr(36)) Then Output = True '$    
 'If InStr(Input,Chr(37)) Then Output = True '%    
 If InStr(Input,Chr(38)) Then Output = True '&    
 If InStr(Input,Chr(39)) Then Output = True ''    
 If InStr(Input,Chr(42)) Then Output = True '*    
 'If InStr(Input,Chr(43)) Then Output = True '+    
 If InStr(Input,Chr(44)) Then Output = True ',    
 'If InStr(Input,Chr(46)) Then Output = True '.    
 If InStr(Input,Chr(47)) Then Output = True '/    
 'If InStr(Input,Chr(58)) Then Output = True ':    
 'If InStr(Input,Chr(59)) Then Output = True ':    
 If InStr(Input,Chr(60)) Then Output = True '<    
 'If InStr(Input,Chr(61)) Then Output = True '=    
 If InStr(Input,Chr(62)) Then Output = True '>    
 If InStr(Input,Chr(63)) Then Output = True '?    
 If InStr(Input,Chr(64)) Then Output = True '@    
 'If InStr(Input,Chr(91)) Then Output = True '[    
 If InStr(Input,Chr(92)) Then Output = True '\    
 'If InStr(Input,Chr(93)) Then Output = True ']    
 If InStr(Input,Chr(94)) Then Output = True '^    
 If InStr(Input,Chr(96)) Then Output = True '`    
 If InStr(Input,Chr(123)) Then Output = True '{    
 'If InStr(Input,Chr(124)) Then Output = True '|    
 If InStr(Input,Chr(125)) Then Output = True '}    
 If InStr(Input,Chr(126)) Then Output = True '~    
 Check_Str = Output
End Function   

'**************************************************
'函数名：CheckIsEmpty
'作  用：用于判断字符串是否为空
'参  数：tstr ---要处理的字符串'             
'**************************************************
Function CheckIsEmpty(tstr)
CheckIsEmpty=false
If IsNull(tstr) or Tstr="" Then Exit Function
Dim Str,re 
Str=Tstr
Set re=new RegExp 
re.IgnoreCase =True 
re.Global=True
str= Replace(str, vbNewLine, "")
str = Replace(str, Chr(9), "")
str = Replace(str, " ", "")
str = Replace(str, "&nbsp;", "")
re.Pattern="<img(.[^>]*)>"
str =re.Replace(Str,"94kk") 
re.Pattern="<(.[^>]*)>" 
Str=re.Replace(Str,"")
Set Re=Nothing
If Str<>"" Then CheckIsEmpty=true 
End Function 


'检查无效字符----------------------
Function CheckStr(byVal ChkStr) 
	Dim Str:Str=ChkStr
	Str=Trim(Str)
	If IsNull(Str) Then
		CheckStr = ""
		Exit Function 
	End If
	Dim re
	Set re=new RegExp
	re.IgnoreCase =True
	re.Global=True
	re.Pattern="(\r\n){3,}"
	Str=re.Replace(Str,"$1$1$1")
	Set re=Nothing
	Str = Replace(Str,"'","''")
	Str = Replace(Str, "!!!", "!")
	Str = Replace(Str, "★★★", "★")
	CheckStr=Str
End Function

'**************************************************
'函数名：IsInteger
'作  用：检测是否有效的数字
'参  数：Para ---要处理的字符串'             
'**************************************************
Function IsInteger(Para) 
	IsInteger=False
	If Not (IsNull(Para) Or Trim(Para)="" Or Not IsNumeric(Para)) Then
		IsInteger=True
	End If
End Function

'**************************************************
'函数名：strint
'作  用：ID号转为字数
'参  数：id ---要处理的字符串'             
'**************************************************
Function strint(id)
If IsNumeric(id) = 0 or id = "" Then id = 0
strint = Clng(id)
End Function
'**************************************************
'函数名：SafeRequest
'作  用：ID号转为字数
'参  数：ParaName ---要处理的字符串  
        'ParaType ---要处理的类型
'**************************************************
 Function SafeRequest(ParaName,ParaType)
       Dim ParaValue
       ParaValue=Request(ParaName)
       If ParaType=1 then
              If not isNumeric(ParaValue) then
                     Response.write "<center>参数" & ParaName & "必须为数字型！</center>"
                     Response.end
              End if
       Else
              ParaValue=replace(ParaValue,"'","''")'替换字符
       End if
       SafeRequest=ParaValue
End Function
'**************************************************
'函数名：Html2Js
'作  用：将Html代码转为Js格式
'参  数：str ---要处理的字符串'             
'**************************************************
Private Function Html2Js(str)
    If str <> "" Then
        str = Replace(str, Chr(34), "\""")
        str = Replace(str, Chr(39), "\'")
        str = Replace(str, Chr(13), "\n")
        str = Replace(str, Chr(10), "\r")
    End If
    Html2Js = str
End Function

'**************************************************
'函数名：CheckStringLength
'作  用：测试字符串长度
'参  数：txt ---要处理的字符串'             
'**************************************************
Function CheckStringLength(txt)
Dim x,y,ii
txt=trim(txt)
x = len(txt)
y = 0
for ii = 1 to x
if asc(mid(txt,ii,1)) < 0 or asc(mid(txt,ii,1)) >255 then '如果是汉字
y = y + 2
else
y = y + 1
end if
next
CheckStringLength = y
End Function

'**************************************************
'函数名：CheckTheChar
'作  用：计算字符串中含指定字符的个数
'参  数：TheString待检测的字符串
'        TheChar指定的字符串'        
'**************************************************
Function CheckTheChar(TheString,TheChar)
Dim n
if inStr(TheString,TheChar) then
for n =1 to Len(TheString)
if Mid(TheString,n,Len(TheChar))=TheChar then 
CheckTheChar=CheckTheChar+1
End if
Next
CheckTheChar=CheckTheChar
else
CheckTheChar=0
end if
End Function

'截取标题过虑函数＝＝＝＝＝＝＝＝＝＝＝＝＝＝
Function ShowText(Text) 
	Text=replace(Text,"<br>","")
	Text=replace(Text,chr(32),"")'空格符
	Text=replace(Text,chr(13),"")'回车符
	Text=replace(Text,chr(10),"")'换行符
	Text=replace(Text,chr(9),"")'tab
	ShowText=replace(Text,"&nbsp;","")
End Function

'**************************************************
'函数名：CutStr
'作  用：截取字符串
'参  数：str ---要处理的字符串
'        StrLen ---- 截取长度
'        Other超过指定长度则显示的字符
'**************************************************
Function CutStr(Str,StrLen,Other)
  dim L,T,I,C
  if Str="" then
    CutStr=""
    exit function
  end if
  Str=Replace(Replace(Replace(Replace(Str,"&nbsp;"," "),"&quot;",Chr(34)),"&gt;",">"),"&lt;","<")
  L=Len(Str)
  T=0
  for i=1 to L
    C=Abs(AscW(Mid(Str,i,1)))
    if C>255 then
      T=T+2
    else
      T=T+1
    end if
    if T>=StrLen then
      CutStr=left(str,i)&Other
      exit for
    else
      CutStr=Str
    end if
  next
  'CutStr=Replace(Replace(Replace(replace(CutStr," ","&nbsp;"),Chr(34),"&quot;"),">","&gt;"),"<","&lt;")'过滤htm代码
  CutStr=CutStr
End Function

'**************************************************
'函数名：cut_string
'作  用：从前面截取字符串到第一个指定字符前
'参  数：string ---要处理的字符串
'        str ---- 指定字符
'**************************************************
Function cut_string(string,Str)
Dim StrLen,CutStr
If string<>"" then
StrLen=InStr(string,Str)-1
CutStr=Left(string,StrLen)
End if
cut_string=CutStr
End Function

'***************************************************
'检查组件是否已经安装
'***************************************************
Function IsObjInstalled(strClassString)
 On Error Resume Next
 IsObjInstalled = False
 Err = 0
 Dim xTestObj
 Set xTestObj = Server.CreateObject(strClassString)
 If 0 = Err Then IsObjInstalled = True
 Set xTestObj = Nothing
 Err = 0
End Function

'**************************************************
'函数名：IsExists
'作  用：判断数据库中的数据表的字段是否存在
'参  数：fieldName ---- 字段名称
'        tableName ---- 数据表名称
'返回值：如果改数据表存在改字段,则返回True,否则返回False
'**************************************************
Function IsExists(fieldName, tableName)
    On Error Resume Next
    IsExists = True
    CONN.execute ("select " & fieldName & " from " & tableName)

    If Err Then
        IsExists = False
    End If
    Err.Clear
End Function

'**************************************************
'函数名：PageSplit
'作  用：翻页类函数
'参  数：currentpage ----当前页数
'        totalpage----总共页数
'        pagename--使用本函数文件的文件名（）
'        后面的参数作其它用
'***************************************************
Function PageSplit(currentpage,totalpage,pagename,type_id,lx)
Dim sp,Pagestart,Pageend,strSplit,j
	if currentpage mod 10 = 0 then                            
	Sp = currentpage \ 10                            
	else                            
	Sp = currentpage \ 10 + 1                            
	end if                            
	Pagestart = (Sp-1)*10+1                            
	Pageend = Sp*10                            
	strSplit = "<a href="&pagename&"?pageid=1"&type_id&"&"&C_ID&"&leixing="&lx&"><font face=webdings title=第一页>9</font></a>&nbsp;"                            
	if Sp > 1 then strSplit = strSplit & "<a href="&pagename&"?pageid="&Pagestart-10&type_id&"&"&C_ID&"&leixing="&lx&"><font face=webdings title=前十页>7</font></a>&nbsp;"                            
	for j=PageStart to Pageend                            
	if j > totalpage then exit for                            
	if j <> currentpage then                            
	strSplit = strSplit & "<a href="&pagename&"?pageid="&j&type_id&"&"&C_ID&"&leixing="&lx&">["&j&"]</a>&nbsp;"			                            
	else                            
	strSplit = strSplit & "<font color=red>["&j&"]</font>&nbsp;"                            
	end if                            
	next                            
	if Sp*10  < totalpage then strSplit = strSplit & "<a href="&pagename&"?pageid="&Pagestart+10&type_id&"&"&C_ID&"&leixing="&lx&"><font face=webdings title=后十页>8</font></a>"                             
	strSplit = strSplit & "<a href="&pagename&"?pageid="&totalpage&type_id&"&"&C_ID&"&leixing="&lx&"><font face=webdings title=""最后一页"">:</font></a>"                            
	PageSplit = strSplit                            
End Function


'**************************************************
'函数名：gen_key
'作  用：产生随机数
'参  数：digits ----
'调用方式strrad=gen_key(5)
'***************************************************
Function gen_key(digits)
dim Output,num,char_array(50)
char_array(0) = "0"
char_array(1) = "1"
char_array(2) = "2"
char_array(3) = "3"
char_array(4) = "4"
char_array(5) = "5"
char_array(6) = "6"
char_array(7) = "7"
char_array(8) = "8"
char_array(9) = "9"
char_array(10) = "a"
char_array(11) = "b"
char_array(12) = "c"
char_array(13) = "d"
char_array(14) = "e"
char_array(15) = "f"
char_array(16) = "g"
char_array(17) = "h"
char_array(18) = "i"
char_array(19) = "j"
char_array(20) = "k"
char_array(21) = "l"
char_array(22) = "m"
char_array(23) = "n"
char_array(24) = "o"
char_array(25) = "p"
char_array(26) = "q"
char_array(27) = "r"
char_array(28) = "s"
char_array(29) = "t"
char_array(30) = "u"
char_array(31) = "v"
char_array(32) = "w"
char_array(33) = "x"
char_array(34) = "y"
char_array(35) = "z"
randomize
do while len(output) < digits
num = char_array(Int((35 - 0 + 1) * Rnd + 0))
output = output + num
loop
gen_key = output & year(now) & month(now) & day(now) & hour(now) & minute(now) & second(now)
End Function

'信息页生成参数
function rep1(aa)
	AA=replace(AA,"news.asp","../news.asp")
	AA=replace(AA,"list.asp","../list.asp")
	AA=replace(AA,"tt_getcode.asp","../tt_getcode.asp")
	AA=replace(AA,"tt_getcodee.asp","../tt_getcodee.asp")
	AA=replace(AA,"img/","../img/")
	AA=replace(AA,"images/","../images/")
	AA=replace(AA,"Include/","../Include/")
	AA=replace(AA,"skin/","../skin/")
    AA=replace(AA,"/../skin","/skin")
	AA=replace(AA,"addmembers.asp","../addmembers.asp")
	AA=replace(AA,"addtourists.asp","../addtourists.asp")
	AA=replace(AA,"Type.asp","../Type.asp")
	AA=replace(AA,"user_xxgl.asp","../user_xxgl.asp")
	AA=replace(AA,"sdlist.asp","../sdlist.asp")
	AA=replace(AA,"sd../list.asp","../sdlist.asp")
	AA=replace(AA,"search.asp","../search.asp")	
	AA=replace(AA,"company.asp","../company.asp")
	AA=replace(AA,"type_list.asp","../type_list.asp")
	AA=replace(AA,"type_../list.asp","../type_list.asp")
	AA=replace(AA,"city_list.asp","../city_list.asp")
	AA=replace(AA,"city_../list.asp","../city_list.asp")
	AA=replace(AA,"Message_board.asp","../Message_board.asp")
	AA=replace(AA,"help.asp","../help.asp")
	AA=replace(AA,"user.asp","../user.asp")
	AA=replace(AA,"login_save.asp","../login_save.asp")
	AA=replace(AA,"RetakePassword_a.asp","../RetakePassword_a.asp")
	AA=replace(AA,"user_regagree.asp","../user_regagree.asp")	
	AA=replace(AA,"user/xinfodt.asp","../user/xinfodt.asp")
	AA=replace(AA,"preview.asp","../preview.asp")
	AA=replace(AA,"user/com.asp","../user/com.asp")
	AA=replace(AA,"user/contribute.asp","../user/contribute.asp")	
	AA=replace(AA,"user/evaluation.asp","../user/evaluation.asp")	
	AA=replace(AA,"index.asp","../index.asp")
	AA=replace(AA,"/../index.asp","/index.asp")
	AA=replace(AA,"qy../index.asp","../qyindex.asp")
	AA=replace(AA,"sd../index.asp","../sdindex.asp")
	AA=replace(AA,"xx../index.asp","../xxindex.asp")
	AA=replace(AA,"user/Bad_info_list.asp","../user/Bad_info_list.asp")
	AA=replace(AA,"user/Bad_info_../list.asp","../user/Bad_info_list.asp")
	AA=replace(AA,"messh.asp","../messh.asp")
	AA=replace(AA,"user/collection.asp","../user/collection.asp")
	AA=replace(AA,"user/Bad_info.asp","../user/Bad_info.asp")	
	AA=replace(AA,"user/reply.asp","../user/reply.asp")
    AA=replace(AA,"user/user_mail.asp","../user/user_mail.asp")	
	AA=replace(AA,"login.asp","../login.asp")	
	AA=replace(AA,"user/rq_message.asp","../user/rq_message.asp")
	AA=replace(AA,"user/rq_messagecs.asp","../user/rq_messagecs.asp")
	AA=replace(AA,"ads/","../ads/")
	AA=replace(AA,"UploadFiles/infopic/","../UploadFiles/infopic/")
	AA=replace(AA,"UploadFiles/adfp/","../UploadFiles/adfp/")
	AA=replace(AA,"UploadFiles/logo/","../UploadFiles/logo/")
	AA=replace(AA,"cityadmin/../login.asp","../cityadmin/login.asp")
	AA=replace(AA,"vip../news.asp","../vipnews.asp")
	AA=replace(AA,"magazine/","../magazine/")	
	rep1=replace(AA,"/../images","/images")
End Function


'生成静态页函数开始＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝
Function JS_AlertEnd(str)'按提示返回页面
response.write(_
"<script language=javascript>alert("""& str &""");history.go(-1);</script>")
response.End()
End Function 


'**************************************************
'函数名：CD_GetCode
'作  用：获取网页源码
'参  数：URL ---- 网址
'        charset ---- 源码
'返回值：如果存在源码,则返回True,否则返回False
'**************************************************
Dim CD_ErrStr
CD_ErrStr = ""
Function CD_GetCode(URL, charset)
	On Error Resume Next
	Dim Http

	If IsNull(URL)=True Or URL="" Then
		CD_GetCode="False"
		CD_ErrStr="网址错误"
		Exit Function
	End If
	Set Http=server.createobject("MSXML2.XMLHTTP")
	Http.open "GET",URL,False
	Http.Send()
		If Http.Readystate<>4 then
			Set Http=Nothing 
			CD_GetCode="False"
			CD_ErrStr="获取不到源代码"
			Exit function
		End if
	CD_GetCode=CD_ToStr(Http.responseBody,charset)
	Set Http=Nothing

	If Err.number<>0 then
		If IsNull(URL)=True Or Len(URL)<18 Or URL="" Then
			CD_GetCode="False"
			CD_ErrStr="网址错误"
			Exit Function
		End If
		Set Http=server.createobject("MSXML2.XMLHTTP")
		On Error Resume Next
		Http.open "GET",URL,False
		Http.Send()
			If Http.Readystate<>4 or Http.Status > 300 then
				Set Http=Nothing 
				CD_GetCode="False"
				CD_ErrStr="获取不到源代码"
				Exit function
			End if
		CD_GetCode=CD_ToStr(Http.responseBody,charset)
		Set Http=Nothing
		Err.Clear
	End If

End Function

Function CD_ToStr(str,charset)
	Dim Objstream
	Set Objstream = Server.CreateObject("adodb.stream")
	objstream.Type = 1
	objstream.Mode =3
	objstream.Open
	objstream.Write str
	objstream.Position = 0
	objstream.Type = 2
	objstream.Charset = charset
	CD_ToStr = objstream.ReadText 
	objstream.Close
	set objstream = nothing
End Function

'**************************************************
'函数名：CD_WriteFile
'作  用：将替换后的内容写入HTML文档并写入文件
'参  数：content --为替换后的字符串
'        filename--为生成的文件名
'**************************************************
Function CD_WriteFile(content,filename)
	Dim fso,f
	Set fso = Server.CreateObject("Scripting.FileSystemObject") 
		Set f = fso.CreateTextFile(filename,true)'如果文件名重复将覆盖旧文件
			f.Write content
		f.Close
		Set f = Nothing
	set fso=Nothing
End Function
'生成静态页函数结束＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝

'*************************************************
'函数名：WriteToFile
'作  用：生成文件过程函数
'参  数：FileUrl--生成文件路径和名称,Str要生成的内容,CharSet生成文件编码方式
'*************************************************
Function WriteToFile(FileUrl,Str,CharSet)
Dim stm
On Error Resume Next
    Set stm = CreateObject("Adodb.Stream")
    stm.Type = 2
    stm.mode = 3
    stm.charset = CharSet
    stm.Open
    stm.WriteText Str
    stm.SaveToFile FileUrl, 2
    stm.flush
    stm.Close
    Set stm = Nothing
End Function



'**************************************************
'函数名：RemoveAllCache
'作  用：实现对站点的所有缓存Application的清理
'调  用：Call RemoveAllCache()        
'**************************************************
Sub RemoveAllCache()  
Dim cachelist,i  
Call InnerHtml("UpdateInfo","<b>开始执行清理当前站点缓存</b>：")  
Cachelist=split(GetallCache(),",")  
If UBound(cachelist)>1 Then  
For i=0 to UBound(cachelist)-1  
DelCahe Cachelist(i)  
Call InnerHtml("UpdateInfo","更新 <b>"&cachelist(i)&"</b> 完成")  
Next  
Call InnerHtml("UpdateInfo","更新了"& UBound(cachelist)-1 &"个缓存对象<br>")  
Else  
Call InnerHtml("UpdateInfo","<b>当前站点全部缓存清理完成。</b>。")  
End If  
End Sub  
Function GetallCache()  
Dim Cacheobj  
For Each Cacheobj in Application.Contents  
GetallCache = GetallCache & Cacheobj & ","  
Next  
End Function  
Sub DelCahe(MyCaheName)  
Application.Lock  
Application.Contents.Remove(MyCaheName)  
Application.unLock  
End Sub  
Sub InnerHtml(obj,msg)  
Response.Write "<li>"&msg&"</li>"  
Response.Flush  
End Sub

Dim onKeyUp'只允许输入数字,用法:<%=onKeyUp% >
onKeyUp="onKeyUp=""if(isNaN(value)){alert('只允许输入数字');value='';}""   onbeforepaste=""clipboardData.setData('text',clipboardData.getData('text').replace(/[^\d|^\.]/g,''))"""

'*************************************************
'函数名：FormatDate
'作  用：时间格式化
'参  数：DateAndTime --要处理的时间
'        Format-格式序号
'*************************************************
Function FormatDate(DateAndTime,Format)
  On Error Resume Next
  Dim yy,y, m, d, h, mi, s, strDateTime
  FormatDate = DateAndTime
  If Not IsNumeric(Format) Then Exit Function
  If Not IsDate(DateAndTime) Then Exit Function
  yy = CStr(Year(DateAndTime))
  y = Mid(CStr(Year(DateAndTime)),3)
  m = CStr(Month(DateAndTime))
  If Len(m) = 1 Then m = "0" & m
  d = CStr(Day(DateAndTime))
  If Len(d) = 1 Then d = "0" & d
  h = CStr(Hour(DateAndTime))
  If Len(h) = 1 Then h = "0" & h
  mi = CStr(Minute(DateAndTime))
  If Len(mi) = 1 Then mi = "0" & mi
  s = CStr(Second(DateAndTime))
  If Len(s) = 1 Then s = "0" & s
  Select Case Format
  Case "0"
  strDateTime =""
  Case "1"
    strDateTime = y & "-" & m & "-" & d & " " & h & ":" & mi & ":" & s
  Case "2"
    strDateTime = yy & m & d & h & mi & s
  Case "3"
    strDateTime = yy & m & d & h & mi
  Case "4"
    strDateTime = yy & "年" & m & "月" & d & "日"
  Case "5"
    strDateTime = m & "-" & d
  Case "6"
    strDateTime = m & "/" & d
  Case "7"
    strDateTime = m & "月" & d & "日"
  Case "8"
    strDateTime = y & "年" & m & "月"
  Case "9"
    strDateTime = y & "-" & m
  Case "10"
    strDateTime = y & "/" & m
  Case "11"
    strDateTime = y & "-" & m & "-" & d
  Case "12"
    strDateTime = y & "/" & m & "/" & d
  Case "13"
    strDateTime = yy & "." & m & "." & d
  Case "14"
    strDateTime = yy & "-" & m & "-" & d
  Case "15"
    strDateTime = DateAndTime
  Case "16"
    strDateTime = yy & m & d
  Case Else  
    strDateTime = DateAndTime
  End Select
  FormatDate = strDateTime
End Function 

'*************************************************
'函数名：tmraumen_Timer
'作  用：计算有效期剩余天时分秒
'参  数：MyDate --结束时间，格式2011-7-29 17:52:01
'*************************************************
Function tmraumen_Timer(MyDate)
     'MyDate 结束日期
     dim datesub '时间差
     dim dd '相差天数
     dim hh '相差小时数
     dim mm '相差分数
     dim ss '相差秒数
     dim strTip '标签提示     
     dim mytime    
     datesub=datediff("s",now,mydate)
     dd=fix(datesub/(60*60*24))
     hh=fix((datesub-dd*60*60*24)/(60*60))
     mm=fix((datesub-dd*60*60*24-hh*60*60)/60)
     ss=fix(datesub-dd*60*60*24-HH*60*60-MM*60)
	 strtip=strtip + cstr(DD) + ":"
     strtip=strtip + cstr(HH) + ":"
     strtip=strtip + cstr(MM) + ":"
     strtip=strtip + cstr(SS) + ""
	 tmraumen_Timer=strtip     
    End Function

'*************************************************
'函数名：ChkPost
'作  用：判断是不是站外提交
'参  数：
'将此代码放到第一行
'if not ChkPost() then 
'response.write "禁止站外提交或访问"
'response.end 
'end if
'*************************************************
Function ChkPost() 
dim HTTP_REFERER,SERVER_NAME 
dim server_v1,server_v2 
chkpost=false 
SERVER_NAME=CheckStr(Request.ServerVariables("SERVER_NAME")) 
HTTP_REFERER=CheckStr(Request.ServerVariables("HTTP_REFERER")) 
server_v1=Cstr(HTTP_REFERER) 
server_v2=Cstr(SERVER_NAME) 
if mid(server_v1,8,len(server_v2))<>server_v2 then 
   chkpost=false 
else 
   chkpost=true 
end if 
End Function

' ============================================
'函数名：DoDelFile
'作  用：通用删除图片文件函数
'参  数：sPathFile为要删除的文件路径及文件名
'============================================
Sub DoDelFile(sPathFile)
	On Error Resume Next
	Dim oFSO
	Set oFSO = Server.CreateObject("Scripting.FileSystemObject")
	if oFSO.fileExists(Server.MapPath(sPathFile)) then
	oFSO.DeleteFile(Server.MapPath(sPathFile))
	End if
	Set oFSO = Nothing
End Sub

' 得到安全字符串,在查询中或有必要强行替换的表单中使用
Function GetSafeStr(str)
	GetSafeStr = Replace(Replace(Replace(Trim(str), "'", ""), Chr(34), ""), ";", "")
End Function
Function ReplaceUnsafe(str)
  Dim badwords,i
badwords = "*|,|'|;|<|>|<=|=|<>|>=|(|)|/|%|^|+|-|&| " '使用 | 分割
badwords = split(badwords,"|",-1,1)
 For i=0 to Ubound(badwords)-1
  str = Replace(str,badwords(i),"") '所有敏感字符替换为空
 Next
 ReplaceUnsafe = str
End Function

'描述字段用
Function RemoveHTML(strHTML)
Dim objRegExp, Match, Matches
Set objRegExp = New Regexp
objRegExp.IgnoreCase = True
objRegExp.Global = True
'取闭合的<>
objRegExp.Pattern = "<.+?>"
'进行匹配
Set Matches = objRegExp.Execute(strHTML)
' 遍历匹配集合，并替换掉匹配的项目
For Each Match in Matches
strHtml=Replace(strHTML,Match.Value,"")
Next
RemoveHTML=strHTML
Set objRegExp = Nothing
End Function
' ============================================
'用户新闻专用结束
' ============================================
 
'==============根据IP自动切换地区开始========================
Class TQQWry
        Dim Country,LocalStr,Buf,OffSet
        Private StartIP,EndIP,CountryFlag
        Public FirstStartIP,LastStartIP,RecordCount,QQWryFile
        Private Stream,EndIPOff
        
        Private Sub Class_Initialize
                Country=""
                LocalStr=""
                StartIP=0
                EndIP=0
                CountryFlag=0 
                FirstStartIP=0 
                LastStartIP=0 
                EndIPOff=0 
                QQWryFile=Server.MapPath("../ip/cz.dat")
        End Sub

    Public Sub SetPath(p)
        QQWryFile = Server.MapPath(p)
    End Sub
        
        Function IP2Int(IP)
                Dim IPArray,i
                IPArray=Split(IP,".",-1)
                FOr i=0 to 3
                        If Not IsNumeric(IPArray(i)) Then IPArray(i)=0
                        If CInt(IPArray(i))<0 Then IPArray(i)=Abs(CInt(IPArray(i)))
                        If CInt(IPArray(i))>255 Then IPArray(i)=255
                Next
                IP2Int=(CInt(IPArray(0))*256*256*256)+(CInt(IPArray(1))*256*256)+(CInt(IPArray(2))*256)+CInt(IPArray(3))'-1
        End Function
        
        Function Int2IP(IntValue)
                p4=IntValue-Fix(IntValue/256)*256
                IntValue=(IntValue-p4)/256
                p3=IntValue-Fix(IntValue/256)*256
                IntValue=(IntValue-p3)/256
                p2=IntValue-Fix(IntValue/256)*256
                IntValue=(IntValue-p2)/256
                p1=IntValue
                Int2IP=Cstr(p1)&"."&Cstr(p2)&"."&Cstr(p3)&"."&Cstr(p4)
        End Function
        
        Private Function GetStartIP(RecNo)
                OffSet=FirstStartIP+RecNo * 7
                Stream.Position=OffSet
                Buf=Stream.Read(7)
                
                EndIPOff=AscB(MidB(Buf,5,1))+(AscB(MidB(Buf,6,1))*256)+(AscB(MidB(Buf,7,1))*256*256) 
                StartIP=AscB(MidB(Buf,1,1))+(AscB(MidB(Buf,2,1))*256)+(AscB(MidB(Buf,3,1))*256*256)+(AscB(MidB(Buf,4,1))*256*256*256)
                GetStartIP=StartIP
        End Function
        
        Private Function GetEndIP()
                Stream.Position=EndIPOff
                Buf=Stream.Read(5)
                EndIP=AscB(MidB(Buf,1,1))+(AscB(MidB(Buf,2,1))*256)+(AscB(MidB(Buf,3,1))*256*256)+(AscB(MidB(Buf,4,1))*256*256*256) 
                CountryFlag=AscB(MidB(Buf,5,1))
                GetEndIP=EndIP
        End Function
        
        Private Sub GetCountry(IP)
                If (CountryFlag=1 Or CountryFlag=2) Then
                        Country=GetFlagStr(EndIPOff+4)
                        If CountryFlag=1 Then
                                LocalStr=GetFlagStr(Stream.Position)
                                If IP>= IP2Int("255.255.255.0") And IP<=IP2Int("255.255.255.255") Then
                                        LocalStr=GetFlagStr(EndIPOff+21)
                                        Country=GetFlagStr(EndIPOff+12)
                                End If
                        Else
                                LocalStr=GetFlagStr(EndIPOff+8)
                        End If
                Else
                        Country=GetFlagStr(EndIPOff+4)
                        LocalStr=GetFlagStr(Stream.Position)
                End If
                Country=Trim(Country)
                LocalStr=Trim(LocalStr)
                If InStr(Country,"CZ88.NET") Then Country = "ip126.com"
                If InStr(LocalStr,"CZ88.NET") Then LocalStr = "ip126.com"
        End Sub
        
        Private Function GetFlagStr(OffSet)
                Dim Flag
                Flag=0
                Do While (True)
                        Stream.Position=OffSet
                        Flag=AscB(Stream.Read(1))
                        If(Flag=1 Or Flag=2 ) Then
                                Buf=Stream.Read(3) 
                                If (Flag=2 ) Then
                                        CountryFlag=2
                                        EndIPOff=OffSet-4
                                End If
                                OffSet=AscB(MidB(Buf,1,1))+(AscB(MidB(Buf,2,1))*256)+(AscB(MidB(Buf,3,1))*256*256)
                        Else
                                Exit Do
                        End If
                Loop
                
                If (OffSet<12 ) Then
                        GetFlagStr=""
                Else
                        Stream.Position=OffSet
                        GetFlagStr=GetStr() 
                End If
        End Function
        
        Private Function GetStr() 
                Dim c
                GetStr=""
                Do While (True)
                        c=AscB(Stream.Read(1))
                        If (c=0) Then Exit Do 
                        
                        If c>127 Then
                                If Stream.EOS Then Exit Do
                                GetStr=GetStr&Chr(AscW(ChrB(AscB(Stream.Read(1)))&ChrB(C)))
                        Else
                                GetStr=GetStr&Chr(c)
                        End If
                Loop 
        End Function
        
        Public Function QQWry(DotIP)
                Dim IP,nRet
                Dim RangB,RangE,RecNo                
                IP=IP2Int(DotIP)                
                Set Stream=CreateObject("ADodb.Stream")
                Stream.Mode=3
                Stream.Type=1
                Stream.Open
                Stream.LoadFromFile QQWryFile
                Stream.Position=0
                Buf=Stream.Read(8)                
                FirstStartIP=AscB(MidB(Buf,1,1))+(AscB(MidB(Buf,2,1))*256)+(AscB(MidB(Buf,3,1))*256*256)+(AscB(MidB(Buf,4,1))*256*256*256)
                LastStartIP=AscB(MidB(Buf,5,1))+(AscB(MidB(Buf,6,1))*256)+(AscB(MidB(Buf,7,1))*256*256)+(AscB(MidB(Buf,8,1))*256*256*256)
                RecordCount=Int((LastStartIP-FirstStartIP)/7)
                If (RecordCount<=1) Then
                        Country="Unknow"
                        QQWry=2
                        Exit Function
                End If                
                RangB=0
                RangE=RecordCount                
                Do While (RangB<(RangE-1)) 
                        RecNo=Int((RangB+RangE)/2) 
                        Call GetStartIP (RecNo)
                        If (IP=StartIP) Then
                                RangB=RecNo
                                Exit Do
                        End If
                        If (IP>StartIP) Then
                                RangB=RecNo
                        Else 
                                RangE=RecNo
                        End If
                Loop                
                Call GetStartIP(RangB)
                Call GetEndIP()
                If (StartIP<=IP) And ( EndIP>=IP) Then

                        nRet=0
                Else

                        nRet=3
                End If
                Call GetCountry(IP)
                QQWry=nRet
        End Function

        Private Sub Class_Terminate()
                On ErrOr Resume Next
                Stream.Close
                If Err Then Err.Clear
                Set Stream=Nothing
        End Sub  
End Class



Function Look_Ip(path,IP)
    Dim Wry, IPType, QQWryVersion, IpCounter
    Set Wry = New TQQWry
    Wry.SetPath path
    On Error Resume Next
    IPType = Wry.QQWry(IP)
    Look_Ip = Wry.Country' & " - " & Wry.LocalStr'Wry.LocalStr获取IP所属公司
    If Err Then
        Err.Clear
        Look_Ip = "查询出错"
    End If
End Function
'==============根据IP自动切换地区结束========================

'*************************************************
'函数名：PicAspJpeg
'作  用：给上传图片加文字或图片水印
'参  数：picname带路径的图片名称
'*************************************************
Sub PicAspJpeg(picname)
If SY_Bold=1 Then
SY_Bold=true
Else
SY_Bold=False
End if
SY_FontColor=right(SY_FontColor,6)
SY_FontColor="&H"&SY_FontColor&""
SY_BackgroundColor=right(SY_BackgroundColor,6)
SY_BackgroundColor="&H"&SY_BackgroundColor&""
Dim current_time
current_time=year(now) & right("0" & month(now),2) & right("0" & day(now),2) & right("0" & hour(now),2) & right("0" & minute(now),2)'格式当前时间 
	'文字水印
	On Error Resume Next
	Dim FSOpic
	Set FSOpic = Server.CreateObject("Scripting.FileSystemObject")
	if SY_Type=1 then		
		if fsopic.FileExists(Server.MapPath(picname)) Then
		If CCur(current_time)-CCur(Left(right(picname,len(picname)-InStrRev(picname,"/")),12))<SY_interval then'SY_interval分钟前旧图片不重复加水印
			Dim Jpeg                                                   '声明变量
			Set Jpeg = Server.CreateObject("Persits.Jpeg")             '调用组件
			Jpeg.Open Server.MapPath(picname)                          '源图片位置		
			Jpeg.Canvas.Font.Color = SY_FontColor          '水印字体颜色
			Jpeg.Canvas.Font.Family = SY_FontName          '水印字体		
			Jpeg.Canvas.Font.Size = SY_FontSize            '水印字体大小
			Jpeg.Canvas.Font.Bold = SY_Bold                '字体是否加粗
			'Jpeg.Canvas.Font.BkMode = "&H"&""&DrawShadowColor&""      '字体背景颜色，不启用为透明
			Jpeg.Canvas.Font.Quality = 3                               '文字清晰度
			Jpeg.Canvas.Print Jpeg.OriginalWidth-SY_X-50,Jpeg.OriginalHeight-30-SY_FontSize,SY_Text'水印文字,xy座标	
			Jpeg.Save Server.MapPath(picname) '//生成有水印的新图片及保存位置
			Set Jpeg = Nothing
		end If
		end if
	'图片水印
	elseif SY_Type=2 then
		If fsopic.FileExists(Server.MapPath(picname)) and fsopic.FileExists(Server.MapPath(""&SY_FileName&"")) then	'打开要加图片水印的图片
		If CCur(current_time)-CCur(Left(right(picname,len(picname)-InStrRev(picname,"/")),12))<SY_interval then'SY_interval分钟前旧图片不重复加水印
			Dim AspJpeg
			Set AspJpeg=Server.CreateObject("Persits.Jpeg")
			AspJpeg.Open Server.MapPath(picname)
			Dim T_AspJpeg,T_Left,T_Top
			Set T_AspJpeg=Server.CreateObject("Persits.Jpeg")
			T_AspJpeg.Open Server.MapPath(""&SY_FileName&"")'打开水印logo
			T_Left=GetPosition_X(AspJpeg.OriginalWidth,T_AspJpeg.Width,SY_X)
			T_Top=GetPosition_Y(AspJpeg.OriginalHeight,T_AspJpeg.Height,SY_Y)			
			AspJpeg.DrawImage T_Left,T_Top,T_AspJpeg,SY_Transparence,SY_BackgroundColor'参数分别为距左、距顶、水印图片、透明度、背景色			
			AspJpeg.Quality=100
			AspJpeg.Save Server.MapPath(picname)
			Set T_AspJpeg=Nothing
		End If
		End If
	end if
	Set FSOpic = Nothing
End Sub

'函数功能：获得图片的X坐标
Function GetPosition_X(ByVal t0,ByVal t1,ByVal t2)
    Select Case SY_coordinates
		Case 0:GetPosition_X=t2'左上
		Case 1:GetPosition_X=t2'左下
		Case 2:GetPosition_X=(t0-t1)/2'居中
		Case 3:GetPosition_X=t0-t1-t2'右上
		Case 4:GetPosition_X=t0-t1-t2'右下
		Case Else:GetPosition_X=0'不显示
	End Select
End Function

'函数功能：获得图片的Y坐标
Function GetPosition_Y(ByVal t0,ByVal t1,ByVal t2)
    Select Case SY_coordinates
		Case 0:GetPosition_Y=t2'左上
		Case 1:GetPosition_Y=t0-t1-t2'左下
		Case 2:GetPosition_Y=(t0-t1)/2'居中
		Case 3:GetPosition_Y=t2'右上
		Case 4:GetPosition_Y=t0-t1-t2'右下
		Case Else:GetPosition_Y=0'不显示
    End Select
End Function

'*************************************************
'函数名：content_pic
'作  用：搜索内容中的图片以|号分隔组成字符串
'参  数：content_pic(content),content为内容
         '得到的结果为|a1|a2...
'*************************************************
Function content_pic(content)
Dim regEx,Matches,Match,Strsimg,Bimg,Strsimg2,Bimg2
Set regEx = New RegExp '建立正则表达式。
regEx.Pattern = "(<img)(.[^<>]*)(src=)('|"&CHR(34)&"| )?(.[^'|\s|"&CHR(34)&"]*)(\.)(jpg|gif|png|bmp|jpeg)('|"&CHR(34)&"|\s|>)(.[^>]*)(>)" '设置模式。
regEx.IgnoreCase = True '设置是否区分字符大小写。
regEx.Global = True '设置全局可用性。
Set Matches = regEx.Execute(content) '执行内容图片搜索。
For Each Match in Matches '遍历匹配集合。 '输入图片地址
Bimg=Bimg&Match.SubMatches(4)&"."&Match.SubMatches(6)&"|"
Next
content_pic=Bimg
End Function

'*************************************************
'函数名：BuildSmallPic
'作  用：根据原图生成缩略图
'参  数： 
's_OriginalPath: 原图片路径及名 例:images/image1.gif 
's_BuildBasePath: 生成图片的基路径,不论是否以“/”结尾均可 例:images或images/ 
'n_MaxWidth: 生成图片最大宽度 
'n_MaxHeight: 生成图片最大高度
's_EndFlag缩略图文件名结尾标识，如_small，结果为image1_small.gif
'如果在前台显示的缩略图是 100*100,这里 n_MaxWidth=100,n_MaxHeight=100
'返回值：BuildSmallPic=s_BuildBasePath & s_BuildFileName'返回生成后的缩略图的路径及名 
'错误处理： 
'如果函数执行过程中出现错误,将返回错误代码,错误代码以 “Error”开头 
'    Error_01:创建AspJpeg组件失败,没有正确安装注册该组件 
'    Error_02:原图片不存在,检查s_OriginalPath参数传入值 
'    Error_03:缩略图存盘失败.可能原因:缩略图保存基地址不存在,检查s_OriginalPath参数传入值;对目录没有写权限;磁盘空间不足 
'    Error_Other:未知错误 
'*************************************************
Function BuildSmallPic(s_OriginalPath,s_BuildBasePath,n_MaxWidth,n_MaxHeight,s_EndFlag) 
    Err.Clear 
    On Error Resume Next      
    '检查组件是否已经注册 
    Dim AspJpeg 
    Set AspJpeg=Server.Createobject("Persits.Jpeg") 
    If Err.Number <> 0 Then 
        Err.Clear 
        BuildSmallPic="Error_01" 
        Exit Function 
    End If 
    '检查原图片是否存在 
    Dim s_MapOriginalPath 
    s_MapOriginalPath=Server.MapPath(s_OriginalPath) 
    AspJpeg.Open s_MapOriginalPath '打开原图片 
    If Err.Number <> 0 Then 
        Err.Clear 
        BuildSmallPic="Error_02" 
        Exit Function 
    End If 
    '按比例取得缩略图宽度和高度 
    Dim n_OriginalWidth,n_OriginalHeight '原图片宽度、高度 
    Dim n_BuildWidth,n_BuildHeight '缩略图宽度、高度 
    Dim div1,div2 
    Dim n1,n2 
    n_OriginalWidth=AspJpeg.Width 
    n_OriginalHeight=AspJpeg.Height 
    div1=n_OriginalWidth/n_OriginalHeight 
    div2=n_OriginalHeight/n_OriginalWidth 
    n1=0 
    n2=0 
    If n_OriginalWidth > n_MaxWidth Then 
        n1=n_OriginalWidth/n_MaxWidth 
    Else 
        n_BuildWidth=n_OriginalWidth 
    End If 
    If n_OriginalHeight > n_MaxHeight Then 
        n2=n_OriginalHeight/n_MaxHeight 
    Else 
        n_BuildHeight=n_OriginalHeight 
    End If 
    If n1 <> 0 Or n2 <> 0 Then 
        If n1 > n2 Then 
            n_BuildWidth=n_MaxWidth 
            n_BuildHeight=n_MaxWidth * div2 
        Else 
            n_BuildWidth=n_MaxHeight * div1 
            n_BuildHeight=n_MaxHeight 
        End If 
    End If 
    '指定宽度和高度生成 
    AspJpeg.Width=n_BuildWidth 
    AspJpeg.Height=n_BuildHeight 
     
    '--将缩略图存盘开始-- 
    Dim pos,s_OriginalFileName,s_OriginalFileExt '位置、原文件名、原文件扩展名 
    pos=InStrRev(s_OriginalPath,"/") + 1 
    s_OriginalFileName=Mid(s_OriginalPath,pos) 
    pos=InStrRev(s_OriginalFileName,".") 
    s_OriginalFileExt=Mid(s_OriginalFileName,pos) 
    Dim s_MapBuildBasePath,s_MapBuildPath,s_BuildFileName '缩略图绝对路径、缩略图文件名     
    If Right(s_BuildBasePath,1) <> "/" Then s_BuildBasePath=s_BuildBasePath & "/" 
    s_MapBuildBasePath=Server.MapPath(s_BuildBasePath)     
    s_BuildFileName=Replace(s_OriginalFileName,s_OriginalFileExt,"") & s_EndFlag & s_OriginalFileExt 
    s_MapBuildPath=s_MapBuildBasePath & "\" & s_BuildFileName 
     
    AspJpeg.Save s_MapBuildPath '保存 
    If Err.Number <> 0 Then 
        Err.Clear 
        BuildSmallPic="Error_03" 
        Exit Function 
    End If 
    '--将缩略图存盘结束-- 
    '注销实例 
    Set AspJpeg=Nothing 
    If Err.Number <> 0 Then 
        BuildSmallPic="Error_Other" 
        Err.Clear 
    End If 
    BuildSmallPic=s_BuildBasePath & s_BuildFileName 
End Function %>