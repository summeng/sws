<%
'=====================================
'���칩����Ϣ��վ����ϵͳ
'��������Ƽ������ҳ�Ʒ
'�ͷ�����http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>
<%'��ȡ������ַ����http://www.ip126.com/list.asp�е�www.ip126.com
Dim BaseUrl
BaseUrl = LCase(Replace(Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("URL"),Split(request.ServerVariables("SCRIPT_NAME"),"/")(ubound(Split(request.ServerVariables("SCRIPT_NAME"),"/"))),""))

'**************************************************
'��������GetUrl
'��  �ã��õ���ǰҳ��ĵ�ַ�������80�˿�Ҳ����� 
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
'��������Alert
'��  �ã�����������ʾ������ȥ�� 
'�÷���Call Alert("��ʾ����","-1")
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


'�����ַ����ˡ��滻��һ
Function HTMLEncode(fString)
If Not IsNull(fString) Then
fString = trim(fString)
'fString = replace(fString, ";", "��")      '�ֺŹ���
fString = replace(fString, "--", "����") '--����
fString = replace(fString, "%20", "")     '�����ַ�����
fString = replace(fString, "==", "")      '==����
fString = replace(fString, ">", "&gt;")
fString = replace(fString, "<", "&lt;")
fString = replace(fString, "%", "&#37;")     '%����
fString = Replace(fString, CHR(32), " ")    '&nbsp;
fString = Replace(fString, CHR(9), " ")     '&nbsp;
fString = Replace(fString, CHR(34), "&quot;")
fString = Replace(fString, CHR(39), "&#39;") '�����Ź���
fString = Replace(fString, CHR(13), "")
fString = Replace(fString, CHR(10) & CHR(10), "</P><P> ")
fString = Replace(fString, CHR(10), "<BR> ")
HTMLEncode = fString
End If
End Function

'�����ַ����ˡ��滻����
Function HtmlEncode2(Content)
IF content="" or isnull(content) Then exit function
  Content = trim(Content)
  Content = Replace(Content,"%20"  , ""       )'�����ַ�����
  Content = Replace(Content,chr(62),  "��"  )' > 
  Content = Replace(Content,chr(60),  "��"  )' < 
  Content = Replace(Content,chr(39),  "��"    )' ' 
  Content = Replace(Content,chr(37),  "��"    )' % 
  Content = Replace(Content, vbcrlf,  ""      )
  Content = Replace(Content,chr(34),  "��")' "
  Content = Replace(Content,chr(40),  "��"    )' ( 
  Content = Replace(Content,chr(41),  "��"    )' ) 
  Content = Replace(Content,chr(91),  "��"    )' [ 
  Content = Replace(Content,chr(93),  "��"    )' ] 
  Content = Replace(Content,chr(123), "��"    )' { 
  Content = Replace(Content,chr(125), "��"    )' } 
  Content = Replace(Content, CHR(13),   "")   
  Content = Replace(Content,CHR(10), "")
  HtmlEncode2 = content 
End Function

'**************************************************
'��������chkdick
'��  �ã������ַ�����
'��  ����char ---Ҫ������ַ���'             
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
if instr(char,"�t") then
call errdick(13)
response.end 
end if
'if instr(char,"��") then
'call errdick(13)
'response.end 
'end if
if instr(char,"�r") then
call errdick(13)
response.end 
end if
'if instr(char,"��") then
'call errdick(13)
'response.end 
'end if
'if instr(char,"��") then
'call errdick(13)
'response.end 
'end if
'if instr(char,"��") then
'call errdick(13)
'response.end 
'end if
'if instr(char,"��") then
'call errdick(13)
'response.end 
'end if
'if instr(char,"��") then
'call errdick(13)
'response.end 
'end if
'if instr(char,"��") then
'call errdick(13)
'response.end 
'end if
'if instr(char,"��") then
'call errdick(13)
'response.end 
'end if
'if instr(char,"��") then
'call errdick(13)
'response.end 
'end if
end function


'*************************************************
'��������RemoveHTML_ttkj
'��  �ã�����HTML����ĺ�����������CSS��JS��ֻ��ʾ����
'��  ����RemoveHTML_ttkj()
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

'ȡ�պϵ�<> 
objRegExp.Pattern = "<style(.+?)/style>" 
'����ƥ�� 
Set Matches = objRegExp.Execute(strHTML) 
' ����ƥ�伯�ϣ����滻��ƥ�����Ŀ 
For Each Match in Matches 
strHtml=Replace(strHTML,Match.Value,"") 
Next

'ȡ�պϵ�<> 
objRegExp.Pattern = "<script(.+?)/script>" 
'����ƥ�� 
Set Matches = objRegExp.Execute(strHTML) 
' ����ƥ�伯�ϣ����滻��ƥ�����Ŀ 
For Each Match in Matches 
strHtml=Replace(strHTML,Match.Value,"") 
Next

'ȡ�պϵ�<> 
objRegExp.Pattern = "<.+?>" 
'����ƥ�� 
Set Matches = objRegExp.Execute(strHTML) 
' ����ƥ�伯�ϣ����滻��ƥ�����Ŀ 
For Each Match in Matches 
strHtml=Replace(strHTML,Match.Value,"") 
Next

RemoveHTML_ttkj=strHTML 
Set objRegExp = Nothing 
End Function 

'��ԭHTML����-----------------------
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


'����ָ���ַ�������strHTML�������˴����HTML�������� ����strTAGs��Ϊ�����˵���HTML����������������Ӣ�Ķ���( , )Ϊ�����nType���������ֹ��˱�ǩ�����ͣ�����ǩ����˫��ǩ������nType=1ʱ�����Թ��˵���ǣ������img,hr�ȵȣ�������ʽ��sPageCont=lFilterBadHTML(sPageCont,"script,iframe,object,table")

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

'�滻�ؼ���ָ���ַ�
function   Replace_Text(fString) 
if   isnull(fString)   then 
Replace_Text= " " 
exit   function 
else 
fString=trim(fString) 
fString=replace(fString, "��", ",") 
fString=replace(fString, "|", ",") 
fString=replace(fString, "��", ",")
fString=server.htmlencode(fString) 
Replace_Text=fString 
end   if 
end   function


'**************************************************
'��������Check_Str
'��  �ã��ж��ַ���
'��  ����Input ---- Ҫ������ַ���'        
'����ֵ���������ŷ���True�����򷵻�False
'**************************************************
Function Check_Str(Input)
Dim Output
 Output = False   
 'If InStr(Input,Chr(32)) Then Output = True '�ո�    
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
'��������CheckIsEmpty
'��  �ã������ж��ַ����Ƿ�Ϊ��
'��  ����tstr ---Ҫ������ַ���'             
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


'�����Ч�ַ�----------------------
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
	Str = Replace(Str, "����", "��")
	CheckStr=Str
End Function

'**************************************************
'��������IsInteger
'��  �ã�����Ƿ���Ч������
'��  ����Para ---Ҫ������ַ���'             
'**************************************************
Function IsInteger(Para) 
	IsInteger=False
	If Not (IsNull(Para) Or Trim(Para)="" Or Not IsNumeric(Para)) Then
		IsInteger=True
	End If
End Function

'**************************************************
'��������strint
'��  �ã�ID��תΪ����
'��  ����id ---Ҫ������ַ���'             
'**************************************************
Function strint(id)
If IsNumeric(id) = 0 or id = "" Then id = 0
strint = Clng(id)
End Function
'**************************************************
'��������SafeRequest
'��  �ã�ID��תΪ����
'��  ����ParaName ---Ҫ������ַ���  
        'ParaType ---Ҫ���������
'**************************************************
 Function SafeRequest(ParaName,ParaType)
       Dim ParaValue
       ParaValue=Request(ParaName)
       If ParaType=1 then
              If not isNumeric(ParaValue) then
                     Response.write "<center>����" & ParaName & "����Ϊ�����ͣ�</center>"
                     Response.end
              End if
       Else
              ParaValue=replace(ParaValue,"'","''")'�滻�ַ�
       End if
       SafeRequest=ParaValue
End Function
'**************************************************
'��������Html2Js
'��  �ã���Html����תΪJs��ʽ
'��  ����str ---Ҫ������ַ���'             
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
'��������CheckStringLength
'��  �ã������ַ�������
'��  ����txt ---Ҫ������ַ���'             
'**************************************************
Function CheckStringLength(txt)
Dim x,y,ii
txt=trim(txt)
x = len(txt)
y = 0
for ii = 1 to x
if asc(mid(txt,ii,1)) < 0 or asc(mid(txt,ii,1)) >255 then '����Ǻ���
y = y + 2
else
y = y + 1
end if
next
CheckStringLength = y
End Function

'**************************************************
'��������CheckTheChar
'��  �ã������ַ����к�ָ���ַ��ĸ���
'��  ����TheString�������ַ���
'        TheCharָ�����ַ���'        
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

'��ȡ������Ǻ�������������������������������
Function ShowText(Text) 
	Text=replace(Text,"<br>","")
	Text=replace(Text,chr(32),"")'�ո��
	Text=replace(Text,chr(13),"")'�س���
	Text=replace(Text,chr(10),"")'���з�
	Text=replace(Text,chr(9),"")'tab
	ShowText=replace(Text,"&nbsp;","")
End Function

'**************************************************
'��������CutStr
'��  �ã���ȡ�ַ���
'��  ����str ---Ҫ������ַ���
'        StrLen ---- ��ȡ����
'        Other����ָ����������ʾ���ַ�
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
  'CutStr=Replace(Replace(Replace(replace(CutStr," ","&nbsp;"),Chr(34),"&quot;"),">","&gt;"),"<","&lt;")'����htm����
  CutStr=CutStr
End Function

'**************************************************
'��������cut_string
'��  �ã���ǰ���ȡ�ַ�������һ��ָ���ַ�ǰ
'��  ����string ---Ҫ������ַ���
'        str ---- ָ���ַ�
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
'�������Ƿ��Ѿ���װ
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
'��������IsExists
'��  �ã��ж����ݿ��е����ݱ���ֶ��Ƿ����
'��  ����fieldName ---- �ֶ�����
'        tableName ---- ���ݱ�����
'����ֵ����������ݱ���ڸ��ֶ�,�򷵻�True,���򷵻�False
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
'��������PageSplit
'��  �ã���ҳ�ຯ��
'��  ����currentpage ----��ǰҳ��
'        totalpage----�ܹ�ҳ��
'        pagename--ʹ�ñ������ļ����ļ�������
'        ����Ĳ�����������
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
	strSplit = "<a href="&pagename&"?pageid=1"&type_id&"&"&C_ID&"&leixing="&lx&"><font face=webdings title=��һҳ>9</font></a>&nbsp;"                            
	if Sp > 1 then strSplit = strSplit & "<a href="&pagename&"?pageid="&Pagestart-10&type_id&"&"&C_ID&"&leixing="&lx&"><font face=webdings title=ǰʮҳ>7</font></a>&nbsp;"                            
	for j=PageStart to Pageend                            
	if j > totalpage then exit for                            
	if j <> currentpage then                            
	strSplit = strSplit & "<a href="&pagename&"?pageid="&j&type_id&"&"&C_ID&"&leixing="&lx&">["&j&"]</a>&nbsp;"			                            
	else                            
	strSplit = strSplit & "<font color=red>["&j&"]</font>&nbsp;"                            
	end if                            
	next                            
	if Sp*10  < totalpage then strSplit = strSplit & "<a href="&pagename&"?pageid="&Pagestart+10&type_id&"&"&C_ID&"&leixing="&lx&"><font face=webdings title=��ʮҳ>8</font></a>"                             
	strSplit = strSplit & "<a href="&pagename&"?pageid="&totalpage&type_id&"&"&C_ID&"&leixing="&lx&"><font face=webdings title=""���һҳ"">:</font></a>"                            
	PageSplit = strSplit                            
End Function


'**************************************************
'��������gen_key
'��  �ã����������
'��  ����digits ----
'���÷�ʽstrrad=gen_key(5)
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

'��Ϣҳ���ɲ���
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


'���ɾ�̬ҳ������ʼ��������������������������������������������������������������������
Function JS_AlertEnd(str)'����ʾ����ҳ��
response.write(_
"<script language=javascript>alert("""& str &""");history.go(-1);</script>")
response.End()
End Function 


'**************************************************
'��������CD_GetCode
'��  �ã���ȡ��ҳԴ��
'��  ����URL ---- ��ַ
'        charset ---- Դ��
'����ֵ���������Դ��,�򷵻�True,���򷵻�False
'**************************************************
Dim CD_ErrStr
CD_ErrStr = ""
Function CD_GetCode(URL, charset)
	On Error Resume Next
	Dim Http

	If IsNull(URL)=True Or URL="" Then
		CD_GetCode="False"
		CD_ErrStr="��ַ����"
		Exit Function
	End If
	Set Http=server.createobject("MSXML2.XMLHTTP")
	Http.open "GET",URL,False
	Http.Send()
		If Http.Readystate<>4 then
			Set Http=Nothing 
			CD_GetCode="False"
			CD_ErrStr="��ȡ����Դ����"
			Exit function
		End if
	CD_GetCode=CD_ToStr(Http.responseBody,charset)
	Set Http=Nothing

	If Err.number<>0 then
		If IsNull(URL)=True Or Len(URL)<18 Or URL="" Then
			CD_GetCode="False"
			CD_ErrStr="��ַ����"
			Exit Function
		End If
		Set Http=server.createobject("MSXML2.XMLHTTP")
		On Error Resume Next
		Http.open "GET",URL,False
		Http.Send()
			If Http.Readystate<>4 or Http.Status > 300 then
				Set Http=Nothing 
				CD_GetCode="False"
				CD_ErrStr="��ȡ����Դ����"
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
'��������CD_WriteFile
'��  �ã����滻�������д��HTML�ĵ���д���ļ�
'��  ����content --Ϊ�滻����ַ���
'        filename--Ϊ���ɵ��ļ���
'**************************************************
Function CD_WriteFile(content,filename)
	Dim fso,f
	Set fso = Server.CreateObject("Scripting.FileSystemObject") 
		Set f = fso.CreateTextFile(filename,true)'����ļ����ظ������Ǿ��ļ�
			f.Write content
		f.Close
		Set f = Nothing
	set fso=Nothing
End Function
'���ɾ�̬ҳ����������������������������������������������������������������������������

'*************************************************
'��������WriteToFile
'��  �ã������ļ����̺���
'��  ����FileUrl--�����ļ�·��������,StrҪ���ɵ�����,CharSet�����ļ����뷽ʽ
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
'��������RemoveAllCache
'��  �ã�ʵ�ֶ�վ������л���Application������
'��  �ã�Call RemoveAllCache()        
'**************************************************
Sub RemoveAllCache()  
Dim cachelist,i  
Call InnerHtml("UpdateInfo","<b>��ʼִ������ǰվ�㻺��</b>��")  
Cachelist=split(GetallCache(),",")  
If UBound(cachelist)>1 Then  
For i=0 to UBound(cachelist)-1  
DelCahe Cachelist(i)  
Call InnerHtml("UpdateInfo","���� <b>"&cachelist(i)&"</b> ���")  
Next  
Call InnerHtml("UpdateInfo","������"& UBound(cachelist)-1 &"���������<br>")  
Else  
Call InnerHtml("UpdateInfo","<b>��ǰվ��ȫ������������ɡ�</b>��")  
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

Dim onKeyUp'ֻ������������,�÷�:<%=onKeyUp% >
onKeyUp="onKeyUp=""if(isNaN(value)){alert('ֻ������������');value='';}""   onbeforepaste=""clipboardData.setData('text',clipboardData.getData('text').replace(/[^\d|^\.]/g,''))"""

'*************************************************
'��������FormatDate
'��  �ã�ʱ���ʽ��
'��  ����DateAndTime --Ҫ�����ʱ��
'        Format-��ʽ���
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
    strDateTime = yy & "��" & m & "��" & d & "��"
  Case "5"
    strDateTime = m & "-" & d
  Case "6"
    strDateTime = m & "/" & d
  Case "7"
    strDateTime = m & "��" & d & "��"
  Case "8"
    strDateTime = y & "��" & m & "��"
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
'��������tmraumen_Timer
'��  �ã�������Ч��ʣ����ʱ����
'��  ����MyDate --����ʱ�䣬��ʽ2011-7-29 17:52:01
'*************************************************
Function tmraumen_Timer(MyDate)
     'MyDate ��������
     dim datesub 'ʱ���
     dim dd '�������
     dim hh '���Сʱ��
     dim mm '������
     dim ss '�������
     dim strTip '��ǩ��ʾ     
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
'��������ChkPost
'��  �ã��ж��ǲ���վ���ύ
'��  ����
'���˴���ŵ���һ��
'if not ChkPost() then 
'response.write "��ֹվ���ύ�����"
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
'��������DoDelFile
'��  �ã�ͨ��ɾ��ͼƬ�ļ�����
'��  ����sPathFileΪҪɾ�����ļ�·�����ļ���
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

' �õ���ȫ�ַ���,�ڲ�ѯ�л��б�Ҫǿ���滻�ı���ʹ��
Function GetSafeStr(str)
	GetSafeStr = Replace(Replace(Replace(Trim(str), "'", ""), Chr(34), ""), ";", "")
End Function
Function ReplaceUnsafe(str)
  Dim badwords,i
badwords = "*|,|'|;|<|>|<=|=|<>|>=|(|)|/|%|^|+|-|&| " 'ʹ�� | �ָ�
badwords = split(badwords,"|",-1,1)
 For i=0 to Ubound(badwords)-1
  str = Replace(str,badwords(i),"") '���������ַ��滻Ϊ��
 Next
 ReplaceUnsafe = str
End Function

'�����ֶ���
Function RemoveHTML(strHTML)
Dim objRegExp, Match, Matches
Set objRegExp = New Regexp
objRegExp.IgnoreCase = True
objRegExp.Global = True
'ȡ�պϵ�<>
objRegExp.Pattern = "<.+?>"
'����ƥ��
Set Matches = objRegExp.Execute(strHTML)
' ����ƥ�伯�ϣ����滻��ƥ�����Ŀ
For Each Match in Matches
strHtml=Replace(strHTML,Match.Value,"")
Next
RemoveHTML=strHTML
Set objRegExp = Nothing
End Function
' ============================================
'�û�����ר�ý���
' ============================================
 
'==============����IP�Զ��л�������ʼ========================
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
    Look_Ip = Wry.Country' & " - " & Wry.LocalStr'Wry.LocalStr��ȡIP������˾
    If Err Then
        Err.Clear
        Look_Ip = "��ѯ����"
    End If
End Function
'==============����IP�Զ��л���������========================

'*************************************************
'��������PicAspJpeg
'��  �ã����ϴ�ͼƬ�����ֻ�ͼƬˮӡ
'��  ����picname��·����ͼƬ����
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
current_time=year(now) & right("0" & month(now),2) & right("0" & day(now),2) & right("0" & hour(now),2) & right("0" & minute(now),2)'��ʽ��ǰʱ�� 
	'����ˮӡ
	On Error Resume Next
	Dim FSOpic
	Set FSOpic = Server.CreateObject("Scripting.FileSystemObject")
	if SY_Type=1 then		
		if fsopic.FileExists(Server.MapPath(picname)) Then
		If CCur(current_time)-CCur(Left(right(picname,len(picname)-InStrRev(picname,"/")),12))<SY_interval then'SY_interval����ǰ��ͼƬ���ظ���ˮӡ
			Dim Jpeg                                                   '��������
			Set Jpeg = Server.CreateObject("Persits.Jpeg")             '�������
			Jpeg.Open Server.MapPath(picname)                          'ԴͼƬλ��		
			Jpeg.Canvas.Font.Color = SY_FontColor          'ˮӡ������ɫ
			Jpeg.Canvas.Font.Family = SY_FontName          'ˮӡ����		
			Jpeg.Canvas.Font.Size = SY_FontSize            'ˮӡ�����С
			Jpeg.Canvas.Font.Bold = SY_Bold                '�����Ƿ�Ӵ�
			'Jpeg.Canvas.Font.BkMode = "&H"&""&DrawShadowColor&""      '���屳����ɫ��������Ϊ͸��
			Jpeg.Canvas.Font.Quality = 3                               '����������
			Jpeg.Canvas.Print Jpeg.OriginalWidth-SY_X-50,Jpeg.OriginalHeight-30-SY_FontSize,SY_Text'ˮӡ����,xy����	
			Jpeg.Save Server.MapPath(picname) '//������ˮӡ����ͼƬ������λ��
			Set Jpeg = Nothing
		end If
		end if
	'ͼƬˮӡ
	elseif SY_Type=2 then
		If fsopic.FileExists(Server.MapPath(picname)) and fsopic.FileExists(Server.MapPath(""&SY_FileName&"")) then	'��Ҫ��ͼƬˮӡ��ͼƬ
		If CCur(current_time)-CCur(Left(right(picname,len(picname)-InStrRev(picname,"/")),12))<SY_interval then'SY_interval����ǰ��ͼƬ���ظ���ˮӡ
			Dim AspJpeg
			Set AspJpeg=Server.CreateObject("Persits.Jpeg")
			AspJpeg.Open Server.MapPath(picname)
			Dim T_AspJpeg,T_Left,T_Top
			Set T_AspJpeg=Server.CreateObject("Persits.Jpeg")
			T_AspJpeg.Open Server.MapPath(""&SY_FileName&"")'��ˮӡlogo
			T_Left=GetPosition_X(AspJpeg.OriginalWidth,T_AspJpeg.Width,SY_X)
			T_Top=GetPosition_Y(AspJpeg.OriginalHeight,T_AspJpeg.Height,SY_Y)			
			AspJpeg.DrawImage T_Left,T_Top,T_AspJpeg,SY_Transparence,SY_BackgroundColor'�����ֱ�Ϊ���󡢾ඥ��ˮӡͼƬ��͸���ȡ�����ɫ			
			AspJpeg.Quality=100
			AspJpeg.Save Server.MapPath(picname)
			Set T_AspJpeg=Nothing
		End If
		End If
	end if
	Set FSOpic = Nothing
End Sub

'�������ܣ����ͼƬ��X����
Function GetPosition_X(ByVal t0,ByVal t1,ByVal t2)
    Select Case SY_coordinates
		Case 0:GetPosition_X=t2'����
		Case 1:GetPosition_X=t2'����
		Case 2:GetPosition_X=(t0-t1)/2'����
		Case 3:GetPosition_X=t0-t1-t2'����
		Case 4:GetPosition_X=t0-t1-t2'����
		Case Else:GetPosition_X=0'����ʾ
	End Select
End Function

'�������ܣ����ͼƬ��Y����
Function GetPosition_Y(ByVal t0,ByVal t1,ByVal t2)
    Select Case SY_coordinates
		Case 0:GetPosition_Y=t2'����
		Case 1:GetPosition_Y=t0-t1-t2'����
		Case 2:GetPosition_Y=(t0-t1)/2'����
		Case 3:GetPosition_Y=t2'����
		Case 4:GetPosition_Y=t0-t1-t2'����
		Case Else:GetPosition_Y=0'����ʾ
    End Select
End Function

'*************************************************
'��������content_pic
'��  �ã����������е�ͼƬ��|�ŷָ�����ַ���
'��  ����content_pic(content),contentΪ����
         '�õ��Ľ��Ϊ|a1|a2...
'*************************************************
Function content_pic(content)
Dim regEx,Matches,Match,Strsimg,Bimg,Strsimg2,Bimg2
Set regEx = New RegExp '����������ʽ��
regEx.Pattern = "(<img)(.[^<>]*)(src=)('|"&CHR(34)&"| )?(.[^'|\s|"&CHR(34)&"]*)(\.)(jpg|gif|png|bmp|jpeg)('|"&CHR(34)&"|\s|>)(.[^>]*)(>)" '����ģʽ��
regEx.IgnoreCase = True '�����Ƿ������ַ���Сд��
regEx.Global = True '����ȫ�ֿ����ԡ�
Set Matches = regEx.Execute(content) 'ִ������ͼƬ������
For Each Match in Matches '����ƥ�伯�ϡ� '����ͼƬ��ַ
Bimg=Bimg&Match.SubMatches(4)&"."&Match.SubMatches(6)&"|"
Next
content_pic=Bimg
End Function

'*************************************************
'��������BuildSmallPic
'��  �ã�����ԭͼ��������ͼ
'��  ���� 
's_OriginalPath: ԭͼƬ·������ ��:images/image1.gif 
's_BuildBasePath: ����ͼƬ�Ļ�·��,�����Ƿ��ԡ�/����β���� ��:images��images/ 
'n_MaxWidth: ����ͼƬ����� 
'n_MaxHeight: ����ͼƬ���߶�
's_EndFlag����ͼ�ļ�����β��ʶ����_small�����Ϊimage1_small.gif
'�����ǰ̨��ʾ������ͼ�� 100*100,���� n_MaxWidth=100,n_MaxHeight=100
'����ֵ��BuildSmallPic=s_BuildBasePath & s_BuildFileName'�������ɺ������ͼ��·������ 
'������ 
'�������ִ�й����г��ִ���,�����ش������,��������� ��Error����ͷ 
'    Error_01:����AspJpeg���ʧ��,û����ȷ��װע������ 
'    Error_02:ԭͼƬ������,���s_OriginalPath��������ֵ 
'    Error_03:����ͼ����ʧ��.����ԭ��:����ͼ�������ַ������,���s_OriginalPath��������ֵ;��Ŀ¼û��дȨ��;���̿ռ䲻�� 
'    Error_Other:δ֪���� 
'*************************************************
Function BuildSmallPic(s_OriginalPath,s_BuildBasePath,n_MaxWidth,n_MaxHeight,s_EndFlag) 
    Err.Clear 
    On Error Resume Next      
    '�������Ƿ��Ѿ�ע�� 
    Dim AspJpeg 
    Set AspJpeg=Server.Createobject("Persits.Jpeg") 
    If Err.Number <> 0 Then 
        Err.Clear 
        BuildSmallPic="Error_01" 
        Exit Function 
    End If 
    '���ԭͼƬ�Ƿ���� 
    Dim s_MapOriginalPath 
    s_MapOriginalPath=Server.MapPath(s_OriginalPath) 
    AspJpeg.Open s_MapOriginalPath '��ԭͼƬ 
    If Err.Number <> 0 Then 
        Err.Clear 
        BuildSmallPic="Error_02" 
        Exit Function 
    End If 
    '������ȡ������ͼ��Ⱥ͸߶� 
    Dim n_OriginalWidth,n_OriginalHeight 'ԭͼƬ��ȡ��߶� 
    Dim n_BuildWidth,n_BuildHeight '����ͼ��ȡ��߶� 
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
    'ָ����Ⱥ͸߶����� 
    AspJpeg.Width=n_BuildWidth 
    AspJpeg.Height=n_BuildHeight 
     
    '--������ͼ���̿�ʼ-- 
    Dim pos,s_OriginalFileName,s_OriginalFileExt 'λ�á�ԭ�ļ�����ԭ�ļ���չ�� 
    pos=InStrRev(s_OriginalPath,"/") + 1 
    s_OriginalFileName=Mid(s_OriginalPath,pos) 
    pos=InStrRev(s_OriginalFileName,".") 
    s_OriginalFileExt=Mid(s_OriginalFileName,pos) 
    Dim s_MapBuildBasePath,s_MapBuildPath,s_BuildFileName '����ͼ����·��������ͼ�ļ���     
    If Right(s_BuildBasePath,1) <> "/" Then s_BuildBasePath=s_BuildBasePath & "/" 
    s_MapBuildBasePath=Server.MapPath(s_BuildBasePath)     
    s_BuildFileName=Replace(s_OriginalFileName,s_OriginalFileExt,"") & s_EndFlag & s_OriginalFileExt 
    s_MapBuildPath=s_MapBuildBasePath & "\" & s_BuildFileName 
     
    AspJpeg.Save s_MapBuildPath '���� 
    If Err.Number <> 0 Then 
        Err.Clear 
        BuildSmallPic="Error_03" 
        Exit Function 
    End If 
    '--������ͼ���̽���-- 
    'ע��ʵ�� 
    Set AspJpeg=Nothing 
    If Err.Number <> 0 Then 
        BuildSmallPic="Error_Other" 
        Err.Clear 
    End If 
    BuildSmallPic=s_BuildBasePath & s_BuildFileName 
End Function %>