<%
'=====================================
'���칩����Ϣ��վ����ϵͳ
'��������Ƽ������ҳ�Ʒ
'�ͷ�����http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>
<%
'�Ƿ��ַ���⺯��
function chkdick(char) 
if instr(char,"'") then 
call errdick(2)
response.end 
end if
if instr(char,"|") then
call errdick(2)
response.end 
end if
if instr(char,"<") then
call errdick(2)
response.end 
end if
if instr(char,">") then
call errdick(2)
response.end 
end if
if instr(char,"&") then
call errdick(2)
response.end 
end if
if instr(char,chr(32)) then
call errdick(2)
response.end 
end if
if instr(char,chr(34)) then
call errdick(2)
response.end 
end if
if instr(char,"%") then
call errdick(2)
response.end 
end if
if instr(char,";") then
call errdick(2)
response.end 
end if
if instr(char," ") then
call errdick(2)
response.end 
end if
if instr(char,"$") then
call errdick(2)
response.end 
end if
end function
'�ʼ���ַ�Ϸ��Լ�⺯��
function IsValidEmail(email)
dim names, name, i, c
IsValidEmail = true
names = Split(email, "@")
if UBound(names) <> 1 then
   IsValidEmail = false
   exit function
end if
for each name in names
   if Len(name) <= 0 then
     IsValidEmail = false
     exit function
   end if
   for i = 1 to Len(name)
     c = Lcase(Mid(name, i, 1))
     if InStr("abcdefghijklmnopqrstuvwxyz_-.", c) <= 0 and not IsNumeric(c) then
       IsValidEmail = false
       exit function
     end if
   next
   if Left(name, 1) = "." or Right(name, 1) = "." then
      IsValidEmail = false
      exit function
   end if
next
if InStr(names(1), ".") <= 0 then
   IsValidEmail = false
   exit function
end if
i = Len(names(1)) - InStrRev(names(1), ".")
if i <> 2 and i <> 3 then
   IsValidEmail = false
   exit function
end if
if InStr(email, "..") > 0 then
   IsValidEmail = false
end if
end function
'�����û�����⺯��
function nothaveChinese(para)
dim str,i,c
nothaveChinese=true
str=cstr(para)
for i = 1 to Len(para)
c=asc(mid(str,i,1))
if c<0 then 
nothaveChinese=false 
exit function
end if
next
end function
%>
