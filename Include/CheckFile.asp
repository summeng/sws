<%
'===============ASP文件操作函数集========================= 
'     作者：天天科技
'     网址：http://www.ip126.com 
'     用途：检查物理文件 
'     注意：所有函数使用的文件地址全部使用绝对地址 
'=========================================================

'******************************************************* 
'作 用: 上传文件扩展名检测 
'函数名: CheckFileExt 
'参 数: fileEXT 上传文件的后缀 
'       UpFileType   允许或禁止上传文件的后缀,多个以"|"分隔 
'       blnAllow 是允许还是禁止上传 UpFileType 中指定的后缀 
'返回值: 合法文件返回 True ,否则返回False 
'******************************************************* 
Function CheckFileExt(fileEXT,UpFileType,blnAllow) 
dim arrExt,return,i 
fileEXT = UCase(fileEXT) 
UpFileType   = UCase(UpFileType)    
arrExt = split(UpFileType,";") 
If blnAllow=true then         '只允许上传指定的文件 
return = false 
for i=0 to UBound(arrExt) 
If fileEXT=arrExt(i) then return=true 
next 
else                          '禁止上传指定的文件 
return = true 
for i=0 to UBound(arrExt) 
If fileEXT=arrExt(i) then return=false 
next 
End If 
CheckFileExt = return 
End Function 


'******************************************************* 
'作 用: 格式化显示文件大小 
'FileSize: 文件大小 
'******************************************************* 
Function FormatSize(FileSize) 
If FileSize<1024 then FormatSize = FileSize & " Byte" 
If FileSize/1024 <1024 And FileSize/1024 > 1 then 
FileSize = FileSize/1024 
FormatSize=round(FileSize*100)/100 & " KB" 
Elseif FileSize/(1024*1024) > 1 Then 
FileSize = FileSize/(1024*1024) 
FormatSize = round(FileSize*100)/100 & " MB" 
End If 
End function 


'-------------------------------------------
'函数名：ChkImg(img)
'作　用：检查图片文件是否合法
'参　数：img，图片路径，相对于网站根目录的绝对路径
'返回值：布尔类型，如果图片合法返回True,否则返回False
'条　件：服务器必须支持AspJpeg组件
'如不支持，为了避免所有图片都不能上传，本函数将直接返回True
'-------------------------------------------
Function ChkImg(img)
    On Error Resume Next '为了捕获错误信息，需要让代码在出错时能继续执行
    Dim RetunValue, ChkJpeg
    RetunValue = True
    '如果路径为空，则认为图片不合法
    If isnull(img) Then ChkImg = False:Exit Function
    Set ChkJpeg = Server.CreateObject("Persits.Jpeg")
    If -2147221005 <> Err Then    '如果组件被支持，则用组件检查图片的合法性
        ChkJpeg.Open Server.mappath(img)
        If Err Then
            RetunValue = False
        End If
    Else    '如果组件不被支持，则跳过直接返回True
        RetunValue = True
    End If
    '必要的善后工作
    If Err.number <> 0 Then Err.clear
    Set ChkJpeg = Nothing
    ChkImg = RetunValue
End Function


'-------------------------------------------
'函数名：Chkfile(my_string)
'作　用：检查图片文件是否经过更改后缀名的脚本文件
'参　数：(my_string，文件路径，文件夹的物理全路径
'返回值：布尔类型，如果不是脚本文件返回True,否则返回False
'-------------------------------------------
Function Chkfile(my_string)
if instr(LCase(my_string),"&lt;%")<>0 or instr(LCase(my_string),"<%")<>0  or instr(LCase(my_string),"Request")<>0  or instr(LCase(my_string),"Session")<>0 or instr(LCase(my_string),"script")<>0 Then
Chkfile= False
Else
Chkfile=True
End if
End Function


'-------------------------------------------
'函数名：CheckFileType(filename)
'作　用：检查文件是否经过更改后缀名的脚本文件
'参　数：filename，文件路径，文件夹的物理全路径
'返回值：布尔类型，如果是脚本文件返回False
'-------------------------------------------
function CheckFileType(filename)
Dim MyFile,MyText,sTextAll,sStr,sNoString,i,filedel
set MyFile = server.CreateObject("Scripting.FileSystemObject")
set MyText = MyFile.OpenTextFile(filename, 1) '读取文本文件
sTextAll = lcase(MyText.ReadAll)
MyText.close
set MyFile = nothing
sStr="<%|.getfolder|.createfolder|.deletefolder|.createdirectory|.deletedirectory|.saveas|wscript.shell|script.encode|server.|.createobject|execute|activexobject|language="
sNoString = split(sStr,"|")
for i=0 to ubound(sNoString)
if instr(sTextAll,sNoString(i)) Then
CheckFileType=False
Response.End()
end if
next
end Function

'******************************************************* 
'作 用: 检测上传的图片文件(jpeg,gif,bmp,png)是否真的为图片 
'函数名: IsImgFile(sFileName) 
'参 数: sFileName 文件名(此处文件名是文件夹的物理全路径) 
'返回值: 确实为图片文件则返回 True ,否则返回False 
'******************************************************* 
Function IsImgFile(sFileName) 
const adTypeBinary=1 
dim return 
dim jpg(1):jpg(0)=CByte(&HFF):jpg(1)=CByte(&HD8) 
dim bmp(1):bmp(0)=CByte(&H42):bmp(1)=CByte(&H4D) 
dim png(3):png(0)=CByte(&H89):png(1)=CByte(&H50):png(2)=CByte(&H4E):png(3)=CByte(&H47) 
dim gif(5):gif(0)=CByte(&H47):gif(1)=CByte(&H49):gif(2)=CByte(&H46):gif(3)=CByte(&H39):gif(4)=CByte(&H38):gif(5)=CByte(&H61) 

on error resume next
return=false 
dim fstream,fileExt,stamp,i 
'得到文件后缀并转化为小写 
FileExt = LCase(GetFileExt(sFileName)) 
'如果文件后缀为 jpg,jpeg,bmp,gif,png 中的任一种 
'则执行真实图片判断 
If strInString(FileExt,"jpg|jpeg|bmp|gif|png")=true then 
Set fstream=Server.createobject("ADODB.Stream") 
fstream.Open 
fstream.Type=adTypeBinary 
fstream.LoadFromFile sFileName 
fstream.position=0 
select case LCase(FileExt) 
case "jpg","jpeg" 
stamp=fstream.read(2) 
for i=0 to 1 
If ascB(MidB(stamp,i+1,1))=jpg(i) then return=true else return=false 
next 
'http://www.knowsky.com
case "gif" 
stamp=fstream.read(6) 
for i=0 to 5 
If ascB(MidB(stamp,i+1,1))=gif(i) then return=true else return=false 
next 
case "png" 
stamp=fstream.read(4) 
for i=0 to 3 
If ascB(MidB(stamp,i+1,1))=png(i) then return=true else return=false 
next 
case "bmp" 
stamp=fstream.read(2) 
for i=0 to 1 
If ascB(MidB(stamp,i+1,1))=bmp(i) then return=true else return=false 
next 
End select 

fstream.Close 
Set fseteam=nothing 
If err.number<>0 then return = false 
else 
return = true 
End If 
IsImgFile = return 
End function %>