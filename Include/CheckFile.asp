<%
'===============ASP�ļ�����������========================= 
'     ���ߣ�����Ƽ�
'     ��ַ��http://www.ip126.com 
'     ��;����������ļ� 
'     ע�⣺���к���ʹ�õ��ļ���ַȫ��ʹ�þ��Ե�ַ 
'=========================================================

'******************************************************* 
'�� ��: �ϴ��ļ���չ����� 
'������: CheckFileExt 
'�� ��: fileEXT �ϴ��ļ��ĺ�׺ 
'       UpFileType   ������ֹ�ϴ��ļ��ĺ�׺,�����"|"�ָ� 
'       blnAllow �������ǽ�ֹ�ϴ� UpFileType ��ָ���ĺ�׺ 
'����ֵ: �Ϸ��ļ����� True ,���򷵻�False 
'******************************************************* 
Function CheckFileExt(fileEXT,UpFileType,blnAllow) 
dim arrExt,return,i 
fileEXT = UCase(fileEXT) 
UpFileType   = UCase(UpFileType)    
arrExt = split(UpFileType,";") 
If blnAllow=true then         'ֻ�����ϴ�ָ�����ļ� 
return = false 
for i=0 to UBound(arrExt) 
If fileEXT=arrExt(i) then return=true 
next 
else                          '��ֹ�ϴ�ָ�����ļ� 
return = true 
for i=0 to UBound(arrExt) 
If fileEXT=arrExt(i) then return=false 
next 
End If 
CheckFileExt = return 
End Function 


'******************************************************* 
'�� ��: ��ʽ����ʾ�ļ���С 
'FileSize: �ļ���С 
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
'��������ChkImg(img)
'�����ã����ͼƬ�ļ��Ƿ�Ϸ�
'�Ρ�����img��ͼƬ·�����������վ��Ŀ¼�ľ���·��
'����ֵ���������ͣ����ͼƬ�Ϸ�����True,���򷵻�False
'������������������֧��AspJpeg���
'�粻֧�֣�Ϊ�˱�������ͼƬ�������ϴ�����������ֱ�ӷ���True
'-------------------------------------------
Function ChkImg(img)
    On Error Resume Next 'Ϊ�˲��������Ϣ����Ҫ�ô����ڳ���ʱ�ܼ���ִ��
    Dim RetunValue, ChkJpeg
    RetunValue = True
    '���·��Ϊ�գ�����ΪͼƬ���Ϸ�
    If isnull(img) Then ChkImg = False:Exit Function
    Set ChkJpeg = Server.CreateObject("Persits.Jpeg")
    If -2147221005 <> Err Then    '��������֧�֣�����������ͼƬ�ĺϷ���
        ChkJpeg.Open Server.mappath(img)
        If Err Then
            RetunValue = False
        End If
    Else    '����������֧�֣�������ֱ�ӷ���True
        RetunValue = True
    End If
    '��Ҫ���ƺ���
    If Err.number <> 0 Then Err.clear
    Set ChkJpeg = Nothing
    ChkImg = RetunValue
End Function


'-------------------------------------------
'��������Chkfile(my_string)
'�����ã����ͼƬ�ļ��Ƿ񾭹����ĺ�׺���Ľű��ļ�
'�Ρ�����(my_string���ļ�·�����ļ��е�����ȫ·��
'����ֵ���������ͣ�������ǽű��ļ�����True,���򷵻�False
'-------------------------------------------
Function Chkfile(my_string)
if instr(LCase(my_string),"&lt;%")<>0 or instr(LCase(my_string),"<%")<>0  or instr(LCase(my_string),"Request")<>0  or instr(LCase(my_string),"Session")<>0 or instr(LCase(my_string),"script")<>0 Then
Chkfile= False
Else
Chkfile=True
End if
End Function


'-------------------------------------------
'��������CheckFileType(filename)
'�����ã�����ļ��Ƿ񾭹����ĺ�׺���Ľű��ļ�
'�Ρ�����filename���ļ�·�����ļ��е�����ȫ·��
'����ֵ���������ͣ�����ǽű��ļ�����False
'-------------------------------------------
function CheckFileType(filename)
Dim MyFile,MyText,sTextAll,sStr,sNoString,i,filedel
set MyFile = server.CreateObject("Scripting.FileSystemObject")
set MyText = MyFile.OpenTextFile(filename, 1) '��ȡ�ı��ļ�
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
'�� ��: ����ϴ���ͼƬ�ļ�(jpeg,gif,bmp,png)�Ƿ����ΪͼƬ 
'������: IsImgFile(sFileName) 
'�� ��: sFileName �ļ���(�˴��ļ������ļ��е�����ȫ·��) 
'����ֵ: ȷʵΪͼƬ�ļ��򷵻� True ,���򷵻�False 
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
'�õ��ļ���׺��ת��ΪСд 
FileExt = LCase(GetFileExt(sFileName)) 
'����ļ���׺Ϊ jpg,jpeg,bmp,gif,png �е���һ�� 
'��ִ����ʵͼƬ�ж� 
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