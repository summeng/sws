<!--#include file="../conn.asp"-->
<!--#include file="../SqlIn.Asp"-->
<!--#include file=cookies.asp-->
<!--#include file="../config.asp"-->
<%
'=====================================
'���칩����Ϣ��վ����ϵͳ
'��������Ƽ������ҳ�Ʒ
'�ͷ�����http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>
<%
dim objfso,root,str,objFolder,colFiles,objFile,Extensions,filedatem,filedated,filedate,colFolders,objSubFolder,PathExclusion,PathExcluded,fso,objStream
Server.ScriptTimeout=50000
dim seoDir
'#################################�޸�1 �Ѹ���ַ�ĳ��Լ���վ��
session("server")="http://"&web&""     
seoDir="/"

set objfso = CreateObject("Scripting.FileSystemObject")
root = Server.MapPath(seoDir)

str = "<?xml version='1.0' encoding='UTF-8'?>" & vbcrlf
str = str & "<urlset xmlns='http://www.google.com/schemas/sitemap/0.84'>" & vbcrlf

Set objFolder = objFSO.GetFolder(root)
'response.write getfilelink(objFolder.Path,objFolder.dateLastModified)
Set colFiles = objFolder.Files
For Each objFile In colFiles
str=str & getfilelink(objFile.Path,objfile.dateLastModified) & vbcrlf
Next
ShowSubFolders(objFolder)


str = str & "</urlset>" & vbcrlf
set fso = nothing

Set objStream = Server.CreateObject("ADODB.Stream")
With objStream
.Open
.Charset = "utf-8"
.Position = objStream.Size
.WriteText=str
.SaveToFile server.mappath("/"&strInstallDir&"sitemap.xml"),2 '���ɵ�XML�ļ���,�������޸�,������и��õķ���,���Ը�
.Close
End With

Set objStream = Nothing
If Not Err Then
response.Write "<script language=javascript>alert('�ɹ�����վ���ͼ!');location.replace('../sitemap.xml');</script>"
Response.End
End If

Sub ShowSubFolders(objFolder)
Set colFolders = objFolder.SubFolders
For Each objSubFolder In colFolders
if folderpermission(objSubFolder.Path) then
str = str & getfilelink(objSubFolder.Path,objSubFolder.dateLastModified) & vbcrlf
Set colFiles = objSubFolder.Files
For Each objFile In colFiles
str = str & getfilelink(objFile.Path,objFile.dateLastModified) & vbcrlf
Next
ShowSubFolders(objSubFolder)
end if
Next
End Sub


Function getfilelink(file,datafile)
file=replace(file,root,"")
file=replace(file,"\","/")
If FileExtensionIsBad(file) then Exit Function
if month(datafile)<10 then filedatem="0"
if day(datafile)<10 then filedated="0"
filedate=year(datafile)&"-"&filedatem&month(datafile)&"-"&filedated&day(datafile)
getfilelink = "<url><loc>"&server.htmlencode(session("server")&file)&"</loc><lastmod>"&filedate&"</lastmod><changefreq>daily</changefreq><priority>1.0</priority></url>"
Response.Flush
End Function


Function Folderpermission(pathName)
'#################################�޸�2 �㲻ϣ������¼��ҳ��, �м�Ŀ¼��\,��Ҫ��/
PathExclusion=Array("\temp","\hua+admin","inc","\data","\databackup","\admin")

Folderpermission =True
for each PathExcluded in PathExclusion
if instr(ucase(pathName),ucase(PathExcluded))>0 then
Folderpermission = False
exit for
end if
next
End Function


Function FileExtensionIsBad(sFileName)
Dim sFileExtension, bFileExtensionIsValid, sFileExt

'#################################�޸�3 �����б���ļ���,��չ���������еĻ�SiteMap�򲻻���¼����չ�����ļ�
Extensions = Array("html","htm","asp")
if len(trim(sFileName)) = 0 then
FileExtensionIsBad=true
Exit Function
end if

sFileExtension = right(sFileName, len(sFileName) - instrrev(sFileName, "."))
bFileExtensionIsValid=false
for each sFileExt in extensions
if ucase(sFileExt)=ucase(sFileExtension) then
bFileExtensionIsValid=True
exit for
end if
next
FileExtensionIsBad = not bFileExtensionIsValid
End Function
%>
<style type="text/css">
<!--
body {
	background-color: #E3EBF9;
}
-->
</style>