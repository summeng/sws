<!--#include file="../conn.asp"-->
<!--#include file="../config.asp"-->
<!--#include file="../SqlIn.Asp"-->
<%
'=====================================
'天天供求信息网站管理系统
'天天网络科技工作室出品
'客服中心http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>
<%Dim action
action=Request.QueryString("action")
If action="rmb_mc" Then response.write "document.write('<font color=""#ff0000"">"&rmb_mc&"</font>');" & vbCrLf
If action="rmb_hb" Then response.write "document.write('<font color=""#ff0000"">"&rmb_hb&"</font>');" & vbCrLf
If action="huiyuanxx" Then response.write "document.write('<font color=""#ff0000"">"&huiyuanxx&"</font>');" & vbCrLf
If action="huiyuansp" Then response.write "document.write('<font color=""#ff0000"">"&huiyuansp&"</font>');" & vbCrLf
If action="z_a" Then response.write "document.write('<font color=""#ff0000"">"&z_a&"</font>');" & vbCrLf
If action="z_b" Then response.write "document.write('<font color=""#ff0000"">"&z_b&"</font>');" & vbCrLf
If action="z_c" Then response.write "document.write('<font color=""#ff0000"">"&z_c&"</font>');" & vbCrLf
If action="z_d" Then response.write "document.write('<font color=""#ff0000"">"&z_d&"</font>');" & vbCrLf
If action="g_a" Then response.write "document.write('<font color=""#ff0000"">"&g_a&"</font>');" & vbCrLf
If action="g_b" Then response.write "document.write('<font color=""#ff0000"">"&g_b&"</font>');" & vbCrLf
If action="g_c" Then response.write "document.write('<font color=""#ff0000"">"&g_c&"</font>');" & vbCrLf
If action="g_d" Then response.write "document.write('<font color=""#ff0000"">"&g_d&"</font>');" & vbCrLf
If action="jf_1" Then response.write "document.write('<font color=""#ff0000"">"&jf_1&"</font>');" & vbCrLf
If action="jf_2" Then response.write "document.write('<font color=""#ff0000"">"&jf_2&"</font>');" & vbCrLf
If action="tgjf" Then response.write "document.write('<font color=""#ff0000"">"&tgjf&"</font>');" & vbCrLf
If action="jf_3" Then response.write "document.write('<font color=""#ff0000"">"&jf_3&"</font>');" & vbCrLf
If action="jf_4" Then response.write "document.write('<font color=""#ff0000"">"&jf_4&"</font>');" & vbCrLf
If action="jf_5" Then response.write "document.write('<font color=""#ff0000"">"&jf_5&"</font>');" & vbCrLf
If action="z_hb" Then response.write "document.write('<font color=""#ff0000"">"&z_hb&"</font>');" & vbCrLf
If action="jf_hb" Then response.write "document.write('<font color=""#ff0000"">"&jf_hb&"</font>');" & vbCrLf
If action="vipje" Then response.write "document.write('<font color=""#ff0000"">"&vipje&"</font>');" & vbCrLf
If action="title" Then response.write "document.write('<font color=""#0000ff"">"&webname&"</font>');" & vbCrLf
If action="web" Then response.write "document.write('<font color=""#0000ff"">"&web&"</font>');" & vbCrLf
%>

