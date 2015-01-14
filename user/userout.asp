<!--#include file="../conn.asp"-->
<!--#include file="../SqlIn.Asp"-->
<!--#include file=../CONfig.ASP-->
<!--#include file="../Include/head.asp"-->
<!--#include file="../Include/Function.asp"-->
<!--#include file="../Include/label.asp"-->
<%
Response.cookies("dick")=""
session("openid")=""
Session("Access_Token")=""
session("binding")=""
Response.Redirect "../index.asp?"&C_ID&""%>
