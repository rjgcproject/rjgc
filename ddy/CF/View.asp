<%
'乘风多用户计数器
'作者QQ：178575
'作者EMail：yliangcf@163.com
'作者网站：http://www.qqcf.com
'详细简介：http://www.qqcf.com/?action=list&list=cfcount
'上面有程序在线演示，安装演示，使用疑难解答，最新版本下载等内容
'因为这些内容可能时常更新，就没有放在程序里，请自己上网站上查看
'有完整版本的演示
%>
<!--#include file="Conn.asp"-->
<!--#include file="CF_MyFunction.asp"-->
<!--#include file="CF_Md5.asp"-->
<!--#Include File="CF_Getcookie.asp"-->
<%
Action=Request("Action")
%>
<!--#Include File="CF_Do.asp"-->

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<HTML><HEAD><TITLE><%=RsSet("SysTitle")%></TITLE>
<META http-equiv=Content-Type content="text/html; charset=gb2312">
<!--#include file="CF_Style.asp"-->
</HEAD>
<BODY>
<!--#include file="Top.asp"-->

<!--#Include File="CF_View_Pre.asp"-->

<!--#include file="Bottom.asp"-->
</BODY></HTML>

<%Call ConnClose()%>