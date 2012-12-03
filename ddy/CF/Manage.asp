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
<%
Response.Expires= -1
Response.Addheader "pragma","no-cache"
Response.AddHeader "cache-control","no-store"
%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<!--#Include File="Conn.asp"-->
<!--#Include File="CF_MyFunCtion.asp"-->
<!--#Include File="CF_Md5.asp"-->
<!--#Include File="CF_Getcookie.asp"-->
<!--#Include File="CF_Manage_Pre.asp"-->
<!--#Include File="CF_Manage_Do.asp"-->


<HTML><HEAD><TITLE><%=RsSet("SysTitle")%></TITLE>
<META http-equiv=Content-Type content="text/html; charset=gb2312">
<!--#Include File="CF_Style.asp"-->
</HEAD>
<BODY>
<!--#include file="Top.asp"-->

<table width="980px" align="center"  style="margin:0px;">
  <tr> 
    <td valign="top" width="202">
<!--#Include File="CF_Manage_Menu.asp"-->
</td>
<td width="778" valign="top">
<!--#Include File="CF_Manage_Select.asp"-->
</td>
</tr>
</table>

<!--#include file="Bottom.asp"-->
</BODY></HTML>
<%Call ConnClose()%>
