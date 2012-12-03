<%@ CODEPAGE = "936" %>
<%
'=================================
'
'      亦凡酷站访问统计系统
'    Ajiang   zhqi123@163.com 
'        www.jhqing.com
'
'     版权所有·抄袭挪用必究
'
'=================================
%>
<!--#include file="inc_config.asp"-->
<%
if Request("logpsw")=masterpsw then
	session.Contents("master")=true
	Response.Redirect "index.asp"
else
	Response.Redirect "help.asp?error=登录密码输入错误。"
end if
%>