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
'权限分配
if session.Contents("master")=false and whatcan<1 then Response.Redirect "help.asp?id=004&error=本统计系统管理员不允许访客查看任何信息。"
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Copyright" content="Ajiang http://www.jhqing.com">
<meta http-equiv="refresh" content="300; url='list.asp'">
<title><%=countname%></title>
<link rel="stylesheet" type="text/css" href="style.css">
</head>
<body style="BACKGROUND-IMAGE:url('images/leftback.gif')" class=backq>
<table>
<tr height="240"><td valign=top>
<p style="margin-left: 17;margin-right: 0;	margin-top: 62;	margin-bottom: 5">
<%if session.Contents("master")=true or whatcan>0 then%><a href="main.asp" target="countmain">※ 总体数据</a><%end if%>
<%if session.Contents("master")=true or whatcan>2 then%><br><a href="tj_all.asp" target="countmain">※ 详细记录</a><%end if%>
<%if session.Contents("master")=true or whatcan>1 then%><br><a href="tj_hour.asp" target="countmain">※ 24小时统计</a>
<br><a href="tj_day.asp" target="countmain">※ 日统计</a>
<br><a href="tj_week.asp" target="countmain">※ 周统计</a>
<br><a href="tj_month.asp" target="countmain">※ 月统计</a>
<br><a href="tj_come.asp" target="countmain">※ 来路统计</a>
<br><a href="tj_page.asp" target="countmain">※ 被访页面</a>
<br><a href="tj_ip.asp" target="countmain">※ IP统计</a>
<br><a href="tj_soft.asp" target="countmain">※ 客户端软件</a>
<br><a href="tj_where.asp" target="countmain">※ 访问者地区</a><%end if%>
</td></tr>
<tr height=60><td valign=top>
<%if session.Contents("master")=true or whatcan>3 then%>
<p style="margin-left: 17;margin-right: 0;	margin-top: 5;	margin-bottom: 5"><a href="tj_search.asp" target="countmain">自定义统计</a>
<%set conn=server.createobject("adodb.connection")
DBPath = Server.MapPath(connpath)
conn.Open "driver={Microsoft Access Driver (*.mdb)};dbq=" & DBPath
Set rs = Server.CreateObject("ADODB.Recordset")

sql="select top 2 * from save"
rs.Open sql,conn,1,1
do while not rs.EOF
%>
<br><a href="tj_searchgo.asp?wherestr=<%=server.URLEncode(rs("wherestr"))%>&outtype=<%=server.URLEncode(rs("outtype"))%>&name=<%=server.URLEncode(rs("name"))%>&content=<%=server.URLEncode(rs("content"))%>" target="countmain">※ <%=left(rs("name"),6)%></a>
<%
rs.MoveNext
loop
rs.Close
set rs=nothing
conn.Close
set conn=nothing

end if%>
<br>&nbsp; &nbsp; &nbsp; &nbsp; <a href="tj_search.asp" target="countmain">更多...</a>
</td></tr>
<%if session.Contents("master")=false then%>
<form action="login.asp"  method=post target="_top">
<tr><td>
<br>&nbsp;<img src="images/tb_login.gif" align=absmiddle> <font color=white>管理员登录
<br>&nbsp;用户:<input name="logname" class="input" size=9>
<br>&nbsp;密码:<input type="password" name="logpsw" class="input" size=9></font>
<input type="submit" name="search" class="backc2" value=" ">
</td></tr>
</form>
<%else%>
<tr><td>
<br>&nbsp;<img src="images/tb_login.gif" align=absmiddle> <font color=white>管理员登出
	<br>&nbsp; 管理员已登录！
<br>&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; <a  target="_top" href='logout.asp' class=a2>退出 <img src="images/arbutton.gif" align="absmiddle" border="0"></a>
</td></tr>
<%end if%>
</table>
</body>
</noframes> 
</html>
