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
'权限检查
if session.Contents("master")=false and whatcan< 6 then Response.Redirect "help.asp?id=004&error=您没有管理自定义统计检索条件库的权限。"

'获取数据
id			=trim(Request("id"))
isok		=trim(Request("isok"))
if isok="" then isok=0
isok=cbool(isok)

'检查数据
if id="" then Response.Redirect "help.asp?id=003&error=请指明您要删除哪条记录。"
'打开数据库
set conn=server.createobject("adodb.connection")
DBPath = Server.MapPath(connpath)
conn.Open "driver={Microsoft Access Driver (*.mdb)};dbq=" & DBPath
Set rs = Server.CreateObject("ADODB.Recordset")

sql="select * from save where id=" & id
rs.Open sql,conn,3,2

if rs.EOF then
	rs.Close
	set rs=nothing
	conn.Close
	set conn=nothing
	Response.Redirect "help.asp?id=003&error=您要删除的记录并不存在。"
end if

%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Copyright" content="Ajiang http://www.jhqing.com">
<title><%=countname%>－删除自定义统计的检索条件</title>
<link rel="stylesheet" type="text/css" href="style.css">
</head>
<body topmargin=5 rightmargin=0 leftmargin=0 vlink=#000000>
<!--#include file="inc_top.asp"-->
<br>

<%

if isok=false then
%>
<table width="500" cellspacing="0" align="center" cellpadding="0" border="0">
  <tr><td colspan="3"><img src="images/photoup.gif"></td></tr>
  <form action="tj_save_del.asp" method=post id=form1 name=form1>
  <tr height="30">
    <td width="1" class="backs"><img src="images/touming.gif" width="1" height="1"></td>
    <td width="498"class="backq">
		&nbsp; <img src="images/tb_save.gif" align=absmiddle><font style="font-size:16px">&nbsp;</font>∷∷∷ 删除检索条件 ∷∷∷<br>

	<p class="p1"><img src='images/tb_error.gif' align=absmiddle><font style='font-size:15px'>&nbsp;</font><font color=red>注意！！</font>
	<p class="p1">您即将删除名为“<%=rs("name")%>”的检索条件记录：
	<p class="p1" style='margin-bottom: 0;'><font class="fonts">检索条件</font>&nbsp;
	<%if rs("wherestr")="" then%>没有检索条件<%else%><%=rs("wherestr")%><%end if%>。
	<p class="p1" style='margin-top: 0;margin-bottom: 0;'><font class="fonts">查询项目</font>&nbsp; <%=rs("outtype")%>。
	<p class="p1" style='margin-top: 0;'><font class="fonts">说　　明</font>&nbsp; <%=rs("content")%>
	<p class="p1">您真的要删除吗？
	<input name="id" size="50" type="hidden" value="<%=id%>">
	<input name="isok" size="50" type="hidden" value="1">
	<p class="p1" align=right><a href="help.asp?id=003">帮助</a>
	<a href='javascript:history.back()'>取消</a> 
	<a href='javascript:document.forms[0].submit();'>删除</a>
	<input type="submit" value=" " name="save" class="backc2"><font style="font-size:16px">&nbsp;</font>&nbsp;
	</td>
    <td width="1" class="backs"><img src="images/touming.gif" width="1" height="1"></td>
  </tr>
  </form>
  <tr><td colspan="4"><img src="images/photodown.gif"></td></tr>
</table>
<br>
<%
rs.Close
else

	rs.delete

	rs.Update
	rs.Close
%>
<table width="500" cellspacing="0" align="center" cellpadding="0" border="0">
  <tr><td colspan="3"><img src="images/photoup.gif"></td></tr>
  <form action="tj_save.asp" method=post id=form1 name=form1>
  <tr height="30">
    <td width="1" class="backs"><img src="images/touming.gif" width="1" height="1"></td>
    <td width="498"class="backq">
		&nbsp; <img src="images/tb_save.gif" align=absmiddle><font style="font-size:16px">&nbsp;</font>∷∷∷ 保存检索条件成功 ∷∷∷<br>

	<p class="p1"><font color=red>恭喜！！</font>
	<p class="p1">已经成功的删除该检索条件。
	<p class="p1" align="right"><a href='tj_search.asp'>继续</a> <a href='tj_search.asp'><img src="images/arbutton.gif" align="absmiddle" border="0"></a> <font style="font-size:16px">&nbsp;</font>&nbsp;
	</td>
    <td width="1" class="backs"><img src="images/touming.gif" width="1" height="1"></td>
  </tr>
  </form>
  <tr><td colspan="4"><img src="images/photodown.gif"></td></tr>
</table>
<br>
<%
end if

set rs=nothing
conn.Close
set conn=nothing
%>
<!--#include file="inc_bottom.asp"-->
</body>
</html>
