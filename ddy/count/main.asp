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

set conn=server.createobject("adodb.connection")
DBPath = Server.MapPath(connpath)
conn.Open "driver={Microsoft Access Driver (*.mdb)};dbq=" & DBPath

'所有总访问数、开始访问日期（从主数据库读取）
    tmprs=conn.execute("Select count(id) As vtotal, first(vtime) as vfirst from view")
    vtotal=tmprs("vtotal")
    vfirst=tmprs("vfirst")
	if isnull(vtotal) then vtotal=0
	if isnull(vfirst) then vfirst=now()

if vtotal=0 then
	conn.Close
	set conn=nothing
	Response.Redirect "help.asp?error=统计系统还没有启用，尚不能查看统计报告。"
end if

'今日访问量、昨日访问量（从简数据朝读取）
    tmprs=conn.execute("Select today,yesterday from vjian")
    vtoday=tmprs("today")
    vyesterday=tmprs("yesterday")
	if isnull(vtoday) then vtoday=0
	if isnull(vyesterday) then vyesterday=0
'今年访问量
    tmprs=conn.execute("Select count(id) as vthisyear from view where vyear=" & year(now))
    vthisyear=tmprs("vthisyear")
	if isnull(vthisyear) then vthisyear=0
'本月访问量
    tmprs=conn.execute("Select count(id) as vthismonth from view where vmonth=" & month(now) & " and vyear=" & year(now))
    vthismonth=tmprs("vthismonth")
	if isnull(vthismonth) then vthismonth=0
'访问天书、平均每天访问量
	vdays=now()-vfirst
	vdayavg=vtotal/vdays
	vdays=int((vdays*10^mPrecision)+0.5)/(10^mPrecision)
	if vdays<1 then vdays="0" & vdays
	vdayavg=int((vdayavg*10^mPrecision)+0.5)/(10^mPrecision)
'预计今日访问量
	vdaylong=now()-date()
	vguess=int(((vtoday/vdaylong)+vyesterday)/2+0.5)
	if vguess< vtoday then vguess=int((vtoday/vdaylong)+0.5)
'当前用户放量
	vuser=cint(Request.Cookies(mNameEn)("lao"))

conn.Close
set conn=nothing
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Copyright" content="Ajiang http://www.jhqing.com">
<title><%=countname%></title>
<link rel="stylesheet" type="text/css" href="style.css">
</head>
<body topmargin=5 rightmargin=0 leftmargin=0 vlink=#000000>
<!--#include file="inc_top.asp"-->
<br>
<table width="500" cellspacing="0" align="center" cellpadding="0" border="0">
  <tr><td colspan="4"><img src="images/photoup.gif"></td></tr>
  <tr height="30">
    <td width="1" class="backs"></td>
    <td colspan="2" width="489"class="backq">
		&nbsp; <img src="images/tb_title.gif" align=absmiddle> &nbsp;∷∷∷ 总 体 数 据 ∷∷∷ 
	</td>
    <td width="1" class="backs"></td>
  </tr>
  <tr height="50">
    <td width="1" class="backs"></td>
    <td width="220"class="backq" valign=top>
		<p style="line-height: 140%; margin-left: 25; margin-right: 0; margin-top: 10; margin-bottom: 5" align=left>
		总访问量: &nbsp; <%=vtotal%><br>
		您的访问量: <%=vuser%><br>
		开始统计于: <%=vfirst%><br>
		今日访问量: <%=vtoday%><br>
		昨日访问量: <%=vyesterday%><br>
		今年访问量: <%=vthisyear%><br>
		本月访问量: <%=vthismonth%><br>
		统计天数: &nbsp; <%=vdays%><br>
		平均日访量: <%=vdayavg%><br>
		预计今日: &nbsp; <%=vguess%><br>
	</td>
    <td width="278"class="backq" valign=top>
		<p style="line-height: 140%; margin-left: 10; margin-right: 0; margin-top: 10; margin-bottom: 0" align=left>
		网站: <%=mName%><br>
		连接: <a href="<%=mURL%>" class="a1" target="_blank"><%=mURL%></a><br>
		站长: <%=masterName%><br>
		信箱: <a href="mailto:<%=masterEmail%>" class="a1"><%=masterEmail%></a><br>
		<p style="text-indent: -35; line-height: 140%; margin-left: 45; margin-right: 25; margin-top: 0; margin-bottom: 5" align=left>简介: <%=SiteBrief%><br>
	</td>
    <td width="1" class="backs"></td>
  </tr>
  <tr><td colspan="4"><img src="images/photodown.gif"></td></tr>
</table>
<br>
<%if session.Contents("master")=false then%>
<table width="500" cellspacing="0" align="center" cellpadding="0" border="0">
  <tr><td colspan="3"><img src="images/photoup.gif"></td></tr>
<form action="login.asp"  method=post target="_top" id=form1 name=form1>
  <tr height="50">
    <td width="1" class="backs"><img src="images/touming.gif" width="1" height="1"></td>
    <td width=498 class="backq">
&nbsp;&nbsp; <img src="images/tb_login.gif" align=absmiddle> 管理员登录
&nbsp; &nbsp; 用户:<input name="logname" class="input" size=9>
&nbsp;密码:<input type="password" name="logpsw" class="input" size=9>
&nbsp; &nbsp;登录 <input type="submit" name="search" class="backc2" value=" ">
	</td>
    <td width="1" class="backs"><img src="images/touming.gif" width="1" height="1"></td>
  </tr>
</form>
  <tr><td colspan="3"><img src="images/photodown.gif"></td></tr>
</table>
<%else%>
<table width="500" cellspacing="0" align="center" cellpadding="0" border="0">
  <tr><td colspan="3"><img src="images/photoup.gif"></td></tr>
<form action="login.asp"  method=post target="_top" id=form1 name=form1>
  <tr height="50">
    <td width="1" class="backs"><img src="images/touming.gif" width="1" height="1"></td>
    <td width=498 class="backq">
&nbsp;&nbsp; <img src="images/tb_login.gif" align=absmiddle> 管理员登录
&nbsp; &nbsp; 管理员已登录！&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; 
<a href='logout.asp' target="_top">退出 <img src="images/arbutton.gif" align="absmiddle" border="0"></a>	</td>
    <td width="1" class="backs"><img src="images/touming.gif" width="1" height="1"></td>
  </tr>
</form>
  <tr><td colspan="3"><img src="images/photodown.gif"></td></tr>
</table>
<%end if%>
<%if yestop=0 and yesleft=0 then%>
<br>
<!--#include file="inc_index.asp"-->
<%end if%>
<!--#include file="inc_bottom.asp"-->
</body>
</html>
