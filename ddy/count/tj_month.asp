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
if session.Contents("master")=false and whatcan<2 then Response.Redirect "help.asp?id=004&error=您没有查看月统计的权限。"
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Copyright" content="Ajiang http://www.jhqing.com">
<title><%=countname%>－月访问统计</title>
<link rel="stylesheet" type="text/css" href="style.css">
</head>
<body topmargin=5 rightmargin=0 leftmargin=0 vlink=#000000>
<!--#include file="inc_top.asp"-->
<br>
<table width=500 cellspacing=0 align=center>
  <tr><td>
    <p style="line-height: 100%; margin-left: 15; margin-top: 5; margin-bottom: 0">
    Tips: 用鼠标点指图形柱或者图形柱下的数字可以看到对应的访问量。</p>
  </td></tr>
</table>
<br>
<table width="500" cellspacing="0" align="center" cellpadding="0" border="0">
  <tr><td colspan="3"><img src="images/photoup.gif"></td></tr>
  <tr height="30">
    <td width="1" class="backs"></td>
    <td width="498"class="backq">
    
		&nbsp; <img src="images/tb_title.gif" align=absmiddle> &nbsp;∷∷∷ 最近12个月访问量 ∷∷∷<br>
<%
set conn=server.createobject("adodb.connection")
DBPath = Server.MapPath(connpath)
conn.Open "driver={Microsoft Access Driver (*.mdb)};dbq=" & DBPath
%>

<table border="0" cellpadding="0" cellspacing="0" width="310" align=center>
<tr height="9"><td></td></tr>
<tr height="101">
<%
Set rs = Server.CreateObject("ADODB.Recordset")
'计算12个月（零头也算一个月）前的时间
datetwelve=dateadd("m",-11,date())
datetwelve=cdate(year(datetwelve) & "-" & month(datetwelve) & "-1")

sql="select vmonth,count(id) as allmonth from view where vtime>=datevalue('" & _
	datetwelve & "') group by vmonth"
rs.Open sql,conn,1,1

dim vallmonth(12)
maxallmonth=0
sumallmonth=0
do while not rs.EOF
	vallmonth(cint(rs("vmonth"))-1)=cint(rs("allmonth"))
	if vallmonth(cint(rs("vmonth"))-1)>maxallmonth then maxallmonth=vallmonth(cint(rs("vmonth"))-1)
	sumallmonth=sumallmonth+vallmonth(cint(rs("vmonth"))-1)
	rs.MoveNext
loop
	'防止除数为零而出错
	if maxallmonth=0 then maxallmonth=1
	if sumallmonth=0 then sumallmonth=1

%>
<td align=right valign=top>
<p style="line-height: 100%; margin-right: 2; margin-top: 0; margin-bottom: 13">
<font face="Arial"><%=int(maxallmonth*10+0.5)/10%></font>
<p style="line-height: 100%; margin-right: 2; margin-top: 0; margin-bottom: 13">
<font face="Arial"><%=int(3*maxallmonth*10/4+0.5)/10%></font>
<p style="line-height: 100%; margin-right: 2; margin-top: 0; margin-bottom: 13">
<font face="Arial"><%=int(maxallmonth*10/2+0.5)/10%></font>
<p style="line-height: 100%; margin-right: 2; margin-top: 0; margin-bottom: 0">
<font face="Arial"><%=int(maxallmonth*10/4+0.5)/10%><br></font></td>
<td width=10><img src="images/tu_back_left.gif"></td>
<%
for i= 0 to 11
themonth=month(now())+i
if themonth>11 then themonth=themonth-12

%>
<td width=20 valign=bottom background="images/tu_back.gif" align=center>
<img style="BORDER-BOTTOM: #000000 1px solid;" src="images/tu.gif"
	height="<%=(vallmonth(themonth)/maxallmonth)*100%>" width="9"
	alt="<%=themonth+1%>月，访问<%=vallmonth(themonth)%>次，<%
	'计算访问量的百分数，精确到小数后1位，小于零的在前面加字母0
	lsbf=int(vallmonth(themonth)*1000/sumallmonth)/10
	if lsbf<1 then lsbf="0" & lsbf
	Response.Write lsbf
	%>%"></td>
<%
next
%>
<td width=10><img src="images/tu_back_right.gif"></td>
<td width=10></td>
</tr>
<tr>
<td align=right>
<p style="line-height: 100%; margin-right: 2; margin-top: 0; margin-bottom: 0">
<font face="Arial">0</font></td>
<td width=10></td>
<%
for i= 0 to 11
themonth=month(now())+i
if themonth>11 then themonth=themonth-12
%>
<td width=20 align=center><a title="<%=themonth+1%>月，访问<%=vallmonth(themonth)%>次，<%
	'计算访问量的百分数，精确到小数后1位，小于零的在前面加字母0
	lsbf=int(vallmonth(themonth)*1000/sumallmonth)/10
	if lsbf<1 then lsbf="0" & lsbf
	Response.Write lsbf
	%>%">
	<font face="Arial" style="letter-spacing: -1"><%=themonth+1%></font></a></td>
<%
next
%>
<td width=10></td>
<td width=10></td>
</tr>
<tr height="5"><td colspan=29></td></tr>
</table>

	</td>
    <td width="1" class="backs"></td>
  </tr>
  <tr><td colspan="4"><img src="images/photodown.gif"></td></tr>
</table>
<%
rs.Close
%>


<br>



<table width="500" cellspacing="0" align="center" cellpadding="0" border="0">
  <tr><td colspan="3"><img src="images/photoup.gif"></td></tr>
  <tr height="30">
    <td width="1" class="backs"></td>
    <td width="498"class="backq">
    
		&nbsp; <img src="images/tb_title.gif" align=absmiddle> &nbsp;∷∷∷ 所有12个月访问量 ∷∷∷<br>

<table border="0" cellpadding="0" cellspacing="0" width="310" align=center>
<tr height="9"><td></td></tr>
<tr height="101">
<%

sql="select vmonth,count(id) as allmonth from view group by vmonth"
rs.Open sql,conn,1,1

for i=0 to 11
	vallmonth(i)=0
next
maxallmonth=0
sumallmonth=0
do while not rs.EOF
	vallmonth(cint(rs("vmonth"))-1)=cint(rs("allmonth"))
	if vallmonth(cint(rs("vmonth"))-1)>maxallmonth then maxallmonth=vallmonth(cint(rs("vmonth"))-1)
	sumallmonth=sumallmonth+vallmonth(cint(rs("vmonth"))-1)
	rs.MoveNext
loop
	'防止除数为零而出错
	if maxallmonth=0 then maxallmonth=1
	if sumallmonth=0 then sumallmonth=1

%>
<td align=right valign=top>
<p style="line-height: 100%; margin-right: 2; margin-top: 0; margin-bottom: 13">
<font face="Arial"><%=int(maxallmonth*10+0.5)/10%></font>
<p style="line-height: 100%; margin-right: 2; margin-top: 0; margin-bottom: 13">
<font face="Arial"><%=int(3*maxallmonth*10/4+0.5)/10%></font>
<p style="line-height: 100%; margin-right: 2; margin-top: 0; margin-bottom: 13">
<font face="Arial"><%=int(maxallmonth*10/2+0.5)/10%></font>
<p style="line-height: 100%; margin-right: 2; margin-top: 0; margin-bottom: 0">
<font face="Arial"><%=int(maxallmonth*10/4+0.5)/10%><br></font></td>
<td width=10><img src="images/tu_back_left.gif"></td>
<%
for i= 0 to 11
themonth=i

%>
<td width=20 valign=bottom background="images/tu_back.gif" align=center>
<img style="BORDER-BOTTOM: #000000 1px solid;" src="images/tu.gif"
	height="<%=(vallmonth(themonth)/maxallmonth)*100%>" width="9"
	alt="<%=themonth+1%>月，访问<%=vallmonth(themonth)%>次，<%
	'计算访问量的百分数，精确到小数后1位，小于零的在前面加字母0
	lsbf=int(vallmonth(themonth)*1000/sumallmonth)/10
	if lsbf<1 then lsbf="0" & lsbf
	Response.Write lsbf
	%>%"></td>
<%
next
%>
<td width=10><img src="images/tu_back_right.gif"></td>
<td width=10></td>
</tr>
<tr>
<td align=right>
<p style="line-height: 100%; margin-right: 2; margin-top: 0; margin-bottom: 0">
<font face="Arial">0</font></td>
<td width=10></td>
<%
for i= 0 to 11
themonth=i
%>
<td width=20 align=center><a title="<%=themonth+1%>月，访问<%=vallmonth(themonth)%>次，<%
	'计算访问量的百分数，精确到小数后1位，小于零的在前面加字母0
	lsbf=int(vallmonth(themonth)*1000/sumallmonth)/10
	if lsbf<1 then lsbf="0" & lsbf
	Response.Write lsbf
	%>%">
	<font face="Arial" style="letter-spacing: -1"><%=themonth+1%></font></a></td>
<%
next
%>
<td width=10></td>
<td width=10></td>
</tr>
<tr height="5"><td colspan=29></td></tr>
</table>

	</td>
    <td width="1" class="backs"></td>
  </tr>
  <tr><td colspan="4"><img src="images/photodown.gif"></td></tr>
</table>
<%
rs.Close
%>

<br>

<table width="500" cellspacing="0" align="center" cellpadding="0" border="0">
  <tr><td colspan="3"><img src="images/photoup.gif"></td></tr>
  <tr height="30">
    <td width="1" class="backs"></td>
    <td width="498"class="backq">
    
		&nbsp; <img src="images/tb_title.gif" align=absmiddle> &nbsp;∷∷∷ 年访问量统计 ∷∷∷<br>

<table border="0" cellpadding="0" cellspacing="0" width="270" align=center>
<tr height="9"><td></td></tr>
<tr height="10">
	<td width="40"></td><td width="230"><img src="images/tu_back_up.gif"></td>
</tr>
<%

sql="select vyear,count(id) as allyear from view group by vyear order by vyear DESC"
rs.Open sql,conn,1,1

maxallyear=0
sumallyear=0
do while not rs.EOF
	if cint(rs("allyear"))>maxallyear then maxallyear=cint(rs("allyear"))
	sumallyear=sumallyear+cint(rs("allyear"))
	rs.MoveNext
loop
	'防止除数为零而出错
	if maxallyear=0 then maxallyear=1
	if sumallyear=0 then sumallyear=1

rs.MoveFirst 
do while not rs.EOF
theyear=rs("vyear")
vallyear=rs("allyear")
%>
<tr>
<td width="40" align=right><a title="<%=theyear%>年，访问<%=vallyear%>次，<%
	'计算访问量的百分数，精确到小数后1位，小于零的在前面加字母0
	lsbf=int(vallyear*1000/sumallyear)/10
	if lsbf<1 then lsbf="0" & lsbf
	Response.Write lsbf
	%>%"><%=theyear%></a>&nbsp;</td>
<td width="230" background="images/tu_back_2.gif" align=left>
<img style="BORDER-left: #000000 1px solid;" src="images/tu.gif"
	width="<%=(vallyear/maxallyear)*183%>" height="9"
	alt="<%=theyear%>年，访问<%=vallyear%>次，<%
	'计算访问量的百分数，精确到小数后1位，小于零的在前面加字母0
	lsbf=int(vallyear*1000/sumallyear)/10
	if lsbf<1 then lsbf="0" & lsbf
	Response.Write lsbf
	%>%"> <%=vallyear%></td>
</tr>
<%
rs.MoveNext
loop
%>
<tr height="10">
	<td width="40"></td><td width="230"><img src="images/tu_back_down.gif"></td>
</tr>

<tr height="5"><td colspan=29></td></tr>
</table>

	</td>
    <td width="1" class="backs"></td>
  </tr>
  <tr><td colspan="4"><img src="images/photodown.gif"></td></tr>
</table>
<%
rs.Close

set rs=nothing
conn.Close 
set conn=nothing
%>
<br>
<!--#include file="inc_bottom.asp"-->
</body>
</html>
