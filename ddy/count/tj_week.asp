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
if session.Contents("master")=false and whatcan<2 then Response.Redirect "help.asp?id=004&error=您没有查看周访问统计的权限。"
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Copyright" content="Ajiang http://www.jhqing.com">
<title><%=countname%>－周访问统计</title>
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
    <td width="498" class="backq">
    
		&nbsp; <img src="images/tb_title.gif" align=absmiddle> &nbsp;∷∷∷ 周访问量统计 ∷∷∷<br>

<table width="90%" align=center><tr><td>

<table border="0" cellpadding="0" cellspacing="0" width="175" align=center>
<tr height="9"><td></td></tr>
<tr height="101">
<%
set conn=server.createobject("adodb.connection")
DBPath = Server.MapPath(connpath)
conn.Open "driver={Microsoft Access Driver (*.mdb)};dbq=" & DBPath

'找到开始统计天数，如果天数不足7天，则跳过前面的空间
tmprs=conn.execute("Select first(vtime) as vfirst from view")
vfirst=tmprs("vfirst")
if isnull(vfirst) then vfirst=now()
vdays=int(date()-vfirst+1)

'声明二维数组，voutday(*,0)为访问量,voutday(*,1)为日期,voutday(*,2)为星期
dim vday(7,3),voutday(7,3)
maxday=0
sumday=0
for i=0 to 6
	vday(i,0)=vdaycon(date()-6+i)
	if vday(i,0)>maxday then maxday=vday(i,0)
	sumday=sumday+vday(i,0)
	vday(i,1)=day(date()-6+i)
	vday(i,2)=weekday(date()-6+i)
next
'防止除数为0而出错
if maxday=0 then maxday=1
if sumday=0 then sumday=1

'根据已统计天数将数值左移
if vdays>=7 then
	for i=0 to 6
		voutday(i,0)=vday(i,0)
		voutday(i,1)=vday(i,1)
		voutday(i,2)=vday(i,2)
	next
else
	for i=0 to 6
		if i<=vdays then
			voutday(i,0)=vday(i+6-vdays,0)
			voutday(i,1)=vday(i+6-vdays,1)
			voutday(i,2)=vday(i+6-vdays,2)
		else
			voutday(i,0)=0
			voutday(i,1)=""
			voutday(i,2)=0
		end if
	next
end if
%>
<td align=right valign=top>
<p style="line-height: 100%; margin-right: 2; margin-top: 0; margin-bottom: 13">
<font face="Arial"><%=int(maxday*10+0.5)/10%></font>
<p style="line-height: 100%; margin-right: 2; margin-top: 0; margin-bottom: 13">
<font face="Arial"><%=int(3*maxday*10/4+0.5)/10%></font>
<p style="line-height: 100%; margin-right: 2; margin-top: 0; margin-bottom: 13">
<font face="Arial"><%=int(maxday*10/2+0.5)/10%></font>
<p style="line-height: 100%; margin-right: 2; margin-top: 0; margin-bottom: 0">
<font face="Arial"><%=int(maxday*10/4+0.5)/10%><br></font></td>
<td width=10><img src="images/tu_back_left.gif"></td>
<%
for i= 0 to 6
%>
<td width=15 valign=bottom background="images/tu_back.gif" align=center>
<img style="BORDER-BOTTOM: #000000 1px solid" src="images/tu.gif"
	height="<%=(voutday(i,0)/maxday)*100%>" width="9"
	alt="<%=voutday(i,1)%>日，星期<%=findweek(cint(voutday(i,2)))%>，访问<%=voutday(i,0)%>次，<%
	'计算访问量的百分数，精确到小数后1位，小于零的在前面加字母0
	lsbf=int(voutday(i,0)*1000/sumday+0.5)/10
	if lsbf<1 then lsbf="0" & lsbf
	Response.Write lsbf
	%>%"></td>
<%
next
%>
<td width=10><img src="images/tu_back_right.gif"></td>
<td width=10></td>
</tr>

<tr height="18">
<td align=right>
<p style="line-height: 100%; margin-right: 2; margin-top: 0; margin-bottom: 0">
<font face="Arial">0</font></td>
<td width=10></td>
<%
for i= 0 to 6
%>
<td width=15 align=center>
<a title="<%=voutday(i,1)%>日，星期<%=findweek(cint(voutday(i,2)))%>，访问<%=voutday(i,0)%>次，<%
	'计算访问量的百分数，精确到小数后1位，小于零的在前面加字母0
	lsbf=int(voutday(i,0)*1000/sumday+0.5)/10
	if lsbf<1 then lsbf="0" & lsbf
	Response.Write lsbf
	%>%"><%
'根据当天的日期用不同的颜色区分日期，周六用绿色，周日用红色
select case voutday(i,2)
case 1%>
<font face="Arial" style="letter-spacing: -1" color="red">
<%case 7%>
<font face="Arial" style="letter-spacing: -1" class="fonts">
<%case else%>
<font face="Arial" style="letter-spacing: -1">
<%end select%>
<%=findweek(voutday(i,2))%></font></a></td>
<%
next
%>
<td width=10></td>
<td width=10></td>
</tr>
<tr height="5"><td></td></tr>
</table>

</td><td>

<table border="0" cellpadding="0" cellspacing="0" width="175" align=center>
<tr height="9"><td></td></tr>
<tr height="101">
<%
Set rs = Server.CreateObject("ADODB.Recordset")
sql="select vweek,count(id) as allweek from view group by vweek"
rs.Open sql,conn,1,1

dim vallweek(7)
maxallweek=0
sumallweek=0
do while not rs.EOF
	vallweek(cint(rs("vweek"))-1)=cint(rs("allweek"))
	if vallweek(cint(rs("vweek"))-1)>maxallweek then maxallweek=vallweek(cint(rs("vweek"))-1)
	sumallweek=sumallweek+vallweek(cint(rs("vweek"))-1)
	rs.MoveNext
loop
'防止除数为0而出错
if maxallweek=0 then maxallweek=1
if sumallweek=0 then sumallweek=1

%>
<td align=right valign=top>
<p style="line-height: 100%; margin-right: 2; margin-top: 0; margin-bottom: 13">
<font face="Arial"><%=int(maxallweek*10+0.5)/10%></font>
<p style="line-height: 100%; margin-right: 2; margin-top: 0; margin-bottom: 13">
<font face="Arial"><%=int(3*maxallweek*10/4+0.5)/10%></font>
<p style="line-height: 100%; margin-right: 2; margin-top: 0; margin-bottom: 13">
<font face="Arial"><%=int(maxallweek*10/2+0.5)/10%></font>
<p style="line-height: 100%; margin-right: 2; margin-top: 0; margin-bottom: 0">
<font face="Arial"><%=int(maxallweek*10/4+0.5)/10%><br></font></td>
<td width=10><img src="images/tu_back_left.gif"></td>
<%
for i= 0 to 6
%>
<td width=15 valign=bottom background="images/tu_back.gif" align=center>
<img style="BORDER-BOTTOM: #000000 1px solid;" src="images/tu.gif"
	height="<%=(vallweek(i)/maxallweek)*100%>" width="9"
	alt="星期<%=findweek(i+1)%>，访问<%=vallweek(i)%>次，<%
	'计算访问量的百分数，精确到小数后1位，小于零的在前面加字母0
	lsbf=int(vallweek(i)*1000/sumallweek)/10
	if lsbf<1 then lsbf="0" & lsbf
	Response.Write lsbf
	%>%"></td>
<%
next
%>
<td width=10><img src="images/tu_back_right.gif"></td>
<td width=10></td>
</tr>
<tr height="18">
<td align=right>
<p style="line-height: 100%; margin-right: 2; margin-top: 0; margin-bottom: 0">
<font face="Arial">0</font></td>
<td width=10></td>
<%
for i= 0 to 6
%>
<td width=15 align=center><a title="星期<%=findweek(i+1)%>，访问<%=vallweek(i)%>次，<%
	'计算访问量的百分数，精确到小数后1位，小于零的在前面加字母0
	lsbf=int(vallweek(i)*1000/sumallweek)/10
	if lsbf<1 then lsbf="0" & lsbf
	Response.Write lsbf
	%>%">
	<font face="Arial" style="letter-spacing: -1"><%=findweek(i+1)%></font></a></td>
<%
next
%>
<td width=10></td>
<td width=10></td>
</tr>
<tr height="5"><td colspan=29></td></tr>
</table>

</td></tr>
<tr height="20" align=center><td>↑ 最近的一周</td><td>↑ 全部时段</td></tr>
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
<%
'计算指定日期的访问量
function vdaycon(theday)
    theday=cdate(theday)
    thetday=cdate(theday+1)
    tmprs=conn.execute("Select count(id) as vdaycon from view where" & _
		" vtime>=datevalue('" & theday & "') and vtime<=datevalue('" & thetday & "')")
    vdaycon=tmprs("vdaycon")
	if isnull(vdaycon) then vdaycon=0
end function

'将星期序号翻译为汉字
function findweek(theweek)
	select case theweek
	case 1
		findweek="日"
	case 2
		findweek="一"
	case 3
		findweek="二"
	case 4
		findweek="三"
	case 5
		findweek="四"
	case 6
		findweek="五"
	case 7
		findweek="六"
	end select
end function
%>