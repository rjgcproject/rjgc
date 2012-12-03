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
if session.Contents("master")=false and whatcan<2 then Response.Redirect "help.asp?id=004&error=您没有查看来路统计的权限。"
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Copyright" content="Ajiang http://www.jhqing.com">
<title><%=countname%>－来路及访问统计</title>
<link rel="stylesheet" type="text/css" href="style.css">
</head>
<body topmargin=5 rightmargin=0 leftmargin=0 vlink=#000000>
<!--#include file="inc_top.asp"-->
<br>
<table width=500 cellspacing=0 align=center>
  <tr><td>
    <p style="line-height: 100%;text-indent: -35; margin-left: 50; margin-top: 5; margin-bottom: 10">
    Tips: 用鼠标点指图形柱或者网址可以看到对应的访问量。要得到某一时段的统计信息，请使用自定义统计。</p>
  </td></tr>
</table>
<%
set conn=server.createobject("adodb.connection")
DBPath = Server.MapPath(connpath)
conn.Open "driver={Microsoft Access Driver (*.mdb)};dbq=" & DBPath
Set rs = Server.CreateObject("ADODB.Recordset")
%>

<table width="500" cellspacing="0" align="center" cellpadding="0" border="0">
  <tr><td colspan="3"><img src="images/photoup.gif"></td></tr>
  <tr height="30">
    <td width="1" class="backs"></td>
    <td width="498"class="backq">
    
		&nbsp; <img src="images/tb_title.gif" align=absmiddle> &nbsp;∷∷∷ 来路及访问量 ∷∷∷<br>

<table border="0" cellpadding="0" cellspacing="0" width="450" align=center>
<tr height="9"><td></td></tr>
<tr height="10">
	<td width="220"></td><td width="230"><img src="images/tu_back_up.gif"></td>
</tr>
<%

sql="select vcome,count(id) as allcome from view group by vcome order by count(id) DESC"
rs.Open sql,conn,1,1

maxallcome=0
sumallcome=0
do while not rs.EOF
	if cint(rs("allcome"))>maxallcome then maxallcome=cint(rs("allcome"))
	sumallcome=sumallcome+cint(rs("allcome"))
	rs.MoveNext
loop
	'防止除数为0而出错
	if maxallcome=0 then maxallcome=1
	if sumallcome=0 then sumallcome=1
rs.MoveFirst 

j=0
do while not rs.EOF
thecome=rs("vcome")
vallcome=rs("allcome")
	thelen=len(thecome)
	if thelen =0 then
		thecome="main.asp"
		svcome="通过收藏或直接输入网址访问"
	end if
	if thelen <= 33 and thelen > 0 then
		svcome=thecome
	end if
	if thelen >= 34 then
		svcome=left(thecome,31) & "..."
	end if
%>
<tr>
<td width="220" align=right><a href="<%=thecome%>" target="_blank"
	title="<%=thecome & vbcrlf%>访问<%=vallcome%>次，<%
	'计算访问量的百分数，精确到小数后1位，小于零的在前面加字母0
	lsbf=int(vallcome*1000/sumallcome)/10
	if lsbf<1 then lsbf="0" & lsbf
	Response.Write lsbf
	%>%"><%=svcome%></a>&nbsp;</td>
<td width="230" background="images/tu_back_2.gif" align=left>
<img style="BORDER-left: #000000 1px solid;" src="images/tu.gif"
	width="<%=(vallcome/maxallcome)*183%>" height="9"
	alt="<%=thecome & vbcrlf%>访问<%=vallcome%>次，<%
	'计算访问量的百分数，精确到小数后1位，小于零的在前面加字母0
	lsbf=int(vallcome*1000/sumallcome)/10
	if lsbf<1 then lsbf="0" & lsbf
	Response.Write lsbf
	%>%"> <%=vallcome%></td>
</tr>
<%
rs.MoveNext
	j=j+1
	'如果记录超过40条，就退出
	if j=40 then exit do
loop
%>
<tr height="10">
	<td width="220"></td><td width="230"><img src="images/tu_back_down.gif"></td>
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
%>