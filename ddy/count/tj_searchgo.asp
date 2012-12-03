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
if session.Contents("master")=false and whatcan<4 then Response.Redirect "help.asp?id=004&error=您没有使用自定义统计的权限。"

'********************************** 处理查询条件 **********************************

'================================
'
'     从浏览器或者数据并处理
'
'================================

'从浏览器获取数据

onyear		=trim(Request("onyear"))
onmonth		=trim(Request("onmonth"))
onday		=trim(Request("onday"))

offyear		=trim(Request("offyear"))
offmonth	=trim(Request("offmonth"))
offday		=trim(Request("offday"))

vip			=trim(Request("vip"))
vwhere		=trim(Request("vwhere"))
vos1		=trim(Request("vos1"))
vos2		=trim(Request("vos2"))
vsoft1		=trim(Request("vsoft1"))
vsoft2		=trim(Request("vsoft2"))
vcome		=trim(Request("vcome"))
vpage		=trim(Request("vpage"))

outtype		=trim(Request("outtype"))
wherestr	=trim(Request("wherestr"))
name		=trim(Request("name"))
content		=trim(Request("content"))

'处理获取的数据
ontime=onyear & "-" & onmonth & "-" & onday
if (ontime <> "--") and (not isdate(ontime)) then Response.Redirect "help.asp?id=001&error=日期填写不合乎要求。"

offtime=offyear & "-" & offmonth & "-" & offday
if (offtime <> "--") and (not isdate(offtime)) then Response.Redirect "help.asp?id=001&error=日期填写不合乎要求。"
if (offtime <> "--") then offtime=cdate(offtime)+1		'因为SQL总是按所指日期的0时计算的，所以应该改为下一日

if (outtype ="") then Response.Redirect "help.asp?id=001&error=您没有指明要查看的项目。"

vos=vos2
if vos = "" then vos=vos1
vsoft=vsoft2
if vsoft="" then vsoft=vsoft1


'************************************ 运算和输出 ************************************

'================================
'
'         转换为 SQL 指令
'
'================================

'将查询条件转换为SQL查询指令
if wherestr="" then
wherestr=" where "

'如果存在查询条件，则将查询条件加入查询字串
if ontime <> "--" then wherestr=wherestr & "and (vtime>=datevalue('" & ontime & "')) "
if offtime <> "--" then wherestr=wherestr & "and (vtime<=datevalue('" & offtime & "')) "
if vip <> "" then wherestr=wherestr & "and (vip like '%" & vip & "%') "
if vwhere <> "" then wherestr=wherestr & "and ((vwhere like '%" & vwhere & "%') or (vwheref like '%" & vwhere & "%')) "
if vos <> "" then vherestr=wherestr & "and (vos like '%" & vos & "%') "
if vsoft <> "" then wherestr=wherestr & "and (vsoft like '%" & vsoft & "%') "
if vcome <> "" then wherestr=wherestr & "and (vcome like '%" & vcome & "%') "
if vpage <> "" then wherestr=wherestr & "and (vpage like '%" & vpage & "%') "

'去除第一个查询条件和 where 之间的 and
wherestr=replace(wherestr,"where and","where")

'如果没有查询条件，则查询字串应为空
if trim(wherestr)="where" then wherestr=""
end if

'如果要查看的内容只有详细记录，则直接转入详细记录页
if outtype="详细" then Response.Redirect "tj_all.asp?wherestr=" & server.URLEncode(wherestr)

'================================
'
'          主 程 序
'
'================================

'打开数据库
set conn=server.createobject("adodb.connection")
DBPath = Server.MapPath(connpath)
conn.Open "driver={Microsoft Access Driver (*.mdb)};dbq=" & DBPath
Set rs = Server.CreateObject("ADODB.Recordset")

%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Copyright" content="Ajiang http://www.jhqing.com">
<title><%=countname%>－自定义统计</title>
<link rel="stylesheet" type="text/css" href="style.css">
</head>
<body topmargin=5 rightmargin=0 leftmargin=0 vlink=#000000>
<!--#include file="inc_top.asp"-->
<br>
<table width=500 cellspacing=0 align=center>
  <tr><td>
    <p style="line-height: 100%; margin-left: 15; margin-top: 5; margin-bottom: 0">
    恭喜: 下面显示的就是您要查询的内容的分析结果。</p>
  </td></tr>
</table>
<br>

<%
'根据查询条件输出相应的内容

if instr(outtype,"小时") then
statname="24小时访问分析"
tuwidth=15
thewhich="vhour"
tuhow=24
unit="时"
statheng statname,tuwidth,thewhich,tuhow,unit,wherestr
end if

if instr(outtype,"日") then
statname="日访问分析"
tuwidth=13
thewhich="vday"
tuhow=31
unit="日"
statheng statname,tuwidth,thewhich,tuhow,unit,wherestr
end if

if instr(outtype,"周") then
statname="周访问分析"
tuwidth=40
thewhich="vweek"
tuhow=7
unit=""
statheng statname,tuwidth,thewhich,tuhow,unit,wherestr
end if

if instr(outtype,"月") then
statname="月访问分析"
tuwidth=20
thewhich="vmonth"
tuhow=12
unit="月"
statheng statname,tuwidth,thewhich,tuhow,unit,wherestr
end if

if instr(outtype,"年") then
statname="年访问分析"
leftwidth=50
tuwhich="vyear"
statshu statname,leftwidth,tuwhich,wherestr
end if

if instr(outtype,"IP") then
statname="IP分析"
leftwidth=150
tuwhich="vIP"
statshu statname,leftwidth,tuwhich,wherestr
end if

if instr(outtype,"地区") then
statname="地区访问分析"
leftwidth=100
tuwhich="vwhere"
statshu statname,leftwidth,tuwhich,wherestr
end if

if instr(outtype,"浏览器") then
statname="不同浏览器用户访问分析"
tuwidth=50
thewhich="vsoft"
tuhow=7
unit=""
statheng statname,tuwidth,thewhich,tuhow,unit,wherestr
end if

if instr(outtype,"操作系统") then
statname="不同操作系统用户访问分析"
tuwidth=50
thewhich="vOS"
tuhow=7
unit=""
statheng statname,tuwidth,thewhich,tuhow,unit,wherestr
end if

if instr(outtype,"来路") then
statname="来路分析"
leftwidth=220
tuwhich="vcome"
statshu statname,leftwidth,tuwhich,wherestr
end if

if instr(outtype,"页面") then
statname="页面点击分析"
leftwidth=220
tuwhich="vpage"
statshu statname,leftwidth,tuwhich,wherestr
end if

'如果同时选择了详细，则提示点击继续
if instr(outtype,"详细") then
%>
<table width="500" cellspacing="0" align="center" cellpadding="0" border="0">
  <tr><td colspan="3"><img src="images/photoup.gif"></td></tr>
  <form action="tj_all.asp" method=post id=form1 name=form1>
  <tr height="30">
    <td width="1" class="backs"></td>
    <td width="498"class="backq">
		&nbsp; <img src="images/tb_title.gif" align=absmiddle><font style="font-size:16px">&nbsp;</font>∷∷∷ 查看详细记录 ∷∷∷<br>
		<p class="p1">因为详细记录需要翻页，不能和以上记录放在同一页，请点击下面的继续按钮查看相应内容。
	<p class="p1" style='margin-top: 0;' align="right">
	<a href='javascript:document.forms[0].submit();'>继续</a>
	<input type="hidden" name="wherestr" value="<%=wherestr%>">
	<input type="submit" value=" " name="save" class="backc2"><font style="font-size:16px">&nbsp;</font>&nbsp;
	</td>
    <td width="1" class="backs"></td>
  </tr>
  </form>
  <tr><td colspan="4"><img src="images/photodown.gif"></td></tr>
</table>
<br>
<%
end if

'保存数据表单
if session.Contents("master")=true or whatcan>=6 then
%>
<table width="500" cellspacing="0" align="center" cellpadding="0" border="0">
  <tr><td colspan="3"><img src="images/photoup.gif"></td></tr>
  <form action="tj_save.asp" method=post>
  <tr height="30">
    <td width="1" class="backs"></td>
    <td width="498"class="backq">
		&nbsp; <img src="images/tb_save.gif" align=absmiddle><font style="font-size:16px">&nbsp;</font>∷∷∷ 保存这次检索条件 ∷∷∷<br>

	<p class="p1"><font class="fonts"><b>检索条件</b></font>&nbsp; <%if wherestr="" then%>没有检索条件<%else%><%=wherestr%><%end if%>。
	<p class="p1" style='margin-bottom: 0;margin-top: 0;'><font class="fonts"><b>查询项目</b></font>&nbsp; <%=outtype%>。
	<p class="p1" style='line-height: 100%;margin-bottom: 0;margin-top: 0;'><font class="fonts"><b>取个名字</b></font>&nbsp; <input name="name" size="16" class="input" value="<%=name%>">&nbsp;
	<INPUT type="radio" name="overwrite" value="0"<%if name="" then%> checked<%end if%>> 同名时提示&nbsp;
	<INPUT type="radio" name="overwrite" value="1"<%if name<>"" then%> checked<%end if%>> 同名时覆盖
	<p class="p1" style='line-height: 100%;margin-bottom: 0;margin-top: 0;'><font class="fonts"><b>加个介绍</b></font>&nbsp; <input name="content" size="50" class="input" value="<%=content%>">
	<p class="p1" style='margin-top: 0;' align="right">
	<a href="help.asp?id=002">帮助</a>
	<a href='javascript:document.forms[0].submit();'>保存</a>
	<input type="hidden" name="wherestr" value="<%=wherestr%>"><input type="hidden" name="outtype" value="<%=outtype%>"><input type="submit" value=" " name="save" class="backc2"><font style="font-size:16px">&nbsp;</font>&nbsp;

	</td>
    <td width="1" class="backs"></td>
  </tr>
  </form>
  <tr><td colspan="4"><img src="images/photodown.gif"></td></tr>
</table>
<br>
<%
end if
%>
<!--#include file="inc_bottom.asp"-->
</body>
</html>
<%
' 关闭数据库
set rs=nothing
conn.Close
set conn=nothing
%>





<%

'******************************** 自定义函数和子程序 ********************************

'================================
'
'   输出图表数据的子程序(横)
'
'================================
sub statheng(statname,tuwidth,tuwhich,thehow,unit,wherestr)

'statname	报告名称
'tuwidth	横图每栏宽度
'thewhich	要查询的项目
'thehow		共有多少种数据
'unit		项目的单位
'wherestr	查询条件

'声明输出内容数组变量
dim val(),alt(),theper()
redim val(thehow),alt(thehow),theper(thehow)

'开始计算
sql="select " & thewhich & ",count(id) as theval from view " & wherestr & _
	" group by " & thewhich & " order by " & thewhich
rs.Open sql,conn,1,1

themax=0
thesum=0

for i=0 to tuhow-1
	if not rs.EOF then
		alt(i)=rs(thewhich)
		val(i)=rs("theval")
		if cint(val(i))>themax then themax=cint(val(i))
		thesum=thesum + cint(val(i))
		rs.MoveNext
	else
		alt(i)=""
		val(i)=0
	end if	
next
'防止除数为0而出错
if themax=0 then themax=1
if thesum=0 then thesum=1

'计算百分数
for i=0 to tuhow-1
	theper(i)=int(val(i)*1000/thesum)/10
	if csng(theper(i))<1 then theper(i)="0" & theper(i)
next

'如果查询项目是星期，则转alt为汉字
if thewhich="vweek" then
  for i=0 to tuhow-1
    alt(i)=findweek(alt(i))
  next
end if
rs.Close

'调用hengout子程序输出到浏览器
hengout statname,tuwidth,themax,tuhow,val,alt,theper,unit

end sub





'================================
'
'   显示分析图表的子程序(横)
'
'================================

sub hengout(statname,tuwidth,themax,tuhow,val,alt,theper,unit)

'statname	报告名称
'tuwidth	横图每栏宽度
'themax		x坐标最大值
'tuhow		柱图有多少栏
'val()		x坐标值数组
'alt()		x坐标文字
'theper()	柱图百分数
'unit		x坐标单位
%>
<table width="500" cellspacing="0" align="center" cellpadding="0" border="0">
  <tr><td colspan="3"><img src="images/photoup.gif"></td></tr>
  <tr height="30">
    <td width="1" class="backs"><img src="images/touming.gif" width="1" height="1"></td>
    <td width="498"class="backq">
    
		&nbsp; <img src="images/tb_title.gif" align=absmiddle> &nbsp;∷∷∷ <%=statname%> ∷∷∷<br>
<table border="0" cellpadding="0" cellspacing="0" width="<%=tuhow * tuwidth + 70%>" align=center>
<tr height="9"><td></td></tr>
<tr height="101">
  <td align=right valign=top>
	<p style="line-height: 100%; margin-right: 2; margin-top: 0; margin-bottom: 13"><font face="Arial"><%=int(themax*10+0.5)/10%></font>
	<p style="line-height: 100%; margin-right: 2; margin-top: 0; margin-bottom: 13"><font face="Arial"><%=int(3*themax*10/4+0.5)/10%></font>
	<p style="line-height: 100%; margin-right: 2; margin-top: 0; margin-bottom: 13"><font face="Arial"><%=int(themax*10/2+0.5)/10%></font>
	<p style="line-height: 100%; margin-right: 2; margin-top: 0; margin-bottom: 0"><font face="Arial"><%=int(themax*10/4+0.5)/10%><br></font></td>
  <td width=10><img src="images/tu_back_left.gif"></td>
<%
for i= 0 to tuhow-1
%>
  <td width="<%=tuwidth%>" valign=bottom background="images/tu_back.gif" align=center>
  <img style="BORDER-BOTTOM: #000000 1px solid;" src="images/tu.gif"
	height="<%=(val(i)/themax)*100%>" width="9"
	alt="<%=alt(i)%><%=unit%>，访问<%=val(i)%>次，<%=theper(i)%>%"></td>
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
for i= 0 to tuhow-1
%>
  <td width="<%=tuwidth%>" align=center><a
	title="<%=alt(i)%><%=unit%>，访问<%=val(i)%>次，<%=theper(i)%>%"><font
	face="Arial" style="letter-spacing: -1"><%=alt(i)%></font></a></td>
<%
next
%>
  <td width=10></td>
  <td width=10></td>
</tr>
<tr height="5"><td></td></tr>
</table>

	</td>
    <td width="1" class="backs"><img src="images/touming.gif" width="1" height="1"></td>
  </tr>
  <tr><td colspan="4"><img src="images/photodown.gif"></td></tr>
</table>
<br>
<%
end sub
%>

<%

'================================
'
'   输出图表数据的子程序(竖)
'
'================================

sub statshu(statname,leftwidth,thewhich,wherestr)

'statname	报告名称
'leftwidth	横图每栏宽度
'thewhich	要查询的项目
'wherestr	查询条件

'声明输出内容数组变量
dim val(),alt(),theper(),alta()

on error resume next
'开始计算
sql="select " & thewhich & ",count(id) as theval from view " & wherestr & _
	" group by " & thewhich & " order by count(id) DESC"
rs.Open sql,conn,1,1

themax=0
thesum=0
tuhow=rs.recordcount
if tuhow>100 then tuhow=100

redim val(tuhow),alt(tuhow),theper(tuhow),alta(tuhow)
i=0
do while not rs.EOF
	alta(i)=rs(thewhich)
	if (thewhich = "vcome") or (thewhich = "vpage") then
		thelen=len(alta(i))
		if thelen =0 then
			alta(i)="#"
			alt(i)="通过收藏或直接输入网址访问"
		end if
		if thelen <= 33 and thelen > 0 then
			alt(i)=alta(i)
		end if
		if thelen >= 34 then
			alt(i)=left(alta(i),31) & "..."
		end if
	else
		alt(i)=alta(i)
	end if
	val(i)=rs("theval")
	if cint(val(i))>themax then themax=cint(val(i))
	thesum=thesum + cint(val(i))
	if i=100 then exit do
	i=i+1
	rs.MoveNext
loop
'防止除数为0而出错
if themax=0 then themax=1
if thesum=0 then thesum=1

'计算百分数
for i=0 to tuhow-1
	theper(i)=int(val(i)*1000/thesum)/10
	if csng(theper(i))<1 then theper(i)="0" & theper(i)
next

'如果查询项目是星期，则转alt为汉字
if thewhich="vweek" then
  for i=0 to tuhow-1
    alt(i)=findweek(alt(i))
  next
end if
rs.Close

'调用hengout子程序输出到浏览器
shuout statname,leftwidth,themax,tuhow,val,alt,alta,theper

end sub






'================================
'
'   显示分析图表的子程序(竖)
'
'================================

sub shuout(statname,leftwidth,themax,tuhow,val,alt,alta,theper)
'statname	报告名称
'leftwidth	图表左边留给文字的宽度
'themax		x坐标最大值
'tuhow		多少行记录
'val()		记录的值(访问数)
'alt()		说明文字
'alta()		文字的连接
'theper()	柱图百分数
%>

<table width="500" cellspacing="0" align="center" cellpadding="0" border="0">
  <tr><td colspan="3"><img src="images/photoup.gif"></td></tr>
  <tr height="30">
    <td width="1" class="backs"><img src="images/touming.gif" width="1" height="1"></td>
    <td width="498"class="backq">
    
		&nbsp; <img src="images/tb_title.gif" align=absmiddle> &nbsp;∷∷∷ <%=statname%> ∷∷∷<br>

<table border="0" cellpadding="0" cellspacing="0" width="<%=leftwidth+230%>" align=center>
  <tr height="9"><td></td></tr>
  <tr height="10">
    <td width="<%=leftwidth%>"></td><td width="230"><img src="images/tu_back_up.gif"></td>
  </tr>
<%for i=0 to tuhow-1%>
  <tr>
    <td width="220" align=right><a href="<%=alta(i)%>" target="_blank"
	title="<%=alta(i) & vbcrlf%>访问<%=val(i)%>次，<%=theper(i)%>%"><%=alt(i)%></a>&nbsp;</td>
  <td width="230" background="images/tu_back_2.gif" align=left>
  <img style="BORDER-left: #000000 1px solid;" src="images/tu.gif"
	width="<%=(val(i)/themax)*183%>" height="9"
	alt="<%=alta(i) & vbcrlf%>，访问<%=val(i)%>次，<%=theper(i)%>%"> <%=val(i)%></td>
</tr>
<%
next
%>
<tr height="10">
	<td width="220"></td><td width="230"><img src="images/tu_back_down.gif"></td>
</tr>
<tr height="5"><td></td></tr>
</table>

	</td>
    <td width="1" class="backs"><img src="images/touming.gif" width="1" height="1"></td>
  </tr>
  <tr><td colspan="4"><img src="images/photodown.gif"></td></tr>
</table>
<br>
<%
end sub


'================================
'
'   将星期转为汉字的自定义函数
'
'================================
function findweek(theweek)
	select case cstr(theweek)
	case "1"
		findweek="星期日"
	case "2"
		findweek="星期一"
	case "3"
		findweek="星期二"
	case "4"
		findweek="星期三"
	case "5"
		findweek="星期四"
	case "6"
		findweek="星期五"
	case "7"
		findweek="星期六"
	end select
end function
%>