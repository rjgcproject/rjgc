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
if session.Contents("master")=false and whatcan<4 then Response.Redirect "help.asp?id=004&error=您没有查看自定义检索记录和进行自定义检索的的权限。"
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
    Tips: 请在下面表格中填写您的查询条件，支持组合查询和模糊查询。</p>
  </td></tr>
</table>
<br>
<%
'检查权限
if session.Contents("master")=true or whatcan>=5 then
%>
<table width="500" cellspacing="0" align="center" cellpadding="0" border="0">
  <tr><td colspan="3"><img src="images/photoup.gif"></td></tr>
  <tr height="30">
    <td width="1" class="backs"></td>
    <td width="498"class="backq">
<%
dim OS(7)
	OS(0)="Win 2000"
	OS(1)="Win XP"
	OS(2)="Win NT"
	OS(3)="Win 9x"
	OS(4)="类Unix"
	OS(5)="Mac"
	OS(6)="Other"

dim soft(7)
	soft(0)="NetCaptor"
	soft(1)="MSIE 6.x"
	soft(2)="MSIE 5.x"
	soft(3)="MSIE 4.x"
	soft(4)="Netscape"
	soft(5)="Opera"
	soft(6)="Other"
%>
		&nbsp; <img src="images/tb_title.gif" align=absmiddle> &nbsp;∷∷∷ 自定义统计项目 ∷∷∷<br>
<table border="0" cellpadding="0" cellspacing="0" width="430" align=center>
<form action="tj_searchgo.asp" method="post" name="search">
<tr height="9"><td></td></tr>
<tr><td>
	时 间 段：<input name="onyear" class="input" size="5"> 年
			  <input name="onmonth" class="input" size="3"> 月
			  <input name="onday" class="input" size="3"> 日
		   ～ <input name="offyear" class="input" size="5"> 年
			  <input name="offmonth" class="input" size="3"> 月
			  <input name="offday" class="input" size="3"> 日
<br>ＩＰ地址：<input name="vip" class="input" size="16">
<br>来自地区：<input name="vwhere" class="input" size="25">
<br>客 户 端：<SELECT name="vOS1" class="input" style="width:70" id="vos1">			  <OPTION value="" selected>操作系统</OPTION><%for i=0 to 6%>			  <OPTION value="<%=OS(i)%>"><%=OS(i)%></OPTION><%next%>
			  </SELECT>
			  <INPUT name="vOS2" class="input" size="10" id="vos">&nbsp; 

			  <SELECT name="vsoft1" class="input" style="width:70">			  <OPTION value="" selected>浏览器</OPTION><%for i=0 to 6%>			  <OPTION value="<%=soft(i)%>"><%=soft(i)%></OPTION><%next%>
			  </SELECT>
			  <INPUT name="vsoft2" class="input" size="10">

<br>来自网页：<input name="vcome" class="input" size="35">
<br>被访网页：<input name="vpage" class="input" size="35">
<br>输出类型：<INPUT type="checkbox" name="outtype" value="小时">小时　　
				<INPUT type="checkbox" name="outtype" value="日">日　　　
				<INPUT type="checkbox" name="outtype" value="周">周　　　
				<INPUT type="checkbox" name="outtype" value="月">月
				<br>　　　　　<INPUT type="checkbox" name="outtype" value="年">年　　　
				<INPUT type="checkbox" name="outtype" value="IP">IP　　　
				<INPUT type="checkbox" name="outtype" value="地区">地区　　
				<INPUT type="checkbox" name="outtype" value="浏览器">浏览器　
				<br>　　　　　<INPUT type="checkbox" name="outtype" value="操作系统">操作系统
				<INPUT type="checkbox" name="outtype" value="来路">来路　　
				<INPUT type="checkbox" name="outtype" value="页面">页面　　
				<INPUT type="checkbox" name="outtype" value="详细">详细
</td></tr>
<tr><td align=right><a href="help.asp?id=001">帮助</a> <a href='javascript:document.forms[0].submit();'>查看结果</a> <input type="submit" name="search" class="backc2" value=" "> &nbsp;</td></tr>
</form>
</table>

	</td>
    <td width="1" class="backs"></td>
  </tr>
  <tr><td colspan="4"><img src="images/photodown.gif"></td></tr>
</table>
<br>
<%
end if		'检查权限

'显示原来保存的自定义检索
'打开数据库
set conn=server.createobject("adodb.connection")
DBPath = Server.MapPath(connpath)
conn.Open "driver={Microsoft Access Driver (*.mdb)};dbq=" & DBPath
Set rs = Server.CreateObject("ADODB.Recordset")

sql="select * from save"
rs.Open sql,conn,1,1

if not rs.EOF then
%>
<table width="500" cellspacing="0" align="center" cellpadding="0" border="0">
  <tr><td colspan="3"><img src="images/photoup.gif"></td></tr>
  <tr height="30">
    <td width="1" class="backs"><img src="images/touming.gif" width="1" height="1"></td>
    <td width="498"class="backq">
		&nbsp; <img src="images/tb_title.gif" align=absmiddle><font style="font-size:16px">&nbsp;</font>∷∷∷ 已保存的检索条件 ∷∷∷<br><br>

<table width="430" border="0" cellspacing=0 cellpadding=0 align=center>
  <tr height="1" class="backs"><td colspan=2><img src="images/touming.gif" width="1" height="1"></td></tr>
<%do while not rs.EOF%>
  <form action="tj_searchgo.asp" method=post>
  <tr height="28">
    <td>
      □ <font class=fonts><%=rs("name")%></font>
    </td>
    <td align=right>
      <input name="content" size="50" type="hidden" value="<%=rs("content")%>"
      ><input type="hidden" name="wherestr" value="<%=rs("wherestr")%>"
      ><input type="hidden" name="outtype" value="<%=rs("outtype")%>"
      ><input type="hidden" name="name" value="<%=rs("name")%>"
      ><%if session.Contents("master")=true or whatcan>=6 then%><a href="tj_save_del.asp?id=<%=rs("id")%>">删除</a> <%end if%>查看结果 <input type="submit" value=" " name="save" class="backc2"> &nbsp;
    </td>
  </tr>
  <%if rs("content") <> "" then%>
  <tr>
    <td colspan=2><p style='text-indent: 25;line-height: 130%;margin-left: 17;margin-right: 5;margin-top: 0;	margin-bottom: 5'><font color="#555555"><%=rs("content")%></font></td>
  </tr>
  <%end if%>
  <tr height="1" class="backs"><td colspan=2><img src="images/touming.gif" width="1" height="1"></td></tr>
  </form>
<%
rs.MoveNext
loop%>
</table>
<br>
	</td>
    <td width="1" class="backs"><img src="images/touming.gif" width="1" height="1"></td>
  </tr>
  <tr><td colspan="4"><img src="images/photodown.gif"></td></tr>
</table>
<br>
<%
end if

rs.Close
set rs=nothing
conn.Close
set conn=nothing
%>
<!--#include file="inc_bottom.asp"-->
</body>
</html>