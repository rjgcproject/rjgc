<table width="500" cellspacing="0" align="center" cellpadding="0" border="0">
  <tr><td colspan="3"><img src="images/photoup.gif"></td></tr>
  <tr height="30">
    <td width="1" class="backs"></td>
    <td width="498"class="backq">
		&nbsp; <img src="images/tb_index.gif" align=absmiddle> &nbsp;∷∷∷ <%=countname%> ∷∷∷
		<p style="line-height: 140%; margin-left: 25; margin-right: 10; margin-top: 10; margin-bottom: 5" align=left>
		<a href="main.asp">◇ 总体数据</a>
		<a href="tj_hour.asp">◇ 24小时</a>
		<a href="tj_day.asp">◇ 日统计</a>
		<a href="tj_month.asp">◇ 月/年统计&nbsp;</a>
		<a href="tj_week.asp">◇ 星期统计</a>
		<a href="tj_come.asp">◇ 来路</a>
		<a href="tj_page.asp">◇ 页面</a>
		<br>
		<a href="tj_ip.asp">◇ 访问者IP</a>
		<a href="tj_soft.asp">◇ 浏览器和操作系统</a>
		<a href="tj_where.asp">◇ 访问者地区</a>
		<%if session.Contents("master")=true or whatcan>2 then%><a href="tj_all.asp">◇ 详细记录</a><%end if%>
		<%if session.Contents("master")=true or whatcan>3 then%><a href="tj_search.asp">◇ 自定义统计</a><%end if%>
	</td>
    <td width="1" class="backs"></td>
  </tr>
  <tr><td colspan="4"><img src="images/photodown.gif"></td></tr>
</table>
