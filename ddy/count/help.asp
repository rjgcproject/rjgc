<%@ CODEPAGE = "936" %>
<!--#include file="inc_config.asp"-->
<%
'从浏览器获取参数
id=Request("id")
error=Request("error")
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Copyright" content="Ajiang http://www.jhqing.com">
<title><%=countname%>－帮助信息</title>
<link rel="stylesheet" type="text/css" href="style.css">
</head>
<body topmargin=5 rightmargin=0 leftmargin=0 vlink=#000000>
<!--#include file="inc_top.asp"-->
<br>
<table width="500" cellspacing="0" align="center" cellpadding="0" border="0">
  <tr><td colspan="3"><img src="images/photoup.gif"></td></tr>
  <tr height="30">
    <td width="1" class="backs"><img src="images/touming.gif" width="1" height="1"></td>
    <td width="498"class="backq">
		&nbsp; <img src="images/tb_help.gif" align=absmiddle> <font style="font-size:16px">&nbsp;</font>帮 助 信 息<br>
	<%if error <> "" then Response.Write "<p class='p1'><font color='red'><img src='images/tb_error.gif' align=absmiddle><font style='font-size:15px'>&nbsp;</font>错误 " & error & "</font>"%>
<%select case id
case ""%>
	<li style="	margin-left: 50; margin-top: 15;margin-bottom: 5"><a href="help.asp?id=001">自定义统计条件应该怎样填写呢？</a>
	<li style="	margin-left: 50; margin-top: 5;	margin-bottom: 5"><a href="help.asp?id=002">如何保存自定义检索分析的检索条件？</a>
	<li style="	margin-left: 50; margin-top: 5;	margin-bottom: 5"><a href="help.asp?id=003">如何修改或删除已经保存的检索条件？</a>
	<li style="	margin-left: 50; margin-top: 5;	margin-bottom: 5"><a href="help.asp?id=004">访客和管理员分别具有哪些权限？</a>
	<li style="	margin-left: 50; margin-top: 5;	margin-bottom: 5"><a href="help.asp?id=005">如何获得本系统的最新版本？</a>
	<li style="	margin-left: 50; margin-top: 5;	margin-bottom: 5"><a href="help.asp?id=006">本系统的版权声明。</a>
	
<%case "001"%>
	<p class="p1"><font class="fonts"><b>自定义统计条件应该怎样填写呢？</b></font>
	<p class="p1">起止日期：年、月、日都必须填写，而且都必须是数字。但是可以只填起始日期，这样表示从填写的起始日期至现在这个时间段，同样也可以只填写截止日期，这样表示从开始统计那天至所填写的这个日期，注意截止日期必须在开通统计日期之后。
	<p class="p1">IP地址和来自地区：这两项都是支持模糊查询的，例如在IP地址项中输入“61.”，则可查到IP地址是“61.163.233.10”或者“202.61.33.25”这样的访问记录的统计分析。
	<p class="p1">操作系统和浏览器：这两项可以从列表中选择，也可以直接在后边文本框中输入，也支持模糊查询，比如在文本框中输入win，则可以查到WIN 2K、WIN XP、WIN 9X的全部记录的统计分析。
	<p class="p1">来自和被访网页：这两个选项可以让您查到特定来路或目标的访问记录的统计分析，支持模糊查询。
	<p class="p1">输出类型：可以选择要查看的分析的类型，可多选，但要注意您选择的越多，需要等待的时间就越长，当然，您也不能一个都不选。

<%case "002"%>
	<p class="p1"><font class="fonts"><b>如何保存自定义检索分析的检索条件？</b></font>
	<p class="p1">在进行了自定义统计的检索之后，在该页面底部有一个名为“保存本次检索条件”的项目框，该框中的选项允许你为要保存的条件取名并加入简介。
	<p class="p1">保存前必须为该条件取一个名字，否则日后将无法使用您保存的检索条件。
	<p class="p1">最好再为要保存的检索条件写一个简介，因为名字写的太长会影响页面美观的，要细致得看懂名字所代表的意思，写个简介是必要的。
	<p class="p1">在名字的输入框后面有两个单选框，可以让您指明当和已有的检索项目重名是覆盖掉原有的呢还是让系统提示重新取名字，除非您本来就是为了修改已有项目并确认名字没写错的话，否则一定保持它选择“同名时提问”，以确保已有数据的安全。
	<p class="p1">只有管理员才可以进行上述操作。
	
<%case "003"%>
	<p class="p1"><font class="fonts"><b>如何修改或删除已经保存的检索条件？</b></font>
	<p class="p1">［修改］您只需要按照新的检索条件进行检索，在检索结果页面的底部的保存项中使用原有的名字保存，并选择“同名时覆盖”，保存后就修改了原来保存的项目。
	<p class="p1">［删除］在“自定义统计”页面的自定义检索条件列表中点击“删除”连接，进入删除页面，系统提示是否真的要删除，点击确定即可。
	<p class="p1">只有管理员才可以进行上述操作。

<%case "004"%>
	<p class="p1"><font class="fonts"><b>访客和管理员分别具有哪些权限？</b></font>
	<p class="p1">亦凡酷站访问统计对访问权限实行等级管理，管理员始终具有 6 级的权力，即拥有所有权力，非管理员（访客）所拥有的权限由管理员在安装时在配置文件中设置。等级划分的办法是：</p>
<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="320" align=center>
	<tr height="18" align=center class=backzq><td width="110">等 级</td><td width="30">0</td><td width="30">1</td><td width="30">2</td><td width="30">3</td><td width="30">4</td><td width="30">5</td><td width="30">6</td></tr>
	<tr height="18" align=center><td align=left>&nbsp;查看总体数据</td><td></td><td>√</td><td>√</td><td>√</td><td>√</td><td>√</td><td>√</td></tr> 
	<tr height="18" align=center class=backzq><td align=left>&nbsp;查看缺省统计数据</td><td></td><td></td><td>√</td><td>√</td><td>√</td><td>√</td><td>√</td></tr> 
	<tr height="18" align=center><td align=left>&nbsp;查看详细记录</td><td></td></td><td></td><td><td>√</td><td>√</td><td>√</td><td>√</td></tr> 
	<tr height="18" align=center class=backzq><td align=left>&nbsp;查看自定义记录</td><td></td><td></td><td></td><td></td><td>√</td><td>√</td><td>√</td></tr> 
	<tr height="18" align=center><td align=left>&nbsp;自定义条件统计</td><td></td><td></td><td></td><td></td><td></td><td>√</td><td>√</td></tr> 
	<tr height="18" align=center class=backzq><td align=left>&nbsp;管理自定义库</td><td></td><td></td><td></td><td></td><td></td><td></td><td>√</td></tr> 
</table>
	<p class="p1">如果访客等级被设置为 0，则访客将没有任何权力，如果访客等级被设置为 6，则访客将具有和管理员相等的权力。默认的访客权限等级为 4。
	<p class="p1">当前访客的权限等级为：<%=whatcan%>

<%case "005"%>
	<p class="p1"><font class="fonts"><b>如何获得本系统的最新版本？</b></font>
	<p class="p1">请您关注我的主页 <a href="http://www.jhqing.com" target="_blank">http://www.jhqing.com</a> 我会在推出新版的第一时间在主页上发布。您也将从亦凡的主页上获得亦凡编写的其他程序。
	<p class="p1">请务必注意保留亦凡的版权信息，我花费很多精力开发各种程序给大家免费使用，会有一种成就感，我不想别人窃取这种荣耀。
	<p class="p1">最后感谢您使用亦凡编写的程序，希望可以得到您的建议和继续支持。

<%case "006"%>
	<p class="p1"><font class="fonts"><b>本系统的版权信息</b></font>
	<p class="p1">本系统版权属于亦凡(zhqi123@163.com)所有，以下情况是允许的：
	<br>　　1、在个人或商业网站上使用本程序并保留亦凡的版权信息；
	<br>　　2、在保持原压缩包完整性的情况下重新分发本程序；
	<br>　　3、修改本程序，如果修改量小于三分之一，请注明“修改自亦凡版网站统计系统”字样并保留亦凡信箱及网站连接；
	<br>　　4、修改本程序超过三分之一的，可以删除有关亦凡的任何信息并重新分发，但在分发前应通知亦凡；
	<br>　　5、在您的程序中引用本程序中代码片断，但请注明“引用自亦凡版网站统计系统”。
	<p class="p1">以下情况是不允许的：
	<br>　　1、修改程序不足三分之一且修改版权重新分发；
	<br>　　2、在您的网站上使用本系统全没有保留亦凡版权信息及连接；
	<br>　　3、直接或间接使用本程序赚取利润；
	<p class="p1">本人允许个人用户使用本程序，但不允许更改版权后传播本程序。曾发现有人修改我的探针程序（仅仅修改了版权）并在自己的主页上发布，说是自己创作的，对于这种情况，亦凡感到非常遗憾，希望这种事情不要再发生在我的统计系统上。
	<p class="p1">本人也不反对将本程序用做商业网站，但任何人都不能从中获利，比如网站建设公司将本系统用在给客户制作的网站上时不能因为本程序而向客户增加费用，也不得将“免费赠送”本系统作为和客户谈判的一个条件。
	<p class="p1">最后感谢您使用亦凡编写的程序，希望可以得到您的建议和继续支持。
	
<%case else%>
<p class="p1">目前还没有与此相关的帮助信息。

<%end select%>
<p class="p1" align="right"><a href='javascript:history.back()'>继续</a> <a href='javascript:history.back()'><img src="images/arbutton.gif" align="absmiddle" border="0"></a> <font style="font-size:16px">&nbsp;</font>&nbsp;
	</td>
    <td width="1" class="backs"><img src="images/touming.gif" width="1" height="1"></td>
  </tr>
  <tr><td colspan="4"><img src="images/photodown.gif"></td></tr>
</table>
<br>
<!--#include file="inc_bottom.asp"-->
</body>
</html>
