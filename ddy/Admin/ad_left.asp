<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%Const dbdns="../"%>
<!--#include file="chk.asp"-->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>后台管理系统</title>
<link rel="stylesheet" type="text/css" href="images/Style.css">

<style type="text/css">
<!--
.ttl {	CURSOR: hand; COLOR: #000000; PADDING-TOP: 4px
}
BODY {
	MARGIN-TOP: 5px; MARGIN-LEFT: 2px; BACKGROUND-COLOR: #fda700; TEXT-ALIGN: center
}
td{
	line-height:170%;
}
-->
</style>

<script language="javascript">
function showHide(obj)
{
	obj.style.display = obj.style.display == "none" ? "block" : "none";
}
</script>
</head>

<body>
<table cellspacing="1" cellpadding="2" width="150" align="center" bgcolor="#999999" 
border="0">
  <tbody>
    <tr>
      <td class="ttl" onClick="showHide(m0)" valign="top" align="left" background="images/top-bj3.jpg"><table cellspacing="0" cellpadding="0" width="127" border="0">
        <tbody>
          <tr>
            <td width="8">&nbsp;</td>
            <td align="left" width="117"><strong class="mtitle">常规操作</strong></td>
          </tr>
        </tbody>
      </table></td>
    </tr>
    <tr id="m0" style="display: block">
      <td valign="top" align="middle" bgcolor="#f3f5f1"><table width="100%" cellpadding="2">
        <tbody>
          
          <tr>
            <td align="left"><img height="7" hspace="5" src="images/arrow.gif" width="5" align="absmiddle" /><a  href="ad_right.asp" target="right">管理首页</a></td>
          </tr>
          <tr>
            <td align="left"><img height="7" hspace="5" src="images/arrow.gif" width="5" align="absmiddle" /><a  href="pass.asp" target="right">修改密码</a></td>
          </tr>
          
          <tr>
            <td align="left"><img height="7" hspace="5" src="images/arrow.gif" width="5" align="absmiddle" /><a  href="quit.asp" target="_parent" onClick="if(!confirm('真的要退出吗?')) return false;">安全退出</a></td>
          </tr>
        </tbody>
      </table></td>
    </tr>
  </tbody>
</table>
<br>
<%If InStr(Admin.GroupId,",1,")>=1 Then%>
<table cellspacing="1" cellpadding="2" width="150" align="center" bgcolor="#999999" 
border="0">
  <tbody>
    <tr>
      <td class="ttl" onClick="showHide(menu6)" valign="top" align="left" background="images/top-bj3.jpg"><table cellspacing="0" cellpadding="0" width="127" border="0">
        <tbody>
          <tr>
            <td width="8">&nbsp;</td>
            <td align="left" width="117"><strong class="mtitle">系统管理</strong></td>
          </tr>
        </tbody>
      </table></td>
    </tr>
    <tr id="menu6" style="display: block">
      <td valign="top" align="middle" bgcolor="#f3f5f1"><table width="100%" cellpadding="2">
        <tbody>
          
          <tr>
            <td align="left"><img height="7" hspace="5" src="images/arrow.gif" width="5" align="absmiddle" /><a  href="Sys_Config.asp" target="right">站点信息配置</a></td>
          </tr>
          <tr>
            <td align="left"><img height="7" hspace="5" src="images/arrow.gif" width="5" align="absmiddle" /><a  href="Sys_admin.asp" target="right">管理员管理</a></td>
          </tr>
		  <%If Db_Type="ACC" Then%>
          <tr>
            <td align="left"><img height="7" hspace="5" src="images/arrow.gif" width="5" align="absmiddle" /><a  href="Sys_db.asp" target="right">数据库管理</a></td>
          </tr>
		  <%End If%>
		   <tr>
		     <td align="left"><img height="7" hspace="5" src="images/arrow.gif" width="5" align="absmiddle" /><a  href="ad_weblog.asp" target="right">后台日志</a></td>
	        </tr>
        </tbody>
      </table></td>
    </tr>
  </tbody>
</table>
<br />
<%End If%>
<%If InStr(Admin.GroupId,",3,")>=1 Then%>
<table cellspacing="1" cellpadding="2" width="150" align="center" bgcolor="#999999" 
border="0">
  <tbody>
    <tr>
      <td class="ttl" onClick="showHide(m2)" valign="top" align="left" background="images/top-bj3.jpg"><table cellspacing="0" cellpadding="0" width="127" border="0">
        <tbody>
          <tr>
            <td width="8">&nbsp;</td>
            <td align="left" width="117"><span class="mtitle">文章管理</span></td>
          </tr>
        </tbody>
      </table></td>
    </tr>
    <tr id="m2" style="display: block">
      <td height="112" align="middle" valign="top" bgcolor="#f3f5f1"><table width="100%" cellpadding="2">
        <tbody>
          <tr>
            <td align="left"><img height="7" hspace="5" src="images/arrow.gif" width="5" align="absmiddle" /><a  href="Class_Manage.asp?ChannelId=1&amp;ParentID=0&amp;Depth=0" target="right">新闻分类</a></td>
          </tr>
          <tr>
            <td align="left"><img height="7" hspace="5" src="images/arrow.gif" width="5" align="absmiddle" /><a  href="Article_Edit.asp?ChannelId=1" target="right">添加新闻</a> <a href="Caiji/" target="_blank" style="color:#0000FF;"></a></td>
          </tr>
          <tr>
            <td align="left"><img height="7" hspace="5" src="images/arrow.gif" width="5" align="absmiddle" /><a  href="Article_List.asp?ChannelId=1" target="right">管理新闻</a> <a href="Article_List.asp?ChannelId=1&IsPass=0" target="right" style="color:#0000FF;">待审</a> <a  href="Article_Move.asp?ChannelId=1" target="right">转移</a></td>
          </tr>
		  <tr>
		    <td align="left"><img height="7" hspace="5" src="images/arrow.gif" width="5" align="absmiddle" /><a  href="Article_List.asp?ChannelId=1&amp;IsDelete=1" target="right">回收站</a></td>
		    </tr>
        </tbody>
      </table></td>
    </tr>
  </tbody>
</table>
<br />
<%End If%>
<%If InStr(Admin.GroupId,",6,")>=1 Then%>
<br />
<%End If%>
<%If InStr(Admin.GroupId,",4,")>=1 Then%>
<table cellspacing="1" cellpadding="2" width="150" align="center" bgcolor="#999999" 
border="0">
  <tbody>
    <tr>
      <td class="ttl" onClick="showHide(m3)" valign="top" align="left" background="images/top-bj3.jpg"><table cellspacing="0" cellpadding="0" width="127" border="0">
        <tbody>
          <tr>
            <td width="8">&nbsp;</td>
            <td align="left" width="117"><span class="mtitle">在线留言管理</span></td>
          </tr>
        </tbody>
      </table></td>
    </tr>
    <tr id="m3" style="display: block">
      <td valign="top" align="middle" bgcolor="#f3f5f1"><table width="100%" cellpadding="2">
        <tbody>
          <tr>
            <td align="left"><img height="7" hspace="5" src="images/arrow.gif" width="5" align="absmiddle" /><a  href="../gbook/admin.asp" target="right" style="color:#0000FF;">留言管理</a></td>
          </tr>
        </tbody>
      </table></td>
    </tr>
  </tbody>
</table>
<br />
<%End If%>
<%If InStr(Admin.GroupId,",5,")>=1 Then%>
<br />
<%End If%>
<%If InStr(Admin.GroupId,",2,")>=1 Then%>
<br />
<%End If%>
<br />
</body>
</html>
<%
Set Rs = Nothing
Call CloseConn()
Set Admin = Nothing
%>
