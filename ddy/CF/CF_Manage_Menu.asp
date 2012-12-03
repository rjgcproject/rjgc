<%
'乘风多用户计数器
'作者QQ：178575
'作者EMail：yliangcf@163.com
'作者网站：http://www.qqcf.com
'详细简介：http://www.qqcf.com/?action=list&list=cfcount
'上面有程序在线演示，安装演示，使用疑难解答，最新版本下载等内容
'因为这些内容可能时常更新，就没有放在程序里，请自己上网站上查看
'有完整版本的演示
%>
<table class="tb_3 td_1 tb_menu" style="margin:5px 0px 0px 0px;width:100%">
  <tr class="tr_1 td_1"> 
    <td >
	用户：<%=UserName%><br> <%
Set Rs= Server.CreateObject("Adodb.RecordSet")
Sql="Select * From CFCount_User Where UserName='"&UserName&"'"
Rs.Open Sql,Conn,1,1
%> 
<%If Rs("ParentName")<>"" And Session("CFCountSubAdmin")<>"" Then
 Response.write "<input type='submit' name='Submit' value='返回父账号' onclick=""location.href='?Action=parentgoto'"">"
End If
If Session("CFCountUser_View")<>"" Then
 Response.write "独立查看者"
End If
%>
</td>
  </tr>
  <tr> 
    <td ><img src=images/a.gif> <b><a href=?Action=tjsurvey>统计概况</a></b></td>
  </tr>
  <tr> 
    <td ><img src=images/closef.gif> <b><a href=?Action=lastvisit>最新访问</a></b></td>
  </tr>
  <tr> 
    <td ><img src=images/closef.gif> <b><a href=?Action=lylist>来源明细</a></b></td>
  </tr>
  <tr> 
    <td ><img src=images/closef.gif> <b><a href=?Action=onlinetj>在线统计</a></b></td>
  </tr>
  <tr> 
    <td ><img src=images/closef.gif> <b><a href=?Action=webtj>页面统计</a></b></td>
  </tr>
  <tr> 
    <td height=24 width=100%><img src=images/openf.gif id=aimgboard1  align="absmiddle"> 
      <b><a href="#">统计报表</a></b></td>
  </tr>
  <tr> 
    <td style="display:yes;" id='submenuboard1'><table cellpadding=0 cellspacing=0 align=left >
        <tr> 
          <td ></td>
          <td ><img src=images/open.gif border=0 align="absmiddle"> 
            <a href="?Action=hourtj">按小时统计</a></td>
        </tr>
        <tr> 
          <td ></td>
          <td ><img src=images/open.gif border=0 align="absmiddle"> 
            <a href="?Action=daytj">按天数统计</a></td>
        </tr>
        <tr> 
          <td ></td>
          <td ><img src=images/open.gif border=0 align="absmiddle"> 
            <a href="?Action=monthtj">按月份统计</a></td>
        </tr>

        <tr> 
          <td ></td>
          <td ><img src=images/open.gif border=0 align="absmiddle"> 
            <a href="?Action=backtj">回头率统计</a></td>
        </tr>
        <tr> 
          <td ></td>
          <td ><img src=images/open.gif border=0 align="absmiddle"> 
            <a href="?Action=alexatj">Alext条统计</a></td>
        </tr>
      </table></td>
  </tr>
  <tr> 
    <td height=24 width=100%><img src=images/openf.gif id=aimgboard1  align="absmiddle"> 
      <b><a href="#">搜索引擎统计</a></b></td>
  </tr>
  <tr> 
    <td style="display:yes;" id='submenuboard1'> <table cellpadding=0 cellspacing=0 align=left >
        <tr> 
          <td ></td>
          <td ><img src=images/open.gif border=0 align="absmiddle"> 
            <a href="?Action=searchtj">数量统计</a></td>
        </tr>
        <tr> 
          <td ></td>
          <td ><img src=images/open.gif border=0 align="absmiddle"> 
            <a href="?Action=keywordtj">关键字统计</a></td>
        </tr>
      </table></td>
  </tr>
<%
If Session("CFCountUser_View")="" Then'不是独立查看者时
%>
<tr> 
    <td height=24 width=100%><img src=images/openf.gif id=aimgboard1  align="absmiddle"> 
      <b><a href="#">计数器设置</a></b></td>
  </tr>
  <tr> 
    <td style="display:yes;" id='submenuboard1'> <table cellpadding=0 cellspacing=0 align=left >
        <tr> 
          <td ></td>
          <td ><img src=images/open.gif border=0 align="absmiddle"> 
            <a href="?Action=jsqset">参数设置</a></td>
        </tr>
        <tr> 
          <td ></td>
          <td ><img src=images/open.gif border=0 align="absmiddle"> 
            <a href="?Action=jsqquickset">参数快速设置</a></td>
        </tr>
        <tr> 
          <td ></td>
          <td ><img src=images/open.gif border=0 align="absmiddle"> 
            <a href="?Action=stylelist">图片样式选择</a></td>
        </tr>
        <tr> 
          <td ></td>
          <td ><img src=images/open.gif border=0 align="absmiddle"> 
            <a href="?Action=imglist">网店类样式选择</a></td>
        </tr>

        <tr> 
          <td ></td>
          <td ><img src=images/open.gif border=0 align="absmiddle"> 
            <a href="?Action=jsqreset">计数器重置</a></td>
        </tr>
        <tr> 
          <td ></td>
          <td ><img src=images/open.gif border=0 align="absmiddle"> 
            <a href="?Action=userdel">账号删除</a></td>
        </tr>

      </table></td>
  </tr>
  <tr> 
    <td height=24 width=100%><img src=images/openf.gif id=aimgboard1  align="absmiddle"> 
      <b><a href="#">修改个人资料</a></b></td>
  </tr>
  <tr> 
    <td style="display:yes;" id='submenuboard1'> <table cellpadding=0 cellspacing=0 align=left >
        <tr> 
          <td ></td>
          <td ><img src=images/open.gif border=0 align="absmiddle"> 
            <a href=?Action=pwdmodify>修改密码</a></td>
        </tr>
        <tr> 
          <td ></td>
          <td ><img src=images/open.gif border=0 align="absmiddle"> 
            <a href=?Action=datamodify>修改资料</a></td>
        </tr>
        <tr> 
          <td ></td>
          <td ><img src=images/open.gif border=0 align="absmiddle"> 
            <a href=?Action=passwordanswermodify>修改密码保护</a></td>
        </tr>
      </table></td>
  </tr>

  <tr> 
    <td ><img src=images/openf.gif> <b><a href=?Action=gbooklist>留言列表</a></b></td>
  </tr>
<%End If%>

<tr> 
    <td ><img src=images/openf.gif> <b><a href=?Action=getcode>获得统计代码</a></b></td>
  </tr>
  <tr> 
    <td ><img src=images/exit.gif> <b><a href=?Action=userlogout>退出管理</a></b></td>
  </tr>
</table>
