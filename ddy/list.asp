<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%
Const dbdns = "./"
Const Htmldns = "./"
Const Base_Target = "target=""_blank"""
%>
<!--#include file="AppCode/Conn.asp"-->
<!--#include file="AppCode/fun/function.asp"-->
<!--#include file="AppCode/Pager.asp"-->
<!--#include file="vbs.asp" -->
<%
Set Page = New TurnPage
ClassID=myCdbl(Request.QueryString("id"))
Sql="select * from Ok3w_Class where ID=" & ClassID
Rs.Open Sql,Conn,0,1
SortPath = Rs("SortPath")
SortName = Rs("SortName")
Rs.Close
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title><%Call Ok3w_Article_Class_PageTitle(SortPath)%>_文章分类_<%=Application("Ok3w_SiteName")%></title>
<meta name="keywords" content="<%=Application("Ok3w_SiteKeyWords")%>">
<meta name="description" content="<%=Application("Ok3w_SiteDescription")%>">

<link rel="stylesheet" href="css/reset.css" />
<link rel="stylesheet" href="css/text.css" />
<link rel="stylesheet" href="css/960_24_col.css" />
<link rel="stylesheet" href="css/index.css" />
<link href="css/menu.css" rel="stylesheet" type="text/css" />
<script type="text/javascript" src="js/jquery-1.4.2.min.js"></script>
<script type="text/javascript" src="js/menu.js" ></script>
</head>

<body>
<div class="container_24 all">
		
		<div class="header">
            <div class="grid_24">
                <img src="img/banner.gif" align="banner" />
            </div>
            <div class="grid_24 nav">                
                    <%Call Ok3w_Article_Class_NavMenus()%>
            </div>
            <div class="clear"> </div>
		</div> 
		<div class="grid_24 content">
		<div class="grid_6 alpha content_left">
        	<!--需要添加的内容，左侧导航条菜单-->
            <div class="menu">
            	<div class="text">
                        <%=SortName%>
                </div>
            </div>
            <div class="menudata3">
            	<!--根据类ID列出该ID下的子类名称-->
                <%Call Ok3w_Article_Class_NavLists(ClassID)%>
            	<!--<img src="img/left_img/download2.png"/>-->
            </div>
            
             
        </div>
        <div class="grid_18 omega content_right">
        	<!--需要添加的内容，右侧内容区-->
            <!--右侧内容区头部位置信息-->
            <div class="content_right_header"> 	
  			  当前位置：    <a href="index.asp">首页</a> &gt;&gt; 
            <%Call Ok3w_Article_Class_Nav(SortPath)%>
            <span id="ctl00_ctl00_PlhMiddle_SiteMapCtrl1_RptSiteMap_ctl01_LblSeperator">&gt;&gt;</span>          
            <span class="red_text">文章列表</span>
            </div>
            
            <!--主内容区-->
            <div class="content_right_body">
                    <%Call Ok3w_Article_bList(ClassID,20)%>
            </div>
        </div>
		<div style="clear:both"></div>
		</div>
        <div class="grid_24 footer">
           <div class="right">
    版权所有：中国地质大学（武汉）信息工程学院软件工程系
    		</div>
    	</div>
</div>

</body>
</html>
