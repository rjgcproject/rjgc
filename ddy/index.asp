<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%
Const dbdns = "./"
Const Htmldns = "./"
Const Base_Target = "target=""_blank"""
%>

<!--#include file="AppCode/Conn.asp"-->
<!--#include file="AppCode/fun/function.asp"-->
<!--#include file="vbs.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<!--<title>欢迎光临中国地质大学（武汉）研究生就业信息网</title>-->
<title><%=Application("Ok3w_SiteTitle")%></title>
<meta name="keywords" content="<%=Application("Ok3w_SiteKeyWords")%>">
<meta name="description" content="<%=Application("Ok3w_SiteDescription")%>">

<link rel="stylesheet" href="css/reset.css" />
<link rel="stylesheet" href="css/text.css" />
<link rel="stylesheet" href="css/960_24_col.css" />
<link rel="stylesheet" href="css/index.css" />

<script language="javascript" src="js/js.js"></script>
<SCRIPT src="js/sohuflash_1.js" type=text/javascript></SCRIPT>
<link href="images/fish.css" rel="stylesheet" type="text/css" />

<style type="text/css">
<!--

-->
</style>

<script type="text/javascript" src="js/jquery-1.4.2.min.js"></script>
<script type="text/javascript" src="js/embed_index.js" ></script>
<script type="text/javascript">
	$().ready(function(){
		onloadlst("index","info-list1",5,"","全职","");//版位标示，列表层ID，条数，职位类别，职位性质，单位性质
		onloadlst("index","info-list2",5,"","实习","");
	});
</script>
<link href="css/menu.css" rel="stylesheet" type="text/css" />
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
        <div class="grid_8 content">
        	<div class="left_top">
            	<!--<img src="img/left_top.gif" />-->
                <%Call Ok3w_Article_ImgFlash("",290,250)%>
            </div>
            <div class="left_middle">
            	<div class="menu">
                	<div class="text">
                    	就业信息
                    </div>
                    <div class="more">
                    <a href="list.asp?id=32">More</a>
                    </div>
                </div>
                <div class="menudata width1">
                	<%Call Ok3w_Article_List(32,7,1,20,False,False,True,"5",False,"")%>
                </div>
                
            </div>
            <div class="left_bottom">
            	<div class="menu">
                	<div class="text">下载专区</div>
                    <div class="more">
                    <a href="list.asp?id=114">More</a>
                    </div>
                </div>
                <div class="menudata width1">
                	<%Call Ok3w_Article_List(114,7,1,20,False,False,True,"5",False,"")%>
                </div>
            </div>
        </div>
        <div class="grid_9 content">
        	<div class="center_top">
            	<div class="menu">
                	<div class="text">
                    	最新通知
                    </div>
                    <div class="more">
                    <a href="list.asp?id=110">More</a>
                    </div>
                </div>
                <div class="menudata">
                	<%Call Ok3w_Article_List(110,7,1,30,False,False,True,"5",False,"")%>
                </div>
            </div>
            <div class="center_middle">
            	<div class="menu">
                	<div class="text">
                    	学术信息
                    </div>
                    <div class="more">
                    <a href="list.asp?id=126">More</a>
                    </div>
                </div>
                <div class="menudata">
                	<%Call Ok3w_Article_List(126,7,1,20,False,False,True,"5",False,"")%>
                </div>
            </div>
            <div class="center_bottom">
            	<div class="menu">
                	<div class="text">
                    	学科建设
                    </div>
                    <div class="more">
                    <a href="list.asp?id=6">More</a>
                    </div>
                </div>
                <div class="menudata">
                	<%Call Ok3w_Article_List(6,7,1,20,False,False,True,"5",False,"")%>
                </div>
            </div>
        </div>
        <div class="grid_7 content">
        	<div class="right_top">
            	<div class="menu">
                	<div class="text">
                    	教学通告
                    </div>
                    <div class="more">
                    <a href="list.asp?id=4">More</a>
                    </div>
                  
                </div>
                <div class="menudata2">
                	<p>
                    <span class="p">【今天】</span>
           			<marquee direction="up" height="207" scrollamount="2" scrolldelay="150" class="ggtitle1a" onmouseover="this.stop()" onmouseout="this.start()">
                    <%Call Ok3w_Article_List(51,4,1,40,False,False,False,"",False,"new")%>
                    </marquee><br>
            		<span class="q">【明天】</span> 
            		<marquee direction="up" height="207" scrollamount="2" scrolldelay="150" class="ggtitle1a" onmouseover="this.stop()"  onmouseout="this.start()">
                    <%Call Ok3w_Article_List(99,4,1,40,False,False,False,"",False,"new")%></marquee><br>
          			</p>
                </div>
            </div>
            <div class="right_bottom">
            	<div class="menu">
                	<div class="text">
                    	友情链接
                    </div>
                    <div class="more">
                    <img src="img/link.png" />
                    </div>
                </div>
                <div class="menudata">
					<select onchange=javascript:window.open(this.options[this.selectedIndex].value) name=D2>
                    <option value="#"> ==企业合作== </option>
                    <option value="http://www.mapgis.com.cn/">中地数码</option>
                    </select>
                    
                    <select onchange=javascript:window.open(this.options[this.selectedIndex].value) name=D1>
                    <option value="#">   ==校内网站== </option>
                    <option value="http://www.cug.edu.cn">中国地质大学（武汉）</option>
                    <option value="http://graduate.cug.edu.cn/">地大研究生院</option>
                    <option value="http://xgxy.cug.edu.cn/ygb/">信息工程学院</option>                    
                    </select>
                    
                    <select onchange=javascript:window.open(this.options[this.selectedIndex].value) name=D1>
                    <option value="#">   ==兄弟院校== </option>
                    <option value="http://www.cugb.edu.cn/">中国地质大学（北京）</option>
                    <option value="http://sse.hust.edu.cn/">华中科技大学软件学院</option>
                    <option value="http://iss.whu.edu.cn/">武汉大学国际软件学院</option>                               
                    </select>
                </div>
            </div>
        </div>
        <div class="grid_24 footer">
           <div class="right">
				版权所有：中国地质大学（武汉）信息工程学院软件工程系
    		</div>
    	</div>
</div>
</body>
</html>
