<%
'乘风多用户计数器 Sql版
'作者QQ：178575
'作者EMail：yliangcf@163.com
'作者网站：http://www.qqcf.com
'详细简介：http://www.qqcf.com/?action=list&list=cfcount
'上面有程序在线演示，安装演示，使用疑难解答，最新版本下载等内容
'因为这些内容可能时常更新，就没有放在程序里，请自己上网站上查看
'有完整版本的演示
%>

<style type="text/css">
/*全局设置开始*/

* {padding:0; margin:0}

body {text-align: left; font-family:Arial; margin:0; padding:0; background: #FFFFFF; font-size:12px; color:#333333;}

div,form,img,ul,li{margin: 0; padding: 0; border: 0;}

div,table{overflow:hidden;}

h1,h2,h3,h4,h5,h6 { margin:0; padding:0;}
ul,ol,li{list-style-type:none}

a:link,a:visited,a:active {color: #02418a; text-decoration: none; }

a:hover {color: #ff0000; text-decoration: underline}




#top{
 width:980px;
 margin:3px auto 0px;
 background-color:#<%=GetSkinColor(RsSet("SkinColor"),1)%>;
 border-right:1px solid #<%=GetSkinColor(RsSet("SkinColor"),0)%>;
 border-top:1px solid #<%=GetSkinColor(RsSet("SkinColor"),0)%>;
 border-left:1px solid #<%=GetSkinColor(RsSet("SkinColor"),0)%>;
 border-bottom:1px solid #<%=GetSkinColor(RsSet("SkinColor"),0)%>;
}

#top_logo {
 width:230px;
 float:left;
}

#top_mainmenu { 
 background:url(images/top_bg0.gif) repeat-x;
 border-top:1px solid #a7cde5;
 border-bottom:1px solid #a7cde5;
 border-left:1px solid #a7cde5;
 border-right:1px solid #a7cde5;
 font-size:14px; 
 float:left;
 height:30px;
 line-height:30px;
 margin-left:70px;
 margin-top:42px;
 width:650px; 
}


#u_mainmenu li {
 padding-right:9px;
 padding-left:10px;
 font-weight:bold; 
 float:left;
 width:60px
}

#bottom{
 text-align:center;
 line-height:160%;
 width:980px;
 margin:5px auto;
 background-color:#<%=GetSkinColor(RsSet("SkinColor"),1)%>;
 border-right:1px solid #<%=GetSkinColor(RsSet("SkinColor"),0)%>;
 border-top:1px solid #<%=GetSkinColor(RsSet("SkinColor"),0)%>;
 border-left:1px solid #<%=GetSkinColor(RsSet("SkinColor"),0)%>;
 border-bottom:1px solid #<%=GetSkinColor(RsSet("SkinColor"),0)%>;
 padding:10px 0px;
}


table {
 font-weight:normal;
 font-size:12px;
 line-height:180%;
 font-family:arial;
}
input {
 height:20px;
 border: 1px solid #<%=GetSkinColor(RsSet("SkinColor"),0)%>;
 BACKGROUND-color:#FFFFFF;

}

button{
 BORDER: #c4dae3 1px solid;
 FONT-SIZE: 12px; 
 CURSOR: hand; 
 HEIGHT: 24px;
 BACKGROUND-color:#FFFFFF;
 font-weight:bold;
 margin-top:5px;
 margin-bottom:5px; 
}




select {
 height:20px;
 border: 1px solid #<%=GetSkinColor(RsSet("SkinColor"),0)%>;
}

textarea {
 border: 1px solid #<%=GetSkinColor(RsSet("SkinColor"),0)%>;
}
/*全局设置结束*/

.button_2{
 BACKGROUND-color:#<%=GetSkinColor(RsSet("SkinColor"),1)%>;
}

.button_3{
 width:140px;
}

.button_4{
 width:140px;
 BACKGROUND-color:#<%=GetSkinColor(RsSet("SkinColor"),1)%>;
}

.co{
 width:980px;
 background-color:#<%=GetSkinColor(RsSet("SkinColor"),1)%>;
 margin-top:5px;
 border:1px solid #<%=GetSkinColor(RsSet("SkinColor"),0)%>;
 margin:3px auto 0px;
}

.co_1{
 width:980px;
 margin:3px auto 0px;
}

.co_left{
 width:202px;
 float:left;
}
.co_right{
 width:778px;
 float:left;
}

.co_d_1{
 width:200px;
 background-color:#<%=GetSkinColor(RsSet("SkinColor"),1)%>;
 margin-top:5px;
 border:1px solid #<%=GetSkinColor(RsSet("SkinColor"),0)%>;
}

.co_d_2 {
 width:771px; /*771+left5px+boder1px*2=778*/
 background-color:#<%=GetSkinColor(RsSet("SkinColor"),1)%>;
 margin:5px 0px 0px 5px;
 border:1px solid #<%=GetSkinColor(RsSet("SkinColor"),0)%>;
 float:right;
}

.l_con_left{
 padding:5px;
 margin-left:0px;
 line-height:160%;
 text-align:left;
}


.l{
 padding:5px;
 margin-left:0px;
 line-height:160%;
 text-align:left;
 list-style-type:none; 
}


.d_1 {
 width:980px;
 background-color:#<%=GetSkinColor(RsSet("SkinColor"),1)%>;
 margin:5px auto 0px;
 border:1px solid #<%=GetSkinColor(RsSet("SkinColor"),0)%>;
}

.d_1_1 {
 padding-left:5px;
 padding-top:5px;
 font-weight:bold;
 font-size:14px;
 height:24px;
}

.d_1_2 {
 width:100%;
 background-color:#ffffff;
 float:left;
 padding:5px 10px;
}

.d_1_3 {
 padding-left:5px;
 padding-top:5px;
 font-weight:bold;
 font-size:14px;
 height:24px;
 text-align:left;
}




ul{
 list-style-type:none;
 margin:0px;
}

li{
 line-height:120%;
 text-align:left;
 list-style-type:none;
 padding-top:5px;
}

.l_1{
 text-align:center;
}

.l_2{
 width:90px;
 float:left;
 padding-top:5px;
 margin:0px;
 text-align:center;
}


.l_3{
 padding-top:5px;
 padding-left:20px;
 text-align:left;
 font-weight:bold;
}

.l_4{
 padding-top:5px;
 padding-left:5px;
 text-align:left;
 font-weight:normal;
}

.l_4_off{
 padding-top:5px;
 padding-left:5px;
 text-align:left;
 font-weight:normal;
 display:none;
}


.l_userlogin_1{
 width:60px;
 float:left;
 padding-left:0px;
 padding-top:12px;
 text-align:right;
 clear:both;
}

.l_userlogin_2{
 width:140px;
 float:left;
 padding-top:5px;
}



.d_menu { 
 background:url(images/top_bg0.gif) repeat-x;
 border-top:1px solid #a7cde5;
 border-bottom:1px solid #a7cde5;
 border-left:1px solid #a7cde5;
 border-right:1px solid #a7cde5;
 font-size:14px; 
 float:left;
 height:30px;
 line-height:30px;
 margin-left:70px;
 margin-top:38px;
 width:650px; 
}

.l_menu {
 padding-right:9px;
 padding-left:10px;
 font-weight:bold;
 float:left;
 padding-bottom:0px;
 padding-top:8px;
}



.tb_1{
 background-color:#ffffff;
 border:1px solid #<%=GetSkinColor(RsSet("SkinColor"),0)%>;
 margin:5px 0px 5px 5px;
 width:768px;
 float:left;
 border-collapse:collapse;
}

.tb_1 td{border-bottom: 1px dotted #<%=GetSkinColor(RsSet("SkinColor"),0)%>;padding-left:6px;empty-cells:show;}

 .tb_2{
 width:980px;
 background-color:#ffffff;
 border:1px solid #<%=GetSkinColor(RsSet("SkinColor"),0)%>;
 margin:15px auto;
 clear:both;
 float:center;
}

 .tb_3{
 background-color:#ffffff;
 border:1px solid #<%=GetSkinColor(RsSet("SkinColor"),0)%>;
 margin:15px auto;
 padding:5px;
 clear:both;
}


.tr_1{
 padding-left:5px;
 padding-top:5px;
 font-weight:bold;
 font-size:14px;
 height:24px;
 text-align:center;
 background-color:#<%=GetSkinColor(RsSet("SkinColor"),1)%>;
}

.tr_2{
 text-align:center;
}



.td_1{
 text-align:left;
}
.td_2{
 text-align:center;
}
.td_3{
 text-align:right;
}

.t_button {
	BORDER-RIGHT: #999999 1px solid; PADDING-RIGHT: 4px; BACKGROUND-POSITION: 1px 1px; BORDER-TOP: #cccccc 1px solid; DISPLAY: inline-block; PADDING-LEFT: 4px; FONT-WEIGHT: bold; PADDING-BOTTOM: 2px; MARGIN: 0px 6px 0px 0px; BORDER-LEFT: #cccccc 1px solid; COLOR: #333333; LINE-HEIGHT: normal; PADDING-TOP: 2px; BORDER-BOTTOM: #999999 1px solid; BACKGROUND-REPEAT: no-repeat; FONT-FAMILY: Tahoma, Arial, Helvetica; WHITE-SPACE: nowrap; BACKGROUND-COLOR: #eeeeee;
}

.td_wauto{width:34px;word-break:keep-all;}
.td_nowrap{white-space:nowrap;}
.tb_underline td{border-bottom: 1px dotted #<%=GetSkinColor(RsSet("SkinColor"),0)%>;padding-left:6px;}
.tb_menu td{padding-left:10px;height:22px}
.tb_color td{background-color:#<%=GetSkinColor(RsSet("SkinColor"),1)%>}
</style>