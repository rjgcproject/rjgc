<%
'基本信息
countname	= "研究生就业办公室——研究生工作部统计系统"	' 统计器名称
yesleft		= 1				' 是否显示左侧的导航，1表示是，0表示否
yestop		= 0				' 横导航栏放置位置，0表示不放，1表示放页顶，2表示放页底
connpath	= "count.mdb"	' 数据库存放位置，这里填写相对index.asp的相对路径
whatcan		= 4				' 访客的权限等级，有关等级划分请看 readme 文档
CookieExpires=100			' cookie设置时间(天),默认为100天

'外观参数
FlashWidth	= 130			' FLASH图标宽度
FlashHeight	= 58			' FLASH图标高度
mPageSize	= 15			' 每页记录数（仅对需分页的页面）
mPrecision	= 5				' 精确到小数点后多少位

'站点信息
mURL		= "http://unit.cug.edu.cn/yjsy/yjsjy"	' 站点连接
mName		= "研究生就业办公室"				' 站点名称
mNameEn		= "yjsjy"			' 站点英文名称
masterName	= "admin"					' 管理员名
masterpsw	= "admin"					' 管理员密码
masterEmail	= "cework@126.com"			' 管理员电子邮件
SiteBrief	= "中国地质大学研究生就业办——研究生工作部。"	' 站点介绍，请勿使用英文的双引号和换行

%>