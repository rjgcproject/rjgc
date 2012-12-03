  provinceItems = [
    {id:'11',name:'北京'},
    {id:'12',name:'天津'},
    {id:'13',name:'河北'},
    {id:'14',name:'山西'},
    {id:'15',name:'内蒙古'},
    {id:'21',name:'辽宁'},
    {id:'22',name:'吉林'},
    {id:'23',name:'黑龙江'},
    {id:'31',name:'上海'},
    {id:'32',name:'江苏'},
    {id:'33',name:'浙江'},
    {id:'34',name:'安徽'},
    {id:'35',name:'福建'},
    {id:'36',name:'江西'},
    {id:'37',name:'山东'},
    {id:'41',name:'河南'},
    {id:'42',name:'湖北'},
    {id:'43',name:'湖南'},
    {id:'44',name:'广东'},
    {id:'45',name:'广西'},
    {id:'46',name:'海南'},
    {id:'50',name:'重庆'},
    {id:'51',name:'四川'},
    {id:'52',name:'贵州'},
    {id:'53',name:'云南'},
    {id:'54',name:'西藏'},
    {id:'61',name:'陕西'},
    {id:'62',name:'甘肃'},
    {id:'63',name:'青海'},
    {id:'64',name:'宁夏'},
    {id:'65',name:'新疆'},
	{id:'71',name:'台湾'},
	{id:'81',name:'香港'},
	{id:'82',name:'澳门'},
	{id:'00',name:'其他'}
    ];
  
	//解析数据集合，有置顶，无分页，每单位只有一条职位
	function onloadlst(board,viewdivID,psize,category,jobType,natureCode)
	{
		var siteUrl='http://lnjy.ncss.org.cn';//一站式服务站点网址，根据本站在admin.ncss.org.cn后台“服务站点设置”中修改
		var jUrl=siteUrl+'/json/general_searchp?callback=?';
		var params = getParams(psize,category,jobType,natureCode);
		$.getJSON(jUrl,params,
				function(data)
				{	var lists = '';
					var divElement="#"+viewdivID;
					var date ;
					var areaCode;
					var area="";
					var topFlag="";

					if(board=="index"){//首页
						$.each(data.lst, function()
						{							
							date = this.postDate.substring(5,10);
							areaCode = this.areaCode.substring(0,2);							
							$.each(provinceItems, function(){
								if(this.id==areaCode){area=this.name;}			 
							});
							if(this.priority==1){
								lists += '<tr><td class="hui21">['
								+ area
								+ ']&nbsp;<a class="hui21" href="' + siteUrl + '/job/view_job?jobId='
								+ this.jobId
								+ '" target="_blank" title="职位名称：'
								+ this.jobTitle
								+ '&#10单位行业：' +this.industry+'&#10单位性质：' +this.nature 
								+'">'
								+ this.recName
								+ '</a>('
								+ date
								+')</td></tr>';
							}else{
								lists += '<tr><td class="hui21">['
								+ area
								+ ']&nbsp;<a class="hui21" href="' + siteUrl + '/job/view_job?jobId='
								+ this.jobId
								+ '" target="_blank" title="职位名称：'
								+ this.jobTitle
								+ '&#10单位行业：' +this.industry+'&#10单位性质：' +this.nature 
								+'">'
								+ this.recName
								+ '</a>('
								+ date
								+')</td></tr>';
							}
							
							
						});	
						
					}					
					
					
					if(lists!='')
						{
							
							$(divElement).html(lists);
						}else
						{
							$(divElement).html("页面加载中...");
						}
				}
				
		);
	}
	    	
	//初始化查询参数
	function getParams(psize,category,jobType,natureCode)
	{
	   var params = {
			recName:"",
			natureCode:natureCode,
			industryCode:"",
			recScale:"",
			jobTitle:"",
			category:category,
			jobType:jobType,
			areaCode:"",
			degreeCode:"",
		    dayLimit:"-1",
			siteId:"",//默认空为全部，00为中心，10001为北京大学
		    psize: parseInt(psize),//每页条数
		    //pindex: parseInt(pindex),//显示第？页
			callback:"test"};
		   return params;
	}