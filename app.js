/**
 * 
 * @authors Your Name (you@example.org)
 * @date    2017-08-11 11:21:03
 * @version $Id$
 */
var express = require('express');
var app = express();
let router = express.Router(); 
var request = require('request');
var cheerio = require('cheerio');
var Excel = require('exceljs');//json解析成excel

var header = {
	    'User-Agent': 'Mozilla/5.0 (Windows NT 6.3; WOW64; rv:39.0) Gecko/20100101 Firefox/39.0',
	    'Cookie': 'JSESSIONID=758F0FB919E1162FF69B11F3B7FF19B0; auth_key=89',//使用Cookie发送登录信息
	    'Connection': 'keep-alive'
  	}


//获取用户信息
function userinfo (url,callback){	
	request({
		url:url,
		headers:header
	}, function(err, res,body) {
		if (!err && res.statusCode == 200) {
	        //var $ = cheerio.load(body); //采用cheerio模块解析html
	        //var items =$('.table td[class=username]').html();//根据html选择器，获得链接所在的html元素
	           
	     	//    items.each(function(index,element){
		    //    console.log($(element).text());
		    // })	
		    var userlist = [];		
		    var rows = JSON.parse(body).rows;
		    rows.forEach(function(item,index){
		    	if (item.deptgroup == '0') {
		    		item.deptgroup = '1';
		    	}
		    	userlist.push({		            
		            username: item.username,
		            email: item.email,
		            password: item.password,
		            phone:item.phone,
		            mobile:item.mobile,
		            userphoto: item.userphoto,
		            gender:item.sex,
		            create_time: item.create_time,	
		            state:item.state,
		            note:item.note,	            
		            roleid: item.roleid,
		            deptid: item.deptid,
		            deptgroupid:item.deptgroup		           
	            });		    
	            
		    })		   	
	    	callback(userlist);//将用户信息返回给router
	    }

	})
}

//获取菜单列表
function getmenu (url,callback){
	//http://crm.huajin100.com/menu/queryAllMenu.do
	request({
		url:url,
		headers:header
	},function(err,res,body){
		if (!err && res.statusCode == 200) {

			var mainmenu = [];	
			var secondrenMenu = [];			
			var rows = JSON.parse(body).rows;
			rows.forEach(function(item,index){
				mainmenu.push({
					mainmenucode:item.menucode,
					//mainmenuid:item.menuno,
					mainselected:item.selected,
					mainmenuname:item.menuname,
					menuimgurl:item.menuimgurl
				})				
				item.children.forEach(function(item,index){
					secondrenMenu.push({
						secondmenucode:item.menucode,
						secondmenuimgurl:item.menuimgurl,
						secondmenuname:item.menuname,
						secondmenuurl:item.menuurl,
						secondselected: item.selected						
					})
				})
			})				
			callback(mainmenu,secondrenMenu,rows)
		}
	})

}
//将json转化成excel表格,方便之后导入数据库
//在navicat中导入excel表格时，会出现无法打开的情况，这时候发现把excel文件打开再导就好了

function toexcel (xlsxname,option,data) { 
	var start_time = new Date();
	var workbook = new Excel.stream.xlsx.WorkbookWriter({
	  filename: './'+xlsxname+'.xlsx'
	});
	var worksheet = workbook.addWorksheet('Sheet');
	worksheet.columns = option;//定义excel表头字段
	// var data = [{
	//   id: 100,
	//   name: 'abc',
	//   phone: '123456789'
	// }];
	var length = data.length;

	// 当前进度
	var current_num = 0;
	var time_monit = 400;
	var temp_time = Date.now();

	console.log('开始添加数据');
	// 开始添加数据
	for(let i in data) {
	  worksheet.addRow(data[i]).commit();
	  current_num = i;
	  if(Date.now() - temp_time > time_monit) {
	    temp_time = Date.now();
	    console.log((current_num / length * 100).toFixed(2) + '%');
	  }
	}
	console.log('添加数据完毕：', (Date.now() - start_time));
	workbook.commit();

	var end_time = new Date();
	var duration = end_time - start_time;

	console.log('用时：' + duration);
	console.log("程序执行完毕");

}

router.get('/', function(req, res, next) {
	//http://crm.huajin100.com/system/userMang.do	
	userinfo('http://crm.huajin100.com/system/userlist.do?currentPage=1&pageSize=10000',function(info){		
		res.json({result:info});
		var option = [//定义excel表头字段
		    { header: 'username', key: 'username' },
		    { header: 'email', key: 'email' },
		    { header: 'password', key: 'password' },
		    { header: 'phone', key: 'phone' },
		    { header: 'mobile', key: 'mobile' },
		    { header: 'userphoto', key: 'userphoto' },
		    { header: 'gender', key: 'gender' },
		    { header: 'create_time', key: 'create_time' },
		    { header: 'state', key: 'state' },
		    { header: 'note', key: 'note' },
		    { header: 'roleid', key: 'roleid' },
		    { header: 'deptid', key: 'deptid' },
		    { header: 'deptgroupid', key: 'deptgroupid' }
		];		
		toexcel ('userinfo',option,info) 
	})	
}).get('/getmenu',function(req, res, next){
		
	getmenu('http://crm.huajin100.com/menu/queryAllMenu.do',function(mainmenu,secondrenMenu,kk){		
		//res.json(mainmenu);
		res.json(kk);
		var option = [//定义excel表头字段
			{ header: 'mainmenucode', key: 'mainmenucode' },
			{ header: 'mainselected', key: 'mainselected' },
			{ header: 'mainmenuname', key: 'mainmenuname' },
			{ header: 'menuimgurl', key: 'menuimgurl' }		 
		];
		var secondoption =  [//定义excel表头字段
			{ header: 'secondmenucode', key: 'secondmenucode' },
			{ header: 'secondmenuimgurl', key: 'secondmenuimgurl' },
			{ header: 'secondmenuname', key: 'secondmenuname' },
			{ header: 'secondmenuurl', key: 'secondmenuurl' },
			{ header: 'secondselected', key: 'secondselected' }		 
		];
		//toexcel ('mainmenu',option,mainmenu); 
		//toexcel ('secondmenu',secondoption,secondrenMenu);
	})	
})

app.use(router);

app.listen(3000, function() {
  console.log('listening at 3000');
});
