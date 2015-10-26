var show_word="输入商品关键字";
function searchSubmitTo(){
	var obj=document.getElementById("KeyWord");
	var formobj=document.getElementById("searchForm");
	var keyType = document.getElementById("keywordtype");
	var classid = document.getElementById("sClassid");
	var shopclassid = document.getElementById("shoptype");
	formobj.action="http://search1.paipai.com/cgi-bin/comm_search1";
	if(!keyType||keyType.value == "goods"||keyType.type=="radio"){
		if(obj.value!=""){
			formobj.submit();
		} else if(document.getElementById("sClassid")){
			if(document.getElementById("sClassid").value!="")
				formobj.submit();
			else location.href="http://search.paipai.com/";
		} else if(document.getElementById("Path")){
			if(document.getElementById("Path").value!="")
				formobj.submit();
			else location.href="http://search.paipai.com/";
		}
	}else{
		if(obj.value!=""){
			formobj.action="http://search1.paipai.com/cgi-bin/shop_search1";
			shopclassid.value = classid.value;
			formobj.submit();
		} else location.href="http://shop.paipai.com/";
	}
	return false;
}

function makeSearch(){
	document.getElementById("keywordtype").value=document.getElementById("searchType_foot").value;
	document.getElementById("KeyWord").value=document.getElementById("KeyWord_foot").value;
	searchSubmitTo();
}

function showSearch(obj,type){
	if(type){
		if(obj.value=="")
			obj.value="";
		obj.style.color='gray';
	}else{
		if(obj.value==show_word||obj.value=="输入关键字进行搜索")
			obj.value="";
		obj.style.color='#000000';
	}
}

function showFind(sortId){
	var obj=document.getElementById("sClassid");
	var length;
	length = obj.length;
	for(var index=0;index<length;index++){
		if(obj.options[index].value == sortId){
			obj.selectedIndex = index;
			break;
		}
	} 
}

function showType(sortId){
	var obj=document.getElementById("keywordtype");
	var length;
	length = obj.length;
	for(var index=0;index<length;index++){
		if(obj.options[index].value == sortId){
			obj.selectedIndex = index;
			break;
		}
	} 
}

rnd.today=new Date(); 
rnd.seed=rnd.today.getTime(); 
function rnd(){ 
	rnd.seed = (rnd.seed*9301+49297) % 233280; 
	return rnd.seed/(233280.0); 
}
function rand(number){ 
	return Math.ceil(rnd()*number)-1; 
}

document.onclick=hideResult;
var nowLink=-1;
var preLink=0;
var keyNumber=0;
function showResult(code,value){
	var obj=document.getElementById('result');
	ShowPaiPaiSearchResult(value);
	nowLink=-1;
	preLink=0;
}
function hideResult(){
	var obj=document.getElementById('result');
	if(obj) obj.style.display='none';
}
function showKeyDown(code){
	keyNumber=document.getElementById("result").childNodes.length;
	if(code==38||code==40){
		preLink=nowLink;
		if(code==38)
			nowLink--;
		else if(code==40)
			nowLink++;
		if(nowLink<0) nowLink=keyNumber-1;
		else if(nowLink>keyNumber-1) nowLink=0;
		var objNowLink=document.getElementById("result").childNodes[nowLink];
		if(objNowLink)
			objNowLink.className="activeLink";
		if(nowLink!=-1){
			var objPreLink=document.getElementById("result").childNodes[preLink];
			if(objPreLink)
				objPreLink.className="";	
			}				
	} else if(code==13){
		if(nowLink!=-1){
			var objNowLink=document.getElementById("result").childNodes[nowLink];
			if(objNowLink){
				keyHref=objNowLink.href;
				keyStart=keyHref.indexOf("KeyWord=");
				keyEnd=keyHref.indexOf("&ADTAG");
				keyValue=keyHref.substring(keyStart+8,keyEnd);
				document.getElementById("KeyWord").value=keyValue;
			}
		}
		if(document.getElementById("searchForm"))
			searchSubmitTo(); 	
	}
}
function loadscript(rURL,rID) {
	var oldscript = document.getElementById(rID);
	var newscript = document.createElement("script");
	newscript.setAttribute("id", rID)
	newscript.setAttribute("src", rURL);
	oldscript.parentNode.replaceChild(newscript, oldscript);
	return true;
}
function ShowPaiPaiSearchResult(KeyWord){
    loadscript("http://search.paipai.com/cgi-bin/isuggest?KeyWord="+KeyWord,"PaipaiSearchJs");
    return true;
}

var arrFirstSort=[[3119,"手机充值中心 - 电话卡"],[12001,"网络游戏虚拟商品"],[24590,"芭芭拉专区"],1,[20501,"女装 - 女士精品"],[27158,"女士内衣 - 袜子 - 泳装"],[20001,"彩妆 - 香水 - 护肤 - 美体"],[2001,"珠宝首饰 - 时尚饰品"],[21001,"女士箱包 - 配件"],[21036,"女鞋"],[28001,"品牌手表 - 流行手表"],[21501,"瑞士军刀 - Zippo - 男士饰品"],[22001,"男装 - 男包 - 服饰配件"],[6001,"运动 - 休闲 - 健身"],[6070,"户外 - 军品 - 旅游 - 机票"],[22501,"婴幼 - 孕妇用品 - 童装"],0,[3002,"手机"],[29551,"手机号码 - 上网卡"],[28038,"笔记本电脑"],[1,"电脑硬件 - 台式电脑"],[5001,"随身视听 - 音响 - 耳机"],[4001,"数码相机 - 摄像机 - 冲印"],[28039,"网络服务 - 电脑软件"],[28009,"数码配件 - 电子元件"],[28046,"办公设备 - 文具 - 耗材"],[11001,"动漫周边 - 游戏周边"],[23001,"卡通产品 - 玩具 - 模型"],1,[28055,"居家 - 日用 - 床品 - 厨房"],[28053,"家电 - 厨房电器 - 保健器械"],[28052,"食品 - 保健品 - 茶叶 - 特产"],[28054,"家具 - 宜家代购"],[28056,"装潢 - 灯具 - 五金 - 卫浴"],[8001,"邮币 - 古董 - 字画 - 收藏"],[24001,"宠物 - 工艺品 - 礼品 - 鲜花"],[23501,"书籍 - 报纸 - 杂志"],[24501,"电影 - 电视 - 音乐 - 曲艺"],[9001,"汽车 - 摩托 - 自行车"],[21591,"成人 - 避孕 - 情趣内衣"],[10001,"住宅 - 商铺 - 办公楼租售"]];

function showFirstSort(showType){
	var showSortStr="";
	var showSortMore="";
	if(showType!="link"){
		if(showType==1) 
			showSortMore="0,";
		showSortStr+='<optgroup>';
		for(var i=0;i<arrFirstSort.length;i++){
			if("object"==typeof(arrFirstSort[i]))
				showSortStr+='<option value="'+showSortMore+arrFirstSort[i][0]+'">'+arrFirstSort[i][1]+'</option>';
			else if(0==arrFirstSort[i])
				showSortStr+='</optgroup><optgroup>';
			else
				showSortStr+='</optgroup><optgroup class="optionShow">';			
		}
		showSortStr+='</optgroup>';
	} else {
		for(var i=0;i<arrFirstSort.length;i++)
			if("object"==typeof(arrFirstSort[i]))
				showSortStr+='<li><a href="http://search.paipai.com/cgi-bin/shop_search?shoptype='+arrFirstSort[i][0]+'">'+arrFirstSort[i][1]+'</a></li>'
	}
	document.write(showSortStr);
}