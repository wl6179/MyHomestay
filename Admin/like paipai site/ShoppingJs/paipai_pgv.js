
function do_ping()
{
	if(typeof(pgvMain) != 'function')
	    return;

    // 1. 获取domain/uri/refer
	var domain="";
	var uri="";
	var search="";
	var ref="";
	try
	{
		domain=top.location.hostname;
		uri=top.location.pathname;
		search=top.location.search;
		ref=top.document.referrer;
	}catch(e)
	{
		domain=window.location.hostname;
		uri=window.location.pathname;
		search=window.location.search;
		ref=window.document.referrer;
	}
	domain=domain.toLowerCase();
	uri=uri.toLowerCase();
	search=search.toLowerCase();
	ref=ref.toLowerCase();
	
	/*关键路径
	 商品发布流程关键路径分析2008.3.1               1)选择类目 2)输入商品信息 3)发布成功
	     my.paipai.com
	 拍拍核心购买流程关键路径1.1.1   				1)搜索 2)商品详情店铺详情 3)下单页面 4)下单成功页面 5)到CFT付款 6)CFT付款成功
	     search.paipai.com search1.paipai.com list.paipai.com auction1.paipai.com shop.paipai.com auction.paipai.com pay.paipai.com
	 拍拍搜索到下单(无需登录)关键路径1.1.2    		1)搜索 2)商品详情 3)下单页面
	     search.paipai.com search1.paipai.com list.paipai.com auction1.paipai.com auction.paipai.com
	 拍拍搜索到下单(需登录)关键路径1.1.3     		1)搜索 2)商品详情 3)登录 4)下单页面
	     search.paipai.com search1.paipai.com list.paipai.com auction1.paipai.com member.paipai.com auction.paipai.com
	 拍拍下单到CFT付款成功(实物物品)关键路径1.1.4	1)下单页面 2)下单成功页面 3)到CFT付款 4)CFT付款成功
	     auction.paipai.com pay.paipai.com
	 拍拍下单到CFT付款成功(虚拟物品)关键路径1.1.5	1)下单页面 2)CFT付款成功
	     auction.paipai.com pay.paipai.com
	*/
	
	/*扩散路径
	 miniportal.paipai.com
	 www.paipai.com
	*/
	
	var reg = /^[1-9][0-9]{4,9}\.paipai\.com$/gi;  
	if(uri == "/cgi-bin/sell_type_choose")
	{
		pgvMain("pathtrace", {keyPathTag: "2008.3.1", nodeIndex: 1}); return;
	}
	else if(uri == "/cgi-bin/commodity_type_choose")
	{
		pgvMain("pathtrace", {keyPathTag: "2008.3.1", nodeIndex: 2});return;
	}
	else if(uri == "/cgi-bin/commodity_update_success")
	{
		pgvMain("pathtrace", {keyPathTag: "2008.3.1", nodeIndex: 3});return;
	}
	//搜索 search.paipai.com  search1.paipai.com list.paipai.com
	else if(domain=="search.paipai.com"||domain=="search1.paipai.com"||domain=="list.paipai.com")
	{
		pgvMain("pathtrace", {keyPathTag: "1.1.1|1.1.2|1.1.3", nodeIndex: "1|1|1", virtualURL: "/", virtualTitle: "搜索"});  return;
	}
	//商品详情 auction1.paipai.com
	else if(domain=="auction1.paipai.com")
	{
		pgvMain("pathtrace", {keyPathTag: "1.1.1|1.1.2|1.1.3", nodeIndex: "2|2|2", virtualURL: "/", virtualTitle: "商品详情"});  return;
	}
	//店铺详情 shop.paipai.com
	else if(domain=="shop.paipai.com" || (reg != null && reg.test(domain)))
	{
		pgvMain("pathtrace", {virtualDomain: "shop.paipai.com", virtualURL: "/", keyPathTag: "1.1.1", nodeIndex: 2, virtualTitle: "店铺详情"});return;
	}
	//下单 auction.paipai.com
	else if(uri == "/cgi-bin/commodity_fixup_confirm")
	{
		pgvMain("pathtrace", {keyPathTag: "1.1.1|1.1.2|1.1.3|1.1.4|1.1.5", nodeIndex: "3|3|4|1|1", virtualTitle: "下单"});return;
	}
	//下单成功 pay.paipai.com
	else if(uri == "/cgi-bin/pay_detail_confirm")
	{
		pgvMain("pathtrace", {keyPathTag: "1.1.1|1.1.4", nodeIndex: "4|2", virtualTitle: "下单成功"});return;
	}	
	//CFT付款成功 pay.paipai.com
	else if(uri == "/cgi-bin/pay_success_return" || uri == "/cgi-bin/mass_pay_return")
	{
		pgvMain("pathtrace", {virtualURL: "/cgi-bin/pay_return", keyPathTag: "1.1.1|1.1.4|1.1.5", nodeIndex: "6|4|2", virtualTitle: "付款成功"});return;
	}
	//登录 member.paipai.com
	else if(uri == "/cgi-bin/login_entry")
	{
		pgvMain("pathtrace", {keyPathTag: "1.1.3", nodeIndex: 3, virtualTitle: "登录"});return;
	}
	//miniportal miniportal.paipai.com
	else if(domain=="miniportal.paipai.com")
	{
		pgvMain("pathtrace", {pathStart: true, useRefUrl: true, override: true, careSameDomainRef: true});return;
	}	
	//首页 www.paipai.com
	else if(domain=="www.paipai.com" && uri=="/")
	{
		pgvMain("pathtrace", {pathStart: true, useRefUrl: true, override: true, careSameDomainRef: true, virtualTitle: "拍拍首页"});return;
	}
			
    pgvMain();
}

//s=(new Date()).getTime();
do_ping();
//e=(new Date()).getTime();
//alert('pgv:'+(e-s));
