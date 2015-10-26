var pvDoc=document;
if(window!=top)
{
    try
	{
	    pvDoc=top.document;
	}
	catch(e)
	{};
}

var cur_ref = pvDoc.referrer;
var req_url = pvDoc.URL; 

function adapt_url()
{
	if(null == req_url || "" == req_url)
	    return;	
    
    var reg = /^http:\/\/[1-9][0-9]{4,8}\.paipai\.com([\/?#].*)?$/gi;  
    
    if(reg == null || !reg.test(req_url))
        return;
        
    var uin = req_url.substr(7, req_url.indexOf(".") - 7);
    	
    var v_t1 = req_url.substr(req_url.indexOf(".") + 1);
    var v_uri_query = v_t1.substr(v_t1.indexOf(".") + 4);;
        
    req_url = "http://shop.paipai.com/" + uin + v_uri_query;
}

function adapt_ref()
{
	if(null == cur_ref || "" == cur_ref)
	    return;

    var reg = /^http:\/\/[1-9][0-9]{4,8}\.paipai\.com([\/?#].*)?$/gi;  
    
    if(reg == null || !reg.test(cur_ref))
        return;
        
    var uin = cur_ref.substr(7, cur_ref.indexOf(".") - 7);
    	
    var v_t1 = cur_ref.substr(cur_ref.indexOf(".") + 1);
    var v_uri_query = v_t1.substr(v_t1.indexOf(".") + 4);;
        
    cur_ref = "http://shop.paipai.com/" + uin + v_uri_query;	
	
}

function ReplaceUrl(url)
{
	len=url.length-1;
	while('/'==url.charAt(len))
	{
		len=len-1;
		url=url.substring(0,len+1);
	}
	return url;
}

function url_encode(url)
{
    return url.replace(/ /g,"%20").replace(/%/g,"%25").replace(/&/g,"%26").replace(/=/g,"%3D").replace(/\?/g,"%3F").replace(/\t/g,"%09").replace(/\r/g,"").replace(/\n/g,"").replace(/#/g,"%23");
}

function url_build()
{
	p=req_url.indexOf("?");
	if(-1==p)p=req_url.indexOf("&");
	if(-1==p)address="u="+req_url;
	else address="u="+req_url.substring(0,p);

    refer="r="+url_encode(cur_ref).substring(0,512);
	
	address=ReplaceUrl(address);
	refer=ReplaceUrl(refer);
	return refer+"&"+address + "&fu=" + url_encode(req_url);
}

function url_build1()
{
	p=req_url.indexOf("com");
	address="u="+req_url.substring(0,p+3);

    refer="r="+url_encode(cur_ref).substring(0,512);
	
	address=ReplaceUrl(address);
	refer=ReplaceUrl(refer);
  
	return refer+"&"+address + "&fu=" + url_encode(req_url);
}

var page_id=0,domain_id=0,link_id=0;
function get_rd_tag()
{
    url=req_url;
	p1=url.indexOf("RDTAG=");
	if(-1==p1)
	    p1=url.indexOf("ADTAG=");
	if(-1==p1)
	    return;
	p2=url.indexOf("&", p1 + 6);
	if(-1==p2)
		p2=url.indexOf("#", p1 + 6);

	rdtag="";
    if(-1==p2)
    {
        rdtag=url.substring(p1+6);
    }
	else
		rdtag=url.substring(p1+6, p2);

    ids=rdtag.split(".");
	if(3==ids.length)
	{
	   page_id=(ids[0] == null || ids[0] == "")?0:parseInt(ids[0]);    
	   domain_id=(ids[1] == null || ids[1] == "")?0:parseInt(ids[1]);
	   link_id=(ids[2] == null || ids[2] == "")?0:parseInt(ids[2]);
	}
	switch(page_id)
    {
		case 201:
			page_id = 601;
			break;
		case 202:
			page_id = 602;
			break;
		case 701:
			page_id = 380;
			break;
		default:
			break;
    }
}

function do_set_cookie(cookieName, cookieValue, expires, path, domain, secure)
{
  document.cookie =
    escape(cookieName) + '=' + escape(cookieValue)
    + (expires ? '; expires=' + expires.toGMTString() : '')
    + (path ? '; path=' + path : '')
    + (domain ? '; domain=' + domain : '')
    + (secure ? '; secure' : '');
}

function do_get_cookie(cookieName)
{ 
  var cookieValue = "";
  var search = cookieName + "=";
  if(document.cookie.length > 0)
  { 
    offset = document.cookie.indexOf(search);
    if (offset != -1)
    { 
      offset += search.length;
      end = document.cookie.indexOf(";", offset);
      if (end == -1) end = document.cookie.length;
      cookieValue = unescape(document.cookie.substring(offset, end))
    }
	else
	{
		return null;
	}
  }
  return cookieValue;
}
function set_market_rd(page_id,domain_id,link_id)
{
    if(page_id<=100||page_id>200||domain_id<=0||link_id<=0)
		return;
    rdtag=page_id+":"+domain_id+":"+link_id;

	today=new Date();
	//expire time is 1 day 86400000 = 3600*24*1000
    expiretime=new Date(today.getTime() + 86400000);
	
    do_set_cookie("MRD", rdtag, expiretime, "/", "paipai.com", 0);
	rdtag=page_id+"."+domain_id+"."+link_id;
	do_set_cookie("MRDT", rdtag, null, "/", "paipai.com", 0);

	crd="0:0:0";
	prd="0:0:0";
	login_src_id="0";
	if((cprd=do_get_cookie("CRD"))==null)cprd="";
	if(cprd != "")
	{
	    ids=cprd.split("-");
	    if(ids.length>=3 && ids[2] != "")
		    login_src_id=ids[2];
	}
	rdtag=crd+"-"+prd+"-"+login_src_id;	
	do_set_cookie("CRD", rdtag, expiretime, "/", "paipai.com", 0);
}

function set_channel_rd(page_id,domain_id,link_id)
{
    if(page_id<=200||page_id>500||domain_id<=0||link_id<=0)
		return;
    rdtag=page_id+":"+domain_id+":"+link_id;
    prd="0:0:0";
	login_src_id="0";
	if((cprd=do_get_cookie("CRD"))==null)cprd="";
	if(cprd != "")
	{
	    ids=cprd.split("-");
	    if(ids.length>=2 && ids[1] != "")
	        prd=ids[1];
	    if(ids.length>=3 && ids[2] != "")
		    login_src_id=ids[2];
	}
	rdtag+="-"+prd+"-"+login_src_id;	
	
	today=new Date();
    expiretime=new Date(today.getTime() + 86400000);
	
	do_set_cookie("CRD", rdtag, expiretime, "/", "paipai.com", 0);
}

function set_product_rd(page_id,domain_id,link_id)
{
    if(page_id<=500||domain_id<=0||link_id<=0)
		return;
    rdtag=page_id+":"+domain_id+":"+link_id;
	crd="0:0:0";
	login_src_id="0";
	if((cprd=do_get_cookie("CRD"))==null)cprd="";
	if(cprd != "")
	{
	    ids=cprd.split("-");
	    if(ids.length>=1 && ids[0] != "")
		    crd=ids[0];
	    if(ids.length>=3 && ids[2] != "")
		    login_src_id=ids[2];
	}
	rdtag=crd+"-"+rdtag+"-"+login_src_id;	

	today=new Date();
    expiretime=new Date(today.getTime() + 86400000);
	
	do_set_cookie("CRD", rdtag, expiretime, "/", "paipai.com", 0);
}

function set_rd_cookie()
{
    if(page_id<=100||domain_id<=0||link_id<=0)
		return;

    if(page_id>=101&&page_id<=200)
    {
        set_market_rd(page_id,domain_id,link_id);
    }
	else if(page_id>=201&&page_id<=500)
	{
	    set_channel_rd(page_id,domain_id,link_id);
	}
	else
	{
        set_product_rd(page_id,domain_id,link_id);
	}
}

function clear_login_src_id()
{
    if(req_url.indexOf("http://member.paipai.com/cgi-bin/login_entry") != -1)
		return;
	
	crd="0:0:0";
	prd="0:0:0";
	login_src_id="0";
	if((cprd=do_get_cookie("CRD"))==null)cprd="";
	if(cprd != "")
	{
	    ids=cprd.split("-");
	    if(ids.length>=1 && ids[0] != "")
		    crd=ids[0];
	    if(ids.length>=2 && ids[1] != "")
		    prd=ids[1];
	}
	rdtag=crd+"-"+prd+"-"+login_src_id;	
	today=new Date();
    expiretime=new Date(today.getTime() + 86400000);
	do_set_cookie("CRD", rdtag, expiretime, "/", "paipai.com", 0);
}

function set_login_src_id(id)
{
	crd="0:0:0";
	prd="0:0:0";
	login_src_id=id;
	if((cprd=do_get_cookie("CRD"))==null)cprd="";
	if(cprd != "")
	{
	    ids=cprd.split("-");
	    if(ids.length>=1 && ids[0] != "")
		    crd=ids[0];
	    if(ids.length>=2 && ids[1] != "")
		    prd=ids[1];
	}
	rdtag=crd+"-"+prd+"-"+login_src_id;	
	today=new Date();
	//expire time is 1 day 86400000 = 3600*24*1000
    expiretime=new Date(today.getTime() + 86400000);
	do_set_cookie("CRD", rdtag, expiretime, "/", "paipai.com", 0);
}

function set_ptag_cookie(prefix, ptag)
{
    if(null == prefix || "" == prefix || null == ptag || "" == ptag)
		return;

    pprd_p=do_get_cookie("PPRD_P");
	if(pprd_p==null || pprd_p == "null") 
	    pprd_p="";
    
	p1=pprd_p.indexOf(prefix);
	if(-1==p1)
	{
	    if(""==pprd_p)
			pprd_p=prefix+ptag;
		else
			pprd_p+="-"+prefix+ptag;
	}
	else
	{
	    p2=pprd_p.indexOf("-", p1);

        var left="",right="";
		left=pprd_p.substring(0,p1);
		if(-1!=p2)
			right=pprd_p.substring(p2);

		pprd_p=left+prefix+ptag+right;
	}

	today=new Date();
	//expire time is 1 day 86400000 = 3600*24*1000
    expiretime=new Date(today.getTime() + 86400000);
	do_set_cookie("PPRD_P", pprd_p, expiretime, "/", "paipai.com", 0);
}

var ptag_page_id = 0,ptag_domain_id = 0,ptag_link_id = 0;

function parse_ptag()
{
	if(req_url == null || req_url == "")
	    return;
  
  url=req_url;
  ptag="";
	p1=url.indexOf("PTAG=");
	if(-1==p1)
	{
		  p1=url.indexOf("PTAG,");
		  if(-1 == p1)
	        return;
	    
	    p3=url.indexOf(".htm", p1 + 5);    
	    p4=url.indexOf("-", p1 + 5);
	    if(-1==p3)
	        return;
	    else if(-1==p4)
	        ptag=url.substring(p1+5, p3);
	    else
	        ptag=url.substring(p1+5, p4);    
	        
	    req_url=req_url+"#PTAG="+ptag;
	}
	else
  {
   		p2=url.indexOf("&", p1 + 5);
	    if(-1==p2)
		      p2=url.indexOf("#", p1 + 5);
      if(-1==p2)
          ptag=url.substring(p1+5);
	    else
		    ptag=url.substring(p1+5, p2);
  }


    pids=ptag.split(".");

	if(3==pids.length)
	{
	   ptag_page_id = (pids[0] == null || pids[0] == "")?0:parseInt(pids[0]);
	   ptag_domain_id = (pids[1] == null || pids[1] == "")?0:parseInt(pids[1]);          
	   ptag_link_id = (pids[2] == null || pids[2] == "")?0:parseInt(pids[2]);          
	}
	else
		return;
	
	if(0 == ptag_page_id || 0 == ptag_domain_id || 0 == ptag_link_id)
		return;

	switch(ptag_page_id)
    {
		case 201:
			ptag_page_id = 30018;
			break;
		case 202:
			ptag_page_id = 30019;
			break;
		case 603:
			ptag_page_id = 30020;
			break;
		case 701:
			ptag_page_id = 30025;
			break;
		default:
			break;
    }
		
	if(ptag_page_id >= 10001 && ptag_page_id <= 20000)
	{
	    set_ptag_cookie("EA.", ptag_page_id+"."+ptag_domain_id+"."+ptag_link_id);
	}
	else if(ptag_page_id >= 20001 && ptag_page_id <= 30000)
	{
	    set_ptag_cookie("IA.", ptag_page_id+"."+ptag_domain_id+"."+ptag_link_id);
	}
	else if(ptag_page_id >= 30001 && ptag_page_id <= 40000)
	{
	    set_ptag_cookie("CT.", ptag_page_id+"."+ptag_domain_id+"."+ptag_link_id);
	}
	else if(ptag_page_id >= 40001 && ptag_page_id <= 50000)
	{
	    set_ptag_cookie("PD.", ptag_page_id+"."+ptag_domain_id+"."+ptag_link_id);
	}
}

function set_pprds_cookie(prefix, ptag)
{
    if(null == prefix || "" == prefix || null == ptag)
		return;
    
    var pprd_s = do_get_cookie("PPRD_S");
	if(pprd_s == null || pprd_s == "null") 
	    pprd_s="";
	    
	p1=pprd_s.indexOf(prefix);
	if(-1==p1)
	{
		if("" != ptag)
		{
	        if(""==pprd_s)
			    pprd_s=prefix+ptag;
		    else
			    pprd_s+="-"+prefix+ptag;
	    }	    	
	}
	else
	{
	    p2=pprd_s.indexOf("-", p1);

        var left="",right="";
		left=pprd_s.substring(0,p1);
		if(-1!=p2)
			right=pprd_s.substring(p2);

        if("" != ptag)
		    pprd_s=left+prefix+ptag+right;
		else
			pprd_s=left+right;
	}

    if("" != pprd_s)
	    do_set_cookie("PPRD_S", pprd_s, null, "/", "paipai.com", 0);
	else
	{
		today=new Date();
        expiretime=new Date(today.getTime() - 86400000);
		do_set_cookie("PPRD_S", pprd_s, expiretime, "/", "paipai.com", 0);
	}
}

function parse_exact_ptag()
{
	if(window.location.hostname != "auction1.paipai.com")
	    return;
	
	if(ptag_page_id <= 0 || ptag_domain_id <= 0 || ptag_link_id <= 0)
	{
		set_pprds_cookie("TJWC.", "");
		return;
	}
	
	if(ptag_page_id == 40006)
	{
		if(ptag_domain_id == 1)
		{
			if(ptag_link_id >= 1 && ptag_link_id <= 3)
			    return;
		}
		else if(ptag_domain_id == 2)
		{
			if(ptag_link_id >= 2 && ptag_link_id <= 4)
			    return;
		}
	}
	
	set_pprds_cookie("TJWC.", ptag_page_id+"."+ptag_domain_id+"."+ptag_link_id);
	return;
}

var se_page_id = 0,se_domain_id = 0,se_link_id = 0;
function parse_se_refer()
{  
	if(cur_ref == null)
	    return;
	var bmatch = false;
	
	var reg = /^http:\/\/([a-zA-Z0-9_]+\.)*paipai(\.[a-zA-Z0-9_]+){1,2}\/.*$/gi;      
	if(reg != null && reg.test(cur_ref))
    {
    	return;
    } 
    	
	reg = /^http:\/\/([a-zA-Z0-9_]+\.)*baidu(\.[a-zA-Z0-9_]+){1,2}\/.*$/gi;      
	if(!bmatch && reg != null && reg.test(cur_ref))
    {
    	se_page_id = 40007, se_domain_id = 1, se_link_id = 1;
    	bmatch = true;
    }  
    
	reg = /^http:\/\/([a-zA-Z0-9_]+\.)*google(\.[a-zA-Z0-9_]+){1,2}\/.*$/gi;     
	if(!bmatch && reg != null && reg.test(cur_ref))
    {
    	se_page_id = 40007, se_domain_id = 1, se_link_id = 2;	
    	bmatch = true;
    }  
    
	reg = /^http:\/\/([a-zA-Z0-9_]+\.)*soso\.com\/.*$/gi;      
	if(!bmatch && reg != null && reg.test(cur_ref))
    {
    	se_page_id = 40007, se_domain_id = 1, se_link_id = 3;
    	bmatch = true;
    }   

	reg = /^http:\/\/([a-zA-Z0-9_]+\.)*3721\.com\/.*$/gi;     
	if(!bmatch && reg != null && reg.test(cur_ref))
    {
    	se_page_id = 40007, se_domain_id = 1, se_link_id = 4;
    	bmatch = true;
    }   
    
	reg = /^http:\/\/([a-zA-Z0-9_]+\.)*sogou\.com\/.*$/gi;      
	if(!bmatch && reg != null && reg.test(cur_ref))
    {
    	se_page_id = 40007, se_domain_id = 1, se_link_id = 5;
    	bmatch = true;
    }

	reg = /^http:\/\/search\.msn\.com\/.*$/gi;      
	if(!bmatch && reg != null && reg.test(cur_ref))
    {
    	se_page_id = 40007, se_domain_id = 1, se_link_id = 6;
    	bmatch = true;
    }      

	reg = /^http:\/\/([a-zA-Z0-9_]+\.)*yisou\.com\/.*$/gi;      
	if(!bmatch && reg != null && reg.test(cur_ref))
    {
    	se_page_id = 40007, se_domain_id = 1, se_link_id = 7;
    	bmatch = true;
    }

 	reg = /^http:\/\/([a-zA-Z0-9_]+\.)*yahoo(\.[a-zA-Z0-9_]+){1,2}\/.*$/gi;     
	if(!bmatch && reg != null && reg.test(cur_ref))
    {
    	se_page_id = 40007, se_domain_id = 1, se_link_id = 8;	
    	bmatch = true;
    }  

	reg = /^http:\/\/([a-zA-Z0-9_]+\.)*live\.com\/.*$/gi;      
	if(!bmatch && reg != null && reg.test(cur_ref))
    {
    	se_page_id = 40007, se_domain_id = 1, se_link_id = 9;
    	bmatch = true;
    }

	reg = /^http:\/\/([a-zA-Z0-9_]+\.)*yodao\.com\/.*$/gi;      
	if(!bmatch && reg != null && reg.test(cur_ref))
    {
    	se_page_id = 40007, se_domain_id = 1, se_link_id = 10;
    	bmatch = true;
    }

	reg = /^http:\/\/([a-zA-Z0-9_]+\.)*iask\.com\/.*$/gi;      
	if(!bmatch && reg != null && reg.test(cur_ref))
    {
    	se_page_id = 40007, se_domain_id = 1, se_link_id = 11;
    	bmatch = true;
    }

	reg = /^http:\/\/([a-zA-Z0-9_]+\.)*zhongsou\.com\/.*$/gi;      
	if(!bmatch && reg != null && reg.test(cur_ref))
    {
    	se_page_id = 40007, se_domain_id = 1, se_link_id = 12;
    	bmatch = true;
    }

	reg = /^http:\/\/([a-zA-Z0-9_]+\.)*qikoo\.com\/.*$/gi;      
	if(!bmatch && reg != null && reg.test(cur_ref))
    {
    	se_page_id = 40007, se_domain_id = 1, se_link_id = 13;
    	bmatch = true;
    }

	reg = /^http:\/\/([a-zA-Z0-9_]+\.)*aol\.com\/.*$/gi;      
	if(!bmatch && reg != null && reg.test(cur_ref))
    {
    	se_page_id = 40007, se_domain_id = 1, se_link_id = 14;
    	bmatch = true;
    }

	reg = /^http:\/\/([a-zA-Z0-9_]+\.)*naver\.com\/.*$/gi;      
	if(!bmatch && reg != null && reg.test(cur_ref))
    {
    	se_page_id = 40007, se_domain_id = 1, se_link_id = 15;
    	bmatch = true;
    }

	reg = /^http:\/\/([a-zA-Z0-9_]+\.)*yandex\.com\/.*$/gi;      
	if(!bmatch && reg != null && reg.test(cur_ref))
    {
    	se_page_id = 40007, se_domain_id = 1, se_link_id = 16;
    	bmatch = true;
    }

	reg = /^http:\/\/([a-zA-Z0-9_]+\.)*ask\.com\/.*$/gi;      
	if(!bmatch && reg != null && reg.test(cur_ref))
    {
    	se_page_id = 40007, se_domain_id = 1, se_link_id = 17;
    	bmatch = true;
    }
    
 	reg = /^http:\/\/([a-zA-Z0-9_]+\.)*vnet(\.[a-zA-Z0-9_]+){1,2}\/.*$/gi;     
	if(!bmatch && reg != null && reg.test(cur_ref))
    {
    	se_page_id = 40007, se_domain_id = 1, se_link_id = 18;	
    	bmatch = true;
    }  
    
    //如果refer匹配搜索引擎, 且PTAG非产品段, 则设置Cookie
    if(bmatch && !(ptag_page_id >= 40001 && ptag_page_id <= 50000))
        set_ptag_cookie("PD.", se_page_id+"."+se_domain_id+"."+se_link_id);
}

function parse_pv_src()
{    
	//1. 客户端来源老RD 114
	if(114 == page_id && domain_id > 0 && link_id > 0)
	{
		set_pprds_cookie("PVS.", "OR." + page_id + "." + domain_id + "." + link_id);
		set_pprds_cookie("PVSE.", "1");
		return; 
	}
	
	//2. 客户端来源老RD 199
	if(199 == page_id && domain_id > 0 && link_id > 0)
	{
	    set_pprds_cookie("PVS.", "OR." + page_id + "." + domain_id + "." + link_id);	
	    set_pprds_cookie("PVSE.", "1");	
		return; 
	}
	
	//3. 客户端来源新RD 10012 10008
	if(ptag_page_id > 0 && ptag_domain_id > 0 && ptag_link_id > 0)
	{
	    if(10008 == ptag_page_id || 10005 == ptag_page_id || 10006 == ptag_page_id || 10007 == ptag_page_id)
	    {
		    set_pprds_cookie("PVS.", "NR." + ptag_page_id + "." + ptag_domain_id + "." + ptag_link_id);
		    set_pprds_cookie("PVSE.", "1");
		    return; 
	    }
	    else if(10001 == ptag_page_id && (1 == ptag_domain_id || 2 == ptag_domain_id || 3 == ptag_domain_id))
	    {
		    set_pprds_cookie("PVS.", "NR." + ptag_page_id + "." + ptag_domain_id + "." + ptag_link_id);
		    set_pprds_cookie("PVSE.", "1");
		    return; 
	    }
	    else if(10012 == ptag_page_id && (8 == ptag_domain_id || 10 == ptag_domain_id))
	    {
		    set_pprds_cookie("PVS.", "NR." + ptag_page_id + "." + ptag_domain_id + "." + ptag_link_id);
		    set_pprds_cookie("PVSE.", "1");
		    return; 
	    }
	    else if(30053 == ptag_page_id && (9 == ptag_domain_id || 15 == ptag_domain_id))
	    {
		    set_pprds_cookie("PVS.", "NR." + ptag_page_id + "." + ptag_domain_id + "." + ptag_link_id);
		    set_pprds_cookie("PVSE.", "1");
		    return; 
	    }
    }
    
    var reg1 = /^http:\/\/([a-zA-Z0-9_]+\.)*paipai\.com\/.*$/gi; 
	var reg2 = /^http:\/\/miniportal.paipai\.com\/.*$/gi; 
	
	if(reg2.test(cur_ref))
	{
	    set_pprds_cookie("PVS.", "MINIPORTAL");
	    set_pprds_cookie("PVSE.", "1");
	    return;
	}
	if(reg1.test(cur_ref))
	{
		// 如果来源标识已经存在,则不进行设置,否则置为用户主动访问
		var pprd_s = do_get_cookie("PPRD_S");
	    if(pprd_s != null && pprd_s.indexOf("PVS.") >= 0) 
	    {
	    	set_pprds_cookie("PVSE.", "0");
	        return;
	    }
	    else
	    {
	    	set_pprds_cookie("PVS.", "PAIPAI");
		    set_pprds_cookie("PVSE.", "1");
	        return;
	    }
	}
	
	//4. 如果来源为搜索引擎
	if(40007 == se_page_id)
	{
		set_pprds_cookie("PVS.", "SE." + se_page_id + "." + se_domain_id + "." + se_link_id);
		set_pprds_cookie("PVSE.", "1");
		return; 
	} 
		
	//5. 来源为空
    if(null == cur_ref || "" == cur_ref)
	{
		// 如果来自自动登录跳转(ft=1标记),
		if(req_url != null && (req_url.indexOf("&ft=1&") >= 0 || req_url.indexOf("?ft=1&") >= 0))
		{
		    set_pprds_cookie("PVS.", "AUTOLOGIN");
		    set_pprds_cookie("PVSE.", "1");
		    return;			
		}
		
		// 如果来源标识已经存在,则不进行设置,否则置为用户主动访问
		var pprd_s = do_get_cookie("PPRD_S");
	    if(pprd_s != null && pprd_s.indexOf("PVS.") >= 0) 
	    {
	    	set_pprds_cookie("PVSE.", "0");
	        return;
	    }
	    
		set_pprds_cookie("PVS.", "USER");
		set_pprds_cookie("PVSE.", "1");
		return;
	}
	
	//6. 截取refer中的域名部分
	var p1 = cur_ref.indexOf("://");
	if(p1 < 0)
	    return;
	
	p1+=3;
	var domainname="";
	len = cur_ref.length;
	for(i=p1;i<len;i++)
	{
		c=cur_ref.charAt(i);
		if('.'==c
		  || (c>='0'&&c<='9')
		  || (c>='a'&&c<='z') 
		  || (c>='A'&&c<='Z'))
		  domainname+=c;
		else
		    break; 
	}
	domainname=domainname.toLowerCase();
	set_pprds_cookie(domainname);
	if(""==domainname)
	    return;
	    
	//7. 来源为*.qq.com
	var reg3 = /^([a-zA-Z0-9_]+\.)+qq\.com/gi; 
	if(reg3.test(domainname))
		set_pprds_cookie("PVS.", "QW."+domainname);
	//8. 来源为其他站点 
	else
		set_pprds_cookie("PVS.", "OW."+domainname);
		
	set_pprds_cookie("PVSE.", "1");
}

function parse_sem_src()
{
	// 1. 非来自baidu的跳转不作处理
	if(40007 != se_page_id
	  || 1 != se_domain_id
	  || 1 != se_link_id)
	    return;
    // 2. 查找关键字参数
    if(null == cur_ref || "" == cur_ref)
        return;    
    p1 = cur_ref.indexOf("wd=");

    if(-1 == p1)
    {
    	p1 = cur_ref.indexOf("word=");
    	if(-1 == p1)
    	    return;
    	else
    	    p1 += 5;
    }
    else
        p1 += 3;
    
    p2 = p1;
    for(; p2 < cur_ref.length; p2++)
    {
    	c = cur_ref.charAt(p2);
        if('&' == c || '?' == c || '=' == c || '#' == c)
            break;
    }
    
    var kw = cur_ref.substr(p1, p2 - p1);
    if(null == kw || "" == kw)
        return;
    // 3. 清理关键字
    kw=kw.toUpperCase().replace(/\+/g,"%20").replace(/ /g,"%20").replace(/\t/g,"%09").replace(/\r/g,"").replace(/\n/g,"").replace(/\./g,"%2E").replace(/-/g,"%2D").replace(/;/g,"%3B").replace(/,/g,"%2C");
    // 4. 设置Cookie
    set_ptag_cookie("SEM.", kw);
}

function ping_build(type,id)
{
	adapt_url();
	adapt_ref();
    get_rd_tag();
    set_rd_cookie();
    parse_ptag();
    parse_exact_ptag();
	parse_se_refer();
	parse_pv_src();
	parse_sem_src();
	
	if(!id)
	id=0;
	if(0==type)
	var url=url_build();
	if(1==type)
	var url=url_build1();
	var resolution="resolution="+window.screen.width+"*"+window.screen.height;
	var color="color="+window.screen.colorDepth;
	var  cookieEnabled=(navigator.cookieEnabled)?true:false;
	if (typeof navigator.cookieEnabled=="undefined" && !cookieEnabled)
	{
		document.cookie="testcookie";
		cookieEnabled=(document.cookie.indexOf("testcookie")!=-1)?true:false;
		document.cookie="";
	}
	if (cookieEnabled)
		cookiesup="cookiesup=1";
	else
		cookiesup="cookiesup=0";
	if(0==id)
	var data = url+"&"+resolution+"&"+color+"&"+cookiesup;
	else
	var data = url+"/"+id+"&"+resolution+"&"+color+"&"+cookiesup;
	data="pageId="+page_id+"&domainId="+domain_id+"&linkId="+link_id+"&"+data;
	data="fpageId="+ptag_page_id+"&fdomainId="+ptag_domain_id+"&flinkId="+ptag_link_id+"&"+data;
	data="sepageid="+se_page_id+"&sedomainid="+se_domain_id+"&selinkid="+se_link_id+"&"+data;
	document.write("<div style='position:absolute; left:0px; top:0px; width:0px; height:0px; z-index:0; visibility: hidden;'><!-- < img src='http://service.paipai.com/cgi-bin/ping?rand=" + Math.round(Math.random()*100000) + "&" + data + "' height='0' width='0' > -->WL</div>");
}

