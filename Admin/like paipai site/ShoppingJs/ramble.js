function ramble()
{
	try{
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

		var url  =pvDoc.URL;
		var title=pvDoc.title;

		var start = url.indexOf("://");
		if (start == -1 ) return;
		start = start + 3;
		var end   = url.indexOf(".", start);
		if (end == -1 ) return;
		var directory = url.slice(start, end);
		var classid_obj = null;
		
		if (!isNaN(parseInt(directory)))    // 12345.paipai.com
		{
			classid_obj = pvDoc.getElementById("scId");
			url= "http://shop.paipai.com/"+directory;
			directory = "shop";
		}
		if (directory == "auction1" || directory == "auction")
		{
			classid_obj = pvDoc.getElementById("dwLeafClassId");  // 商品详情 <input type="hidden" name="dwLeafClassId" value="<%$dwLeafClassId$%>" />
		}
		else if (directory == "shop1" || directory == "shop")
		{
			classid_obj = pvDoc.getElementById("shopTypeId");  // 店铺详情 <input name="scId" id="scId" type="hidden" value="<%$sCategoryID$%>">
		}
		else if (directory == "search1" || directory =="search" || directory =="list" || directory == "sse")
		{
			classid_obj = pvDoc.getElementById("sClassid");  // search list页面		
		}

		var classid = 0;	
		if (classid_obj != null && classid_obj.value != null)
		{
			classid = parseInt(classid_obj.value);
			if (isNaN(classid)) classid = 0;
		}
		
		var data = "classid=" + classid + "&url=" + escape(url) + "&title=" + title;
		document.write("<div style='position:absolute; left:0px; top:0px; width:0px; height:0px; z-index:0; visibility: hidden;'><!--<img src='http://service.paipai.com/cgi-bin/ramble?" + data + "' height='0' width='0'>-->WL</div>");
		//alert(data);
	}catch(e){ };
}
ramble();