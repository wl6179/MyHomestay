function nextSale(order){
	if(order=="up")	saleNum++;
	else saleNum--;
	if(saleNum>2)	saleNum=0
	else if(saleNum<0) saleNum=2;
	for(i=0;i<2;i++)
		document.getElementById("saleList"+i).style.display="none";
	document.getElementById("saleList"+saleNum).style.display="";
}
function showLink(name,order){
	for(i=0;i<2;i++){
		document.getElementById(name+"Push"+i).style.display="none";
		document.getElementById(name+"Link"+i).className=""
	}
	document.getElementById(name+"Push"+order).style.display="";
	document.getElementById(name+"Link"+order).className=name+"LinkHere";
}
function cutSaleName(str){
	var len = 0;
	var strLen=0;
	var strReal = 0;
	var lenReal=str.replace(/[^\x00-\xff]/g,"**").length;
	if(lenReal<=26) return str+"<br />";
	else if(lenReal<=40)  return str;
	while(strReal<=40){
		if(str.charCodeAt(strLen)>255)
			strReal += 2;
		else
			strReal += 1;
		if(strReal<=40) strLen++;
	}
	return str.substr(0,strLen)+"бн";
}