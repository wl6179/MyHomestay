var show_word="������Ʒ�ؼ���";
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
		if(obj.value==show_word||obj.value=="����ؼ��ֽ�������")
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

var arrFirstSort=[[3119,"�ֻ���ֵ���� - �绰��"],[12001,"������Ϸ������Ʒ"],[24590,"�Ű���ר��"],1,[20501,"Ůװ - Ůʿ��Ʒ"],[27158,"Ůʿ���� - ���� - Ӿװ"],[20001,"��ױ - ��ˮ - ���� - ����"],[2001,"�鱦���� - ʱ����Ʒ"],[21001,"Ůʿ��� - ���"],[21036,"ŮЬ"],[28001,"Ʒ���ֱ� - �����ֱ�"],[21501,"��ʿ���� - Zippo - ��ʿ��Ʒ"],[22001,"��װ - �а� - �������"],[6001,"�˶� - ���� - ����"],[6070,"���� - ��Ʒ - ���� - ��Ʊ"],[22501,"Ӥ�� - �и���Ʒ - ͯװ"],0,[3002,"�ֻ�"],[29551,"�ֻ����� - ������"],[28038,"�ʼǱ�����"],[1,"����Ӳ�� - ̨ʽ����"],[5001,"�������� - ���� - ����"],[4001,"������� - ����� - ��ӡ"],[28039,"������� - �������"],[28009,"������� - ����Ԫ��"],[28046,"�칫�豸 - �ľ� - �Ĳ�"],[11001,"�����ܱ� - ��Ϸ�ܱ�"],[23001,"��ͨ��Ʒ - ��� - ģ��"],1,[28055,"�Ӽ� - ���� - ��Ʒ - ����"],[28053,"�ҵ� - �������� - ������е"],[28052,"ʳƷ - ����Ʒ - ��Ҷ - �ز�"],[28054,"�Ҿ� - �˼Ҵ���"],[28056,"װ�� - �ƾ� - ��� - ��ԡ"],[8001,"�ʱ� - �Ŷ� - �ֻ� - �ղ�"],[24001,"���� - ����Ʒ - ��Ʒ - �ʻ�"],[23501,"�鼮 - ��ֽ - ��־"],[24501,"��Ӱ - ���� - ���� - ����"],[9001,"���� - Ħ�� - ���г�"],[21591,"���� - ���� - ��Ȥ����"],[10001,"סլ - ���� - �칫¥����"]];

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