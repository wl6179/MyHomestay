
var tempObj=document.getElementById("divCN-EN");
hutia(tempObj);

function hutia(obj){
 var re="", nn, ss;
 if(obj.nodeType==1){
  for(var i=0;i<obj.childNodes.length;i++)hutia(obj.childNodes[i]);
 }else if(obj.nodeType==3){
  nn=document.createElement("span");
  nn.className="cn";
  obj.parentNode.insertBefore(nn,obj);
  ss=obj.nodeValue.replace(/&/g,"&amp;").replace(/ /g,"&nbsp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/\n/g,"<br>").replace(/\"/g,"&quot;");
  obj.parentNode.removeChild(obj);
  nn.innerHTML=ss.replace(/[^\u0391-\uFFE5]+/g,hutia_en);
 }
}

function hutia_en(str){
  return("<span class=\"en\">"+str+"</span>");
}
