// JavaScript Document
function change(obj)
{
for(i=1;i<6;i++)
{
if(i==obj)
{
eval("document.getElementById('week"+i+"').className='titlef1'");
eval("document.getElementById('link"+i+"').className='nav'");
eval("document.getElementById('tab"+i+"').style.display='block'");
<!--document.forms[i-1].oftype.value=obj;--> 
 

}
else
{
eval("document.getElementById('week"+i+"').className='titlef2'");
eval("document.getElementById('link"+i+"').className='ftitle'");
eval("document.getElementById('tab"+i+"').style.display='none'");
}
}
}

function showdate()
{
for(i=1;i<6;i++)
{
eval("document.getElementById('week"+i+"').className='titlef2'");
eval("document.getElementById('link"+i+"').className='ftitle'");
eval("document.getElementById('tab"+i+"').style.display='none'");

}
var d = new Date();
var num = d.getDay();
if (num == 0)
{
num = 5;
}
eval("document.getElementById('week"+num+"').className='titlef1'");
eval("document.getElementById('link"+num+"').className='nav'");
eval("document.getElementById('tab"+num+"').style.display='block'");
 }