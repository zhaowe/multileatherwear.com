
N = 70; 
Y = new Array();
X = new Array();
S = new Array();
A = new Array();
B = new Array();
M = new Array();
V = (document.layers)?1:0;

iH=(document.layers)?window.innerHeight:window.document.body.clientHeight;
iW=(document.layers)?window.innerWidth:window.document.body.clientWidth;
for (i=0; i < N; i++){                                                                
 Y[i]=Math.round(Math.random()*iH);
 X[i]=Math.round(Math.random()*iW);
 S[i]=Math.round(Math.random()*5+2);
 A[i]=0;
 B[i]=Math.random()*0.1+0.1;
 M[i]=Math.round(Math.random()*1+1);
}
if (V){
for (i = 0; i < N; i++)
{document.write("<LAYER NAME='sn"+i+"' LEFT=0 TOP=0 BGCOLOR='#FFFFF0' CLIP='0,0,"+M[i]+","+M[i]+"'></LAYER>")}
}
else{
document.write('<div style="position:absolute;top:0px;left:0px">');
document.write('<div style="position:relative">');
for (i = 0; i < N; i++)
{document.write('<div id="si" style="position:absolute;top:0;left:0;width:'+M[i]+';height:'+M[i]+';background:#fffff0;font-size:'+M[i]+'"></div>')}
document.write('</div></div>');
}
function snow(){
var H=(document.layers)?window.innerHeight:window.document.body.clientHeight;
var W=(document.layers)?window.innerWidth:window.document.body.clientWidth;
var T=(document.layers)?window.pageYOffset:document.body.scrollTop;
var L=(document.layers)?window.pageXOffset:document.body.scrollLeft;
for (i=0; i < N; i++){
sy=S[i]*Math.sin(90*Math.PI/180);
sx=S[i]*Math.cos(A[i]);
Y[i]+=sy;
X[i]+=sx; 
if (Y[i] > H){
Y[i]=-10;
X[i]=Math.round(Math.random()*W);
M[i]=Math.round(Math.random()*1+1);
S[i]=Math.round(Math.random()*5+2);
}
if (V){document.layers['sn'+i].left=X[i];document.layers['sn'+i].top=Y[i]+T}
else{si[i].style.pixelLeft=X[i];si[i].style.pixelTop=Y[i]+T} 
A[i]+=B[i];
}
setTimeout('snow()',50);
}
