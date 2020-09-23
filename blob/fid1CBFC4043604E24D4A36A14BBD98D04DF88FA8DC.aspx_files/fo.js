//  Copyright (c) 2000-2007 ZEDO Inc. All Rights Reserved.
function B2(){
var t2=navigator.userAgent.toLowerCase();var y6=(t2.indexOf('mac')!=-1);var o6=(!y6&&(t2.indexOf('msie 5')!=-1)||(t2.indexOf('msie 6')!=-1));
if(o6){
document.writeln('<scr'+'ipt language=VBS'+'cript>');
document.writeln('on error resume next');
document.writeln('r0=IsObject(CreateObject("ShockwaveFlash.ShockwaveFlash.5"))');
document.writeln('if(r0<=0)then r0=IsObject(CreateObject("ShockwaveFlash.ShockwaveFlash.4"))');
document.writeln('</scr'+'ipt>');
}
else if(navigator.mimeTypes&&
navigator.mimeTypes["application/x-shockwave-flash"]&&
navigator.mimeTypes["application/x-shockwave-flash"].enabledPlugin){
if(navigator.plugins&&navigator.plugins["Shockwave Flash"]){
var z4=navigator.plugins["Shockwave Flash"].description;
if(parseInt(z4.substring(z4.indexOf(".")-1))>=4){
r0=1;
}}}
var i4=navigator.javaEnabled();
a0=1;
if(i4){a0 |=4;}
if(r0){a0 |=8;}
if(o6){
if(document.body){
document.body.style.behavior='url(#default#clientCaps)';
if(document.body.connectionType=='lan'){
a0 |=16;
}}}
return a0;
}
var c0=0;var v0=0;var o0='0';var z0=0;var v4='';var r0=0;var p5='';var e3='';var t4='';var d4="";
if(typeof zflag_nid!='undefined'){
c0=zflag_nid;
zflag_nid=0;
}
if(typeof zflag_sid!='undefined'){
v0=zflag_sid;
zflag_sid=0;
}
if(typeof zflag_cid!='undefined'){
o0=zflag_cid;
zflag_cid=0;
}
if(typeof zflag_sz!='undefined'){
z0=zflag_sz;
if(z0<0||z0>95){
z0=0;
}
zflag_sz=0;
}
if(typeof zflag_kw!='undefined'){
zflag_kw=zflag_kw.replace(/&/g,'zzazz');
v4=escape(zflag_kw);
zflag_kw="";
}
if(typeof zflag_geo!='undefined'){
if(!isNaN(zflag_geo)){
e3="&g="+zflag_geo;
zflag_geo=0;
}}
if(typeof zflag_param!='undefined'){
d4="&p="+zflag_param;
zflag_param="";
}
if(typeof zflag_click!='undefined'){
zzTrd=escape(zflag_click);
t4='&l='+zzTrd;
zflag_click="";
}
var zzStr='';var zzCountry=255;var zzMetro=0;var zzState=0;var zzSection=v0;var zzD=window.document;var zzRand=(Math.floor(Math.random()* 1000000)% 10000);var zzCustom='';var zzPat='';var zzSkip='';
var zzExp='';var zzTrd='';var zzDm1=0;var zzDm2=0;var zzDm3=0;var zzDm4=0;var zzDm5=0;var zzDm6=0;var zzDm7=0;var zzDm8=0;var zzDm9=0;var zzDm10=0;var zzAGrp=0;var zzAct=new Array();
var zzActVal=new Array();
p5=B2();
if(p5<0||p5>31){
p5=1;
}
n0='<scr'+'ipt language="JavaScript" src="http://c7.zedo.com/bar/v14-000/c5/jsc/fm.js?c='+o0+'&n='+c0+'&r='+p5+'&d='+z0+'&q='+v4+'&s='+v0+e3+d4+t4+'&z='+Math.random()+'"></scr'+'ipt>';
document.write(n0);
